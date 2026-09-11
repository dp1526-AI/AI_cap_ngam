import time, os, re, threading, json
#import tkinter as tk
#from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
from docx.shared import Pt
from google import genai
from google.genai import types

# Cấu hình API Key (Hỗ trợ chuẩn 100% các khóa bắt đầu bằng AQ. hoặc AIza...)
API_KEY = "Nhập mã API"
KNOWLEDGE_FILE = "knowledge_history.json"

class CableAIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Phân Tích & Thẩm Định Biên Bản Cáp Ngầm EVNHCMC - v5.1 (GenAI SDK)")
        self.root.geometry("850x720")
        
        self.path_qt = tk.StringVar()
        self.path_bb = tk.StringVar()
        
        # Khởi tạo Client từ SDK mới
        self.client = genai.Client(api_key=API_KEY)
        
        self.create_widgets()
        self.init_knowledge_base()

    def init_knowledge_base(self):
        if not os.path.exists(KNOWLEDGE_FILE):
            with open(KNOWLEDGE_FILE, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        # --- Phần chọn File ---
        frame_files = tk.LabelFrame(self.root, text="Thiết lập tài liệu đầu vào (Tự động thẩm định)", padx=10, pady=10)
        frame_files.pack(fill="x", padx=20, pady=10)

        tk.Button(frame_files, text="1. Chọn Quy trình QT-CT-02", command=self.btn_select_qt, width=25, bg="#E1E1E1").grid(row=0, column=0, pady=5)
        tk.Entry(frame_files, textvariable=self.path_qt, width=55).grid(row=0, column=1, padx=10)

        tk.Button(frame_files, text="2. Chọn Biên bản CBM (PD)", command=self.btn_select_bb, width=25, bg="#E1E1E1").grid(row=1, column=0, pady=5)
        tk.Entry(frame_files, textvariable=self.path_bb, width=55).grid(row=1, column=1, padx=10)

        # --- Phần Nhật ký & Kết quả ---
        frame_log = tk.LabelFrame(self.root, text="Nhật ký kiểm tra & Kết quả phân tích chuyên gia", padx=10, pady=10)
        frame_log.pack(fill="both", expand=True, padx=20, pady=5)

        self.log_area = scrolledtext.ScrolledText(frame_log, height=22, font=("Consolas", 10))
        self.log_area.pack(fill="both", expand=True)

        # --- Nút chức năng ---
        frame_btns = tk.Frame(self.root, pady=10)
        frame_btns.pack()

        self.run_btn = tk.Button(frame_btns, text="🚀 BẮT ĐẦU THẨM ĐỊNH & PHÂN TÍCH", bg="#0078D7", fg="white", 
                                 font=("Arial", 12, "bold"), padx=20, pady=10, command=self.start_thread)
        self.run_btn.pack()

    def write_log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def upload_to_gemini(self, path):
        self.write_log(f"⏳ Đang tải file lên Gemini: {os.path.basename(path)}...")
        file = self.client.files.upload(file=path)
        while file.state.name == "PROCESSING":
            time.sleep(1.5)
            file = self.client.files.get(name=file.name)
        return file

    def btn_select_qt(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.verify_qt_file_immediately, args=(path,)).start()

    def verify_qt_file_immediately(self, path):
        self.write_log("\n🔍 Đang kiểm tra nhanh file Quy trình vừa chọn...")
        try:
            pdf_qt = self.upload_to_gemini(path)
            prompt = """
            Kiểm tra tệp PDF này có phải là quy trình thử nghiệm cáp ngầm 'QT-CT-02' của EVNHCMC hay không.
            Trả về duy nhất định dạng JSON:
            {"is_qt_ct_02": true/false, "reason": "Lý do ngắn gọn"}
            """
            res = self.client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[pdf_qt, prompt],
                config=types.GenerateContentConfig(
                    system_instruction="Bạn là trợ lý kiểm tra tài liệu kỹ thuật.",
                    response_mime_type="application/json"
                )
            )
            
            match_json = re.search(r'\{.*\}', res.text, re.DOTALL)
            if match_json:
                check_data = json.loads(match_json.group(0))
                if check_data.get("is_qt_ct_02"):
                    self.path_qt.set(path)
                    self.write_log("✅ XÁC NHẬN: File Quy trình hợp lệ (Chuẩn QT-CT-02).")
                    messagebox.showinfo("Thành công", "File Quy trình hợp lệ (Chuẩn QT-CT-02)!")
                else:
                    self.path_qt.set("")
                    self.write_log("❌ LỖI: File không phải quy trình QT-CT-02!")
                    self.write_log(f"📌 Lý do: {check_data.get('reason')}")
                    messagebox.showerror("Sai Quy Trình", f"File đã chọn KHÔNG PHẢI là quy trình QT-CT-02!\nLý do: {check_data.get('reason')}\nVui lòng chọn lại đúng file!")
            else:
                self.path_qt.set(path)
        except Exception as e:
            self.write_log(f"⚠️ Lỗi kiểm tra file Quy trình: {e}")

    def btn_select_bb(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.path_bb.set(path)

    def start_thread(self):
        if not self.path_qt.get():
            messagebox.showwarning("Thiếu file", "Vui lòng chọn đúng File Quy trình QT-CT-02!")
            return
        if not self.path_bb.get():
            messagebox.showwarning("Thiếu file", "Vui lòng chọn File Biên bản thử nghiệm CBM!")
            return
        
        thread = threading.Thread(target=self.process_analysis)
        thread.start()

    def save_to_knowledge_base(self, record_info):
        try:
            with open(KNOWLEDGE_FILE, "r+", encoding="utf-8") as f:
                history = json.load(f)
                history.append(record_info)
                f.seek(0)
                json.dump(history[-20:], f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.write_log(f"⚠️ Không thể lưu lịch sử học: {e}")

    def fill_report(self, data_dict):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            template_file = os.path.join(current_dir, "Bao_Cao_mau.docx")
            
            if not os.path.exists(template_file):
                self.write_log("⚠️ CẢNH BÁO: Không tìm thấy 'Bao_Cao_mau.docx'. Không thể xuất file Word.")
                return False

            doc = Document(template_file)
            clean_data = {k: str(v).replace("*", "").replace("#", "").replace("$", "").strip() 
                          for k, v in data_dict.items()}

            def replace_in_paragraphs(paragraphs):
                for p in paragraphs:
                    for key, value in clean_data.items():
                        if key in p.text:
                            p.text = p.text.replace(key, value)
                            for run in p.runs:
                                run.font.name = 'Arial'
                                run.font.size = Pt(11)

            replace_in_paragraphs(doc.paragraphs)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        replace_in_paragraphs(cell.paragraphs)
            
            output_path = os.path.join(current_dir, f"Bao_Cao_Hoan_Thien_{int(time.time())}.docx")
            doc.save(output_path)
            self.write_log(f"✅ ĐÃ XUẤT BÁO CÁO WORD THÀNH CÔNG: {output_path}")
            return True
        except Exception as e:
            self.write_log(f"❌ Lỗi xuất Word: {e}")
            return False

    def process_analysis(self):
        self.run_btn.config(state=tk.DISABLED)
        
        try:
            # 1. Thẩm định nhanh Biên bản đầu vào
            self.write_log("\n🔍 Đang kiểm tra loại file Biên bản thử nghiệm...")
            pdf_bb = self.upload_to_gemini(self.path_bb.get())
            pdf_qt = self.upload_to_gemini(self.path_qt.get())

            system_instruction = """
            Bạn là Chuyên gia Kiểm định Cáp ngầm Cao thế & Trung thế Bậc cao của EVNHCMC.
            Nhiệm vụ: Phân tích chỉ tiêu Phóng điện cục bộ (CBM/PD), đối soát nghiêm ngặt quy trình QT-CT-02
            và thẩm định ĐỘ TRUNG THỰC/CHÍNH XÁC của phần kết luận trên biên bản hiện trường.
            """

            # Kiểm tra nhanh Biên bản
            chk_bb_prompt = """
            Trả về JSON duy nhất: {"is_cbm_pd": true/false, "reason": "chi tiết nếu sai"}
            Tệp 2 có phải là Biên bản thử nghiệm phóng điện cục bộ (PD/CBM) cáp ngầm không?
            """
            bb_res = self.client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[pdf_bb, chk_bb_prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            
            match_bb = re.search(r'\{.*\}', bb_res.text, re.DOTALL)
            if match_bb:
                bb_data = json.loads(match_bb.group(0))
                if not bb_data.get("is_cbm_pd"):
                    self.write_log("❌ LỖI BIÊN BẢN: File không phải Biên bản thử nghiệm CBM (PD)!")
                    messagebox.showerror("Sai Biên Bản", f"Tệp Biên bản không hợp lệ!\nLý do: {bb_data.get('reason')}\nYêu cầu chọn đúng file CBM!")
                    self.run_btn.config(state=tk.NORMAL)
                    return

            # 2. Phân tích Chuyên sâu & Đánh giá Kết luận
            self.write_log("🧠 Chuyên gia AI đang phân tích toàn bộ dữ liệu & Đánh giá kết luận biên bản...")
            
            history_context = ""
            if os.path.exists(KNOWLEDGE_FILE):
                with open(KNOWLEDGE_FILE, "r", encoding="utf-8") as f:
                    hist = json.load(f)
                    if hist:
                        history_context = f"\n[LỊCH SỬ ĐỐI SOÁT ĐỂ ĐẢM BẢO NHẤT QUÁN]:\n{json.dumps(hist[-3:], ensure_ascii=False)}"

            main_prompt = f"""
            Dựa vào Quy trình QT-CT-02 (Tệp 1) và Biên bản Thử nghiệm Phóng điện cục bộ CBM (Tệp 2), hãy thực hiện phân tích chuyên sâu.
            {history_context}

            Hãy xuất thông tin ĐÚNG BẮT BUỘC theo các thẻ sau:

            [THONG_TIN_CHUNG]
            Tên tuyến cáp, chủng loại, chiều dài, điện áp vận hành, đơn vị thực hiện, ngày thử nghiệm.
            [/THONG_TIN_CHUNG]

            [DANH_GIA]
            Phân tích số liệu CBM chi tiết: Điện áp khởi phát PD, Biên độ phóng điện tối đa (pC hoặc dBmV), vị trí điểm PD (m), tần suất xung. So sánh cụ thể với ngưỡng tiêu chuẩn trong QT-CT-02.
            [/DANH_GIA]

            [DANH_GIA_KET_LUAN]
            - Trích dẫn nguyên văn "Nội dung Kết luận" ghi trên Biên bản hiện trường.
            - Đánh giá Chuyên gia: So sánh kết luận gốc đó với dữ liệu kỹ thuật thực tế đo được.
            - Khẳng định rõ ràng: Kết luận trên biên bản là KHỚP (Chính xác) hay KHÔNG KHỚP (Sai lệch/Bỏ sót rủi ro/Chẩn đoán Đạt hay Không Đạt sai quy trình QT-CT-02). Nêu rõ lý do.
            [/DANH_GIA_KET_LUAN]

            [NHAN_DINH]
            Nhận định chuyên môn về mức độ lão hóa cách điện, rủi ro phóng điện hộp nối/đầu cáp, dự báo nguy cơ sự cố.
            [/NHAN_DINH]

            [KHUYEN_NGHI]
            Đề xuất vận hành (Rút ngắn chu kỳ CBM, sửa chữa hộp nối, giảm tải hoặc thay thế).
            [/KHUYEN_NGHI]
            """

            response = self.client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[pdf_qt, pdf_bb, main_prompt],
                config=types.GenerateContentConfig(system_instruction=system_instruction)
            )
            text = response.text
            
            self.write_log("\n" + "="*20 + " KẾT QUẢ THẨM ĐỊNH CHI TIẾT " + "="*20 + "\n")
            self.write_log(text)

            def extract_content(tag, full_text):
                pattern = rf"\[{tag}\](.*?)\[/{tag}\]"
                match = re.search(pattern, full_text, re.DOTALL)
                return match.group(1).strip() if match else "Dữ liệu trống"

            data = {
                "{{THONG_TIN_CHUNG}}": extract_content("THONG_TIN_CHUNG", text),
                "{{DANH_GIA}}": extract_content("DANH_GIA", text),
                "{{DANH_GIA_KET_LUAN}}": extract_content("DANH_GIA_KET_LUAN", text),
                "{{NHAN_DINH}}": extract_content("NHAN_DINH", text),
                "{{KHUYEN_NGHI}}": extract_content("KHUYEN_NGHI", text)
            }

            self.save_to_knowledge_base({
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                "file_bb": os.path.basename(self.path_bb.get()),
                "evaluation_match": data["{{DANH_GIA_KET_LUAN}}"][:250]
            })

            if self.fill_report(data):
                messagebox.showinfo("Thành công", "Đã phân tích & xuất báo cáo Word thành công!")
            
        except Exception as e:
            self.write_log(f"❌ Lỗi hệ thống: {e}")
            messagebox.showerror("Lỗi", str(e))
        
        self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = CableAIApp(root)
    root.mainloop()
