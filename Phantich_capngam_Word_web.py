import time, os, re, threading, json, shutil, tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
from docx.shared import Pt

# Khai báo thư viện Google GenAI SDK mới
import google_genai as genai
from google_genai import types

st.set_page_config(
    page_title="AI Phân Tích & Thẩm Định Cáp Ngầm EVNHCMC",
    page_icon="⚡",
    layout="wide"
)

KNOWLEDGE_FILE = "knowledge_history.json"

if not os.path.exists(KNOWLEDGE_FILE):
    with open(KNOWLEDGE_FILE, "w", encoding="utf-8") as f:
        json.dump([], f, ensure_ascii=False, indent=4)

def save_to_knowledge_base(record_info):
    try:
        with open(KNOWLEDGE_FILE, "r+", encoding="utf-8") as f:
            history = json.load(f)
            history.append(record_info)
            f.seek(0)
            json.dump(history[-20:], f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.warning(f"⚠️ Không thể lưu lịch sử: {e}")

def fill_report_bytes(data_dict, template_path="Bao_Cao_mau.docx"):
    if not os.path.exists(template_path):
        return None

    doc = Document(template_path)
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

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def upload_bytes_to_gemini(client, uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    file = client.files.upload(file=tmp_path)
    while file.state.name == "PROCESSING":
        time.sleep(1)
        file = client.files.get(name=file.name)
    
    os.remove(tmp_path)
    return file

st.title("⚡ AI Phân Tích & Thẩm Định Biên Bản Cáp Ngầm - EVNHCMC")
st.caption("Phiên bản Web SDK mới - Hỗ trợ toàn diện API Key dạng AQ. và AIzaSy")

with st.sidebar:
    st.header("⚙️ Cấu hình hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key (AQ... hoặc AIzaSy...):", type="password")
    st.divider()
    st.markdown("**Hồ sơ yêu cầu:**")
    st.markdown("- Quy trình: `QT-CT-02` (PDF)")
    st.markdown("- Biên bản: `Biên bản CBM/PD` (PDF)")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. File Quy trình kỹ thuật")
    uploaded_qt = st.file_uploader("Tải lên Quy trình thử nghiệm (PDF)", type=["pdf"], key="qt_file")

with col2:
    st.subheader("2. File Biên bản hiện trường")
    uploaded_bb = st.file_uploader("Tải lên Biên bản CBM/PD (PDF)", type=["pdf"], key="bb_file")

if st.button("🚀 BẮT ĐẦU THẨM ĐỊNH & PHÂN TÍCH", type="primary", use_container_width=True):
    if not api_key:
        st.error("❌ Vui lòng nhập API Key ở thanh bên trái.")
    elif not uploaded_qt or not uploaded_bb:
        st.warning("⚠️ Vui lòng tải lên đầy đủ cả 2 tệp: File Quy trình và File Biên bản!")
    else:
        status_box = st.status("Đang tiến hành xử lý hồ sơ...", expanded=True)
        
        try:
            client = genai.Client(api_key=api_key)

            status_box.write("⏳ Đang nạp tài liệu lên máy chủ AI Cloud...")
            pdf_qt = upload_bytes_to_gemini(client, uploaded_qt)
            pdf_bb = upload_bytes_to_gemini(client, uploaded_bb)

            system_instruction = """
            Bạn là Chuyên gia Kiểm định Cáp ngầm Cao thế & Trung thế Bậc cao của EVNHCMC.
            Nhiệm vụ: Phân tích chỉ tiêu Phóng điện cục bộ (CBM/PD), đối soát nghiêm ngặt quy trình QT-CT-02
            và thẩm định ĐỘ TRUNG THỰC/CHÍNH XÁC của phần kết luận trên biên bản hiện trường.
            """

            status_box.write("🔍 Đang kiểm tra tính hợp lệ của hai tệp đầu vào...")
            chk_prompt = """
            Hãy kiểm tra 2 tệp PDF và trả về duy nhất định dạng JSON:
            {
                "is_qt_ct_02": true/false,
                "is_cbm_pd": true/false,
                "reason": "Mô tả lý do cụ thể nếu có file không đúng"
            }
            """
            verify_res = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[pdf_qt, pdf_bb, chk_prompt],
                config=types.GenerateContentConfig(
                    system_instruction="Bạn là trợ lý kiểm tra hồ sơ kỹ thuật.",
                    response_mime_type="application/json"
                )
            )
            
            valid = True
            try:
                chk_data = json.loads(verify_res.text)
                if not chk_data.get("is_qt_ct_02"):
                    status_box.update(label="Thẩm định thất bại!", state="error")
                    st.error(f"❌ **Lỗi File Quy Trình:** Tệp 1 không phải là quy trình QT-CT-02. Chi tiết: {chk_data.get('reason')}")
                    valid = False
                elif not chk_data.get("is_cbm_pd"):
                    status_box.update(label="Thẩm định thất bại!", state="error")
                    st.error(f"❌ **Lỗi File Biên Bản:** Tệp 2 không phải là biên bản thử nghiệm CBM (PD). Chi tiết: {chk_data.get('reason')}")
                    valid = False
            except Exception:
                pass
            
            if valid:
                status_box.write("✅ Hồ sơ hợp lệ! Đang bóc tách thông số và thẩm định kết luận...")
                
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

                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[pdf_qt, pdf_bb, main_prompt],
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction
                    )
                )
                text = response.text

                status_box.update(label="Thẩm định hoàn tất thành công!", state="complete", expanded=False)

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

                save_to_knowledge_base({
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "file_bb": uploaded_bb.name,
                    "evaluation_match": data["{{DANH_GIA_KET_LUAN}}"][:250]
                })

                st.subheader("📑 Kết quả thẩm định & Đối soát kỹ thuật")
                
                with st.expander("📌 1. Thông tin chung tuyến cáp", expanded=True):
                    st.write(data["{{THONG_TIN_CHUNG}}"])

                with st.expander("📊 2. Đánh giá số liệu đo CBM so với QT-CT-02", expanded=True):
                    st.write(data["{{DANH_GIA}}"])

                with st.expander("🔍 3. Thẩm định tính chuẩn xác của Kết luận Biên bản", expanded=True):
                    if "KHÔNG KHỚP" in data["{{DANH_GIA_KET_LUAN}}"].upper():
                        st.error(data["{{DANH_GIA_KET_LUAN}}"])
                    else:
                        st.success(data["{{DANH_GIA_KET_LUAN}}"])

                with st.expander("⚠️ 4. Nhận định rủi ro & Cách điện", expanded=True):
                    st.write(data["{{NHAN_DINH}}"])

                with st.expander("💡 5. Khuyến nghị giải pháp vận hành", expanded=True):
                    st.write(data["{{KHUYEN_NGHI}}"])

                doc_bytes = fill_report_bytes(data)
                if doc_bytes:
                    st.download_button(
                        label="📥 TẢI XUỐNG BÁO CÁO HOÀN THIỆN (.DOCX)",
                        data=doc_bytes,
                        file_name=f"Bao_Cao_CBM_{int(time.time())}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.info("ℹ️ Không tìm thấy file `Bao_Cao_mau.docx` trong thư mục chạy để tạo file tải về.")

        except Exception as e:
            status_box.update(label="Có lỗi phát sinh!", state="error")
            st.error(f"❌ Lỗi hệ thống: {e}")
