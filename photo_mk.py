import streamlit as st
from docx import Document
from docx.shared import Inches
import io

# 1. 연결 확인용 메시지 (성공하면 삭제하셔도 됩니다)
st.success("✅ 프로그램이 최신 버전으로 업데이트되었습니다!")

st.title("🏗️ 공사 사진대지 개별 생성기")

# 세션 상태 (항목 저장용)
if 'sections' not in st.session_state:
    st.session_state.sections = [{"title": "", "before": [], "after": []}]

# 항목 추가/삭제 함수
def add_section():
    st.session_state.sections.append({"title": "", "before": [], "after": []})
def remove_section(idx):
    if len(st.session_state.sections) > 1:
        st.session_state.sections.pop(idx)

st.button("➕ 새로운 공정 항목 추가", on_click=add_section)

# 입력 칸 만들기
for i, section in enumerate(st.session_state.sections):
    with st.expander(f"항목 {i+1}: {section['title'] if section['title'] else '제목을 입력하세요'}", expanded=True):
        col_t, col_d = st.columns([7, 1])
        section['title'] = col_t.text_input("작업 명칭", key=f"t_{i}", value=section['title'])
        if col_d.button("🗑️ 삭제", key=f"d_{i}"):
            remove_section(i)
            st.rerun()

        c_bef, c_aft = st.columns(2)
        section['before'] = c_bef.file_uploader(f"[{i+1}] 전 사진들", type=['jpg','jpeg','png'], accept_multiple_files=True, key=f"b_{i}")
        section['after'] = c_aft.file_uploader(f"[{i+1}] 후 사진들", type=['jpg','jpeg','png'], accept_multiple_files=True, key=f"a_{i}")

st.divider()

# --- 워드 파일 생성 로직 (함수화) ---
def make_doc(mode="before"):
    doc = Document()
    name = "철거 전" if mode == "before" else "철거 후"
    doc.add_heading(f"공사 사진 대지 ({name})", 0)

    for s in st.session_state.sections:
        photos = s['before'] if mode == "before" else s['after']
        if not photos: continue # 사진 없으면 통과
        
        doc.add_heading(s['title'] if s['title'] else "제목 없음", level=1)
        table = doc.add_table(rows=0, cols=1)
        table.style = 'Table Grid'
        
        for p in photos:
            row = table.add_row().cells
            row[0].paragraphs[0].add_run().add_picture(p, width=Inches(5.5))
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 버튼 배치 (이 부분은 조건문 밖에 있어서 무조건 보입니다) ---
st.write("### 📥 아래 버튼을 눌러 각각 다운로드 하세요")
col1, col2 = st.columns(2)

# 전용 문서 버튼
before_file = make_doc("before")
col1.download_button(
    label="📂 [철거 전] 문서 다운로드",
    data=before_file,
    file_name="01_철거전_사진대지.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)

# 후용 문서 버튼
after_file = make_doc("after")
col2.download_button(
    label="📂 [철거 후] 문서 다운로드",
    data=after_file,
    file_name="02_철거후_사진대지.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)