import streamlit as st
from docx import Document
from docx.shared import Inches
import io

st.set_page_config(page_title="현장 사진대지 생성기", layout="wide")
st.title("🏗️ 공사 항목별 사진대지 개별 생성")

# 세션 상태 초기화
if 'sections' not in st.session_state:
    st.session_state.sections = [{"title": "", "before": [], "after": []}]

def add_section():
    st.session_state.sections.append({"title": "", "before": [], "after": []})

def remove_section(index):
    if len(st.session_state.sections) > 1:
        st.session_state.sections.pop(index)

# 상단 항목 추가 버튼
st.button("➕ 새로운 공정(항목) 추가", on_click=add_section)

# 입력 UI 구성
for i, section in enumerate(st.session_state.sections):
    with st.container():
        st.markdown(f"### [항목 {i+1}]")
        col_title, col_del = st.columns([8, 1])
        with col_title:
            section['title'] = st.text_input(f"작업 명칭 입력", key=f"title_{i}", value=section['title'], placeholder="예: 정화조 철거, 내장재 해체 등")
        with col_del:
            if st.button("🗑️ 삭제", key=f"del_{i}"):
                remove_section(i)
                st.rerun()

        col_before, col_after = st.columns(2)
        with col_before:
            st.subheader("📷 철거 '전' 사진")
            section['before'] = st.file_uploader(f"항목 {i+1} 전 사진들", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True, key=f"before_{i}")
        with col_after:
            st.subheader("📸 철거 '후' 사진")
            section['after'] = st.file_uploader(f"항목 {i+1} 후 사진들", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True, key=f"after_{i}")
        st.divider()

# --- 개별 워드 파일 생성 함수 ---
def create_docx_content(mode="before"):
    doc = Document()
    suffix = " (철거 전)" if mode == "before" else " (철거 후)"
    doc.add_heading("공사 사진 대지" + suffix, 0)

    for section in st.session_state.sections:
        photos = section['before'] if mode == "before" else section['after']
        if not photos:
            continue
            
        current_title = section['title'] if section['title'] else "제목 없음"
        doc.add_heading(current_title, level=1)
        
        # 사진을 한 줄에 하나씩 배치하는 표
        table = doc.add_table(rows=0, cols=1)
        table.style = 'Table Grid'
        hdr = table.add_row().cells
        hdr[0].text = "촬영 사진 (" + ("전" if mode == "before" else "후") + ")"
        
        for photo in photos:
            row = table.add_row().cells
            # 사진 너비를 5.5인치로 큼직하게 설정 (A4 출력 최적화)
            row[0].paragraphs[0].add_run().add_picture(photo, width=Inches(5.5))
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- 하단 개별 다운로드 버튼 ---
st.write("### 🖨️ 원하는 문서를 선택해서 다운로드하세요")
btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    before_data = create_docx_content("before")
    st.download_button(
        label="📄 [철거 전] 문서만 생성 및 다운로드",
        data=before_data,
        file_name="01_철거전_사진대지.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

with btn_col2:
    after_data = create_docx_content("after")
    st.download_button(
        label="📄 [철거 후] 문서만 생성 및 다운로드",
        data=after_data,
        file_name="02_철거후_사진대지.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )