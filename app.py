# -*- coding: utf-8 -*-
import os
import sys
import tempfile
import streamlit as st

# Add current directory to path so python can find report_generator.py
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
from report_generator import generate_report

# Configure Streamlit page layout and aesthetics
st.set_page_config(
    page_title="KOOS RESULT 결과보고서 자동 생성기",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling using CSS
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+KR:wght@300;400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', 'Noto Sans KR', sans-serif;
    }
    
    .main-title {
        font-size: 2.8rem;
        font-weight: 700;
        background: linear-gradient(135deg, #2D6CDF, #00C9FF);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        font-size: 1.1rem;
        color: #6c757d;
        margin-bottom: 2rem;
    }
    
    .card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #eef1f6;
        margin-bottom: 20px;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #2D6CDF, #1B4FB0);
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
        border: none;
        box-shadow: 0 4px 10px rgba(45, 108, 223, 0.2);
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(45, 108, 223, 0.3);
    }
    
    .sidebar-header {
        font-size: 1.3rem;
        font-weight: 700;
        color: #2D6CDF;
        margin-bottom: 15px;
    }
    
    .sidebar-section {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        border-left: 4px solid #2D6CDF;
    }
    
    .download-link-btn {
        display: inline-block;
        width: 100%;
        text-align: center;
        background-color: #f1f3f7;
        color: #2D6CDF;
        padding: 8px 12px;
        border-radius: 6px;
        text-decoration: none;
        font-weight: 600;
        margin-bottom: 8px;
        border: 1px solid #DFE4EC;
        transition: all 0.2s ease;
    }
    .download-link-btn:hover {
        background-color: #EAF1FF;
        border-color: #2D6CDF;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Helper function to read file as bytes
def get_file_bytes(filepath):
    if os.path.exists(filepath):
        with open(filepath, "rb") as f:
            return f.read()
    return None

# Locate default files
current_dir = os.path.dirname(os.path.abspath(__file__))
default_template_path = os.path.join(current_dir, "템플릿_결보표양.pptx")
if not os.path.exists(default_template_path):
    default_template_path = os.path.join(current_dir, "assets", "템플릿_결보표양.pptx")

default_config_path = os.path.join(current_dir, "설정.xlsx")
if not os.path.exists(default_config_path):
    default_config_path = os.path.join(current_dir, "assets", "설정_예시.xlsx")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.markdown("<div class='sidebar-header'>📊 KOOS RESULT</div>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 0.9rem; color:#666;'>교육 만족도 결과보고서 PPT를 자동생성해주는 웹 서비스입니다.</p>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("<div style='font-weight:600; margin-bottom:10px;'>리소스 다운로드</div>", unsafe_allow_html=True)
    
    # Download buttons for template and sample settings
    template_bytes = get_file_bytes(default_template_path)
    if template_bytes:
        st.download_button(
            label="📄 기본 PPT 템플릿 다운로드",
            data=template_bytes,
            file_name="템플릿_결보표양.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
    
    config_bytes = get_file_bytes(default_config_path)
    if config_bytes:
        st.download_button(
            label="⚙️ 기본 설정 엑셀 다운로드",
            data=config_bytes,
            file_name="설정.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    st.markdown("---")
    st.markdown("<div style='font-weight:600; margin-bottom:10px;'>이용 안내</div>", unsafe_allow_html=True)
    with st.expander("💡 척도 자동 감지 규칙"):
        st.markdown(
            """
            - 5점 그렇다형, 만족형, 우수형, 도움형
            - 4점 그렇다형, 만족형
            - 7점 그렇다형 등을 자동 감지합니다.
            - 응답의 60% 이상이 특정 척도 라벨을 가질 때 작동합니다.
            """
        )
    with st.expander("💡 주관식 자동 분류"):
        st.markdown(
            """
            - 문항 제목에 `좋았던·장점·우수·인상` 등이 포함되면 **좋았던 점**으로 분류됩니다.
            - `아쉬·개선·보완·불편·단점·건의` 등이 포함되면 **아쉬웠던 점**으로 분류됩니다.
            - 아쉬웠던 점의 '없음', '없습니다' 등 무의미한 답변은 기본적으로 자동 제외됩니다.
            """
        )

# ---------------- Main Page ----------------
st.markdown("<div class='main-title'>교육 만족도 결과보고서 자동 생성기</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>설문 로우데이터와 설정 값만 입력하면 시각화 결과 보고서 PPT가 즉시 만들어집니다.</div>", unsafe_allow_html=True)

# Grid Layout
col_left, col_right = st.columns([1.1, 0.9])

with col_left:
    st.markdown("### 1. 입력 파일 업로드")
    
    # 1. Raw survey data upload
    raw_file = st.file_uploader(
        "📊 설문 로우데이터 엑셀 또는 CSV 파일 업로드 (필수)",
        type=["xlsx", "xls", "xlsm", "csv"],
        help="구글폼 등에서 내려받은 응답 원본 파일입니다."
    )
    
    st.markdown("---")
    st.markdown("### 2. 보고서 세부 정보 설정")
    
    # Selection of configuration method
    config_mode = st.radio(
        "설정 값 입력 방식 선택",
        options=["직접 입력 (추천 - 엑셀 파일 없이 직접 타이핑)", "엑셀 파일 업로드 (설정.xlsx 사용)"],
        horizontal=True
    )
    
    config_dict = None
    config_file = None
    
    if "직접 입력" in config_mode:
        st.info("💡 아래 입력 칸에 정보를 적어주세요. '설정.xlsx'를 따로 업로드하지 않아도 됩니다.")
        
        tab_basic, tab_overview = st.tabs(["📝 표지 기본정보", "🏫 교육 개요"])
        
        with tab_basic:
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                cover_label = st.text_input("표지 상단 라벨", value="결과보고서")
                title1 = st.text_input("표지 대제목 1줄", value="SK TNS FLP 대상")
                title2 = st.text_input("표지 대제목 2줄", value="인사이트 트립 결과보고서")
            with col_b2:
                instructor = st.text_input("강사명", value="설상훈 교수")
                student_count = st.number_input("수강인원(명)", min_value=1, value=9, step=1,
                                                help="응답률 계산에 활용됩니다. (응답자 수 ÷ 수강인원 × 100)")
                
        with tab_overview:
            course_name = st.text_input("과정명", value="인사이트 트립")
            schedule = st.text_input("교육일정", value="2026.07.08")
            method = st.text_input("교육방식", value="대면 교육")
            target = st.text_input("교육대상", value="핵심 실무진")
            goal = st.text_area("교육목표", value="트렌드 학습 및 비즈니스 인사이트 도출")
            
        # Build configuration dict
        config_dict = {
            "basic": {
                "표지_상단라벨": cover_label,
                "표지_제목1": title1,
                "표지_제목2": title2,
                "강사명": instructor,
                "수강인원": int(student_count)
            },
            "overview": {
                "과정명": course_name,
                "교육일정": schedule,
                "교육방식": method,
                "교육대상": target,
                "교육목표": goal
            },
            "questions": []
        }
        
    else:
        st.info("📂 '설정.xlsx' 파일을 업로드해 주세요. (다운로드 탭에서 예시를 내려받아 수정할 수 있습니다.)")
        config_file = st.file_uploader(
            "⚙️ 설정 엑셀 파일 업로드",
            type=["xlsx", "xls", "xlsm"],
            help="표지 제목, 교육 개요 등이 기록된 설정 엑셀 파일입니다."
        )

with col_right:
    st.markdown("### 3. 추가 옵션")
    
    # Option: Keep no opinion
    keep_no_opinion = st.checkbox(
        "아쉬웠던 점의 '없음'류 의견도 그대로 표시",
        value=False,
        help="체크 해제 시 '없음', '없습니다', '해당사항 없음' 같은 짧고 무의미한 의견을 필터링합니다."
    )
    
    # Custom template PPTX
    use_custom_template = st.checkbox(
        "사용자 정의 PPT 템플릿 파일 업로드",
        value=False,
        help="체크 시 나만의 PPT 템플릿(.pptx)을 업로드하여 반영할 수 있습니다."
    )
    
    template_file = None
    if use_custom_template:
        template_file = st.file_uploader(
            "🎨 템플릿 PPTX 파일 업로드",
            type=["pptx"],
            help="결과보고서 마스터 레이아웃이 적용된 파워포인트 템플릿입니다."
        )
    else:
        st.caption(f"💡 기본 템플릿 (`{os.path.basename(default_template_path or '템플릿_결보표양.pptx')}`)을 사용합니다.")
        
    st.markdown("---")
    st.markdown("### 4. 보고서 생성")
    
    # Button to generate report
    generate_btn = st.button("결과보고서 만들기 🚀", use_container_width=True)
    
    # Container for generation logs and status
    log_area = st.empty()
    download_area = st.empty()

    if generate_btn:
        # Validations
        if not raw_file:
            st.error("❌ 설문 로우데이터 파일을 업로드해 주세요.")
        elif "엑셀 파일 업로드" in config_mode and not config_file:
            st.error("❌ 설정 엑셀 파일(설정.xlsx)을 업로드해 주세요.")
        elif use_custom_template and not template_file:
            st.error("❌ 템플릿 PPTX 파일을 업로드해 주세요.")
        else:
            # Create a temporary directory to save and parse files
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save Raw File
                raw_ext = os.path.splitext(raw_file.name)[1]
                raw_temp_path = os.path.join(tmpdir, f"raw_data{raw_ext}")
                with open(raw_temp_path, "wb") as f:
                    f.write(raw_file.getbuffer())
                
                # Save Config File or Dict
                config_arg = None
                if config_file:
                    config_ext = os.path.splitext(config_file.name)[1]
                    config_temp_path = os.path.join(tmpdir, f"config{config_ext}")
                    with open(config_temp_path, "wb") as f:
                        f.write(config_file.getbuffer())
                    config_arg = config_temp_path
                else:
                    config_arg = config_dict
                    
                # Save Template File
                template_temp_path = default_template_path
                if template_file:
                    template_temp_path = os.path.join(tmpdir, "template.pptx")
                    with open(template_temp_path, "wb") as f:
                        f.write(template_file.getbuffer())
                
                if not template_temp_path or not os.path.exists(template_temp_path):
                    st.error("❌ 기본 템플릿 파일을 찾을 수 없습니다. 템플릿 파일을 직접 업로드해 주세요.")
                else:
                    # Output path
                    output_temp_path = os.path.join(tmpdir, "output.pptx")
                    
                    # Set up logging callback for Streamlit UI
                    logs = []
                    
                    def streamlit_log(message):
                        logs.append(message)
                        # Display latest 8 log lines in real-time
                        log_area.code("\n".join(logs[-10:]))
                    
                    # Run the generator
                    try:
                        with st.spinner("결과보고서 생성 중..."):
                            streamlit_log("[시작] 교육보고서 생성을 시작합니다...")
                            generate_report(
                                raw_path=raw_temp_path,
                                config=config_arg,
                                template_path=template_temp_path,
                                output_path=output_temp_path,
                                drop_no_opinion=not keep_no_opinion,
                                log=streamlit_log
                            )
                            
                        # If successful, read generated file to serve for download
                        if os.path.exists(output_temp_path):
                            with open(output_temp_path, "rb") as f:
                                result_pptx_bytes = f.read()
                            
                            # Determine download filename
                            dl_filename = "교육결과보고서.pptx"
                            if config_dict:
                                title = config_dict["basic"].get("표지_제목2") or config_dict["basic"].get("표지_제목1")
                                if title:
                                    dl_filename = f"{title.replace(' ', '_')}_결과보고서.pptx"
                            elif config_file:
                                try:
                                    import openpyxl
                                    wb = openpyxl.load_workbook(config_temp_path, data_only=True)
                                    if "기본정보" in wb.sheetnames:
                                        ws = wb["기본정보"]
                                        title = ""
                                        for r in range(1, ws.max_row + 1):
                                            k = ws.cell(row=r, column=1).value
                                            v = ws.cell(row=r, column=2).value
                                            if k and str(k).strip() in ("표지_제목2", "표지_제목1") and v:
                                                title = str(v).strip()
                                                break
                                        if title:
                                            dl_filename = f"{title.replace(' ', '_')}_결과보고서.pptx"
                                except Exception:
                                    pass
                            
                            st.success("🎉 결과보고서 생성이 완료되었습니다!")
                            
                            # Render Download Button
                            download_area.download_button(
                                label="📥 결과보고서 다운로드 (.pptx)",
                                data=result_pptx_bytes,
                                file_name=dl_filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ 결과 파일 생성에 실패했습니다. 로그를 확인하세요.")
                    except Exception as e:
                        st.error(f"❌ 생성 도중 오류가 발생했습니다: {str(e)}")
                        st.exception(e)
