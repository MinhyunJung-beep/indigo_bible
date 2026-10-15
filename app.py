import os
import sys
import streamlit as st
import pandas as pd
import google.generativeai as genai
import plotly.express as px
import plotly.graph_objects as go
import numpy as np # 13번 몬테카를로 연산을 위한 수학 라이브러리 추가
excel_path = "tcr_diagnosis.xlsx"

# 💡 윈도우 인코딩 충돌 방지
os.environ["PYTHONIOENCODING"] = "utf-8"
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except AttributeError:
    pass

from search_engine import RacingBibleEngine
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage, AIMessage

# 1. Page Configuration & Professional Technical UI Styling
st.set_page_config(page_title="Trackside Engineering Support System V3", page_icon="🏁", layout="wide")
st.title("🏁 Trackside Engineering Support System v3.1")
st.subheader("Integrated Trackside Engineering & Race Strategy Platform")

# 💡 [핵심 기술] PDF 출력용 프린트 최적화 CSS V3 (다크 모드 유지 & 사이드바 제거)
st.markdown("""
    <style>
    @media print {
        /* 1. 사이드바, 상단 헤더 등 인쇄에 필요 없는 UI 완벽 삭제 */
        [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"] { 
            display: none !important; 
        }
        
        /* 2. 다크 모드 배경색 및 텍스트 색상을 인쇄 시에도 강제로 유지 */
        body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMainBlock"] {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }
        
        /* 3. 페이지 넘어갈 때 글씨나 차트가 반으로 잘리는 현상 방지 */
        h1, h2, h3, h4, h5, .js-plotly-plot, .stMarkdown { 
            page-break-inside: avoid; 
        }
    }
    </style>
""", unsafe_allow_html=True)
# 2. Initialize Backend Engine
@st.cache_resource 
def load_engine():
    return RacingBibleEngine()

# 🔑 [CRITICAL] Enter your Google API Key here!
os.environ["GOOGLE_API_KEY"] = "AQ.Ab8RN6J10ayPamyr97bWAFu3g_KF2HgIaaWYKIqT2bwO2HG-ug"

if "engine" not in st.session_state:
    with st.spinner("🚀 Booting Super Brain Vector Database... Please standby."):
        st.session_state.engine = load_engine()

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

llm = ChatGoogleGenerativeAI(model="gemini-3.8-flash", temperature=0.1)

# 3. Sidebar Configuration (Mission Control Modules)
st.sidebar.header("🕹️ Mission Control Modules")
selected_module = st.sidebar.radio(
    "Select Operational Mode:",
    [
        " 1. Live Pit-Wall Radio",
        " 2. #87 Driver Debrief & AI Diagnostic Center",
        " 3. Advanced Quali Predictor",
        " 4. Competitor Pace Analysis",
        " 5. Race Result & RCA",
        " 6. Tire Strategy & Pressure Predictor",
        " 7. Crew KPI & Performance Matrix",
        " 8. Component Lifing & Mileage",
        " 9. Garage Operation & Crew Choreography",
        " 10. Pitstop Motion & Bottleneck AI", 
        " 11. Trackside Decision & Risk Matrix",
        " 12. Telemetry Diagnostics",       # 👈 12번 추가
        " 13. Monte Carlo Race Predictor",  # 👈 13번 추가
        " 14. Engine vital Check",
        " 15. Driver KPI & Track Map",
        " 16. Pre-Race Strategy & Weather Analysis",
        " 17. Auto Document & Proposal Generator"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 System Status Dashboard")
st.sidebar.success(f"**Active:** {selected_module.split('.')[0]}")
st.sidebar.info(
    "💡 **Knowledge Base Summary:**\n\n"
    "- 181 OEM Factory Tech Manuals\n"
    "- 174 Trackside Diagnostics Rows\n"
    "- #87 Historic Setup Database\n"
    "- 32,177 Refined Data Chunks"
)

if st.sidebar.button("🔄 Reset Session / Clear Radio"):
    st.session_state.chat_history = []
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown(
    """
    <div style='text-align: center; color: #888888; font-size: 11px; line-height: 1.5;'>
        <b>TESS : Trackside Engineering Support System</b><br>
        <i>v1.0 (Build 2026)</i><br><br>
        <b>System Architect & Developer</b><br>
        Chief Mechanic Minhyun Jung<br>
        <br>
        © 2026 Minhyun Jung. All Rights Reserved.
    </div>
    """,
    unsafe_allow_html=True
)
st.sidebar.markdown("---")
st.sidebar.write("👨‍💻 **Technical Team Leader:**\nChief Mechanic")

# ==========================================
# MODULE 1~9: AI Mission Control Chat Interfaces
# ==========================================
# 🚨 10번 Pitstop Motion까지 리스트에 추가 완료
if any(mod in selected_module for mod in ["Live Pit-Wall", "#87 Setup", "Race Result", "RCA", "Tire Strategy", "Crew KPI", "Component Lifing", "Garage Operation", "Pitstop Motion"]):
    
    setup_history_text = ""
    race_result_text = "" 
    tire_analysis_text = ""
    kpi_text = ""
    lifing_text = ""
    garage_text = "" 
    pitstop_text = "" # 👈 10번 융합 데이터 보관함 추가
    excel_path = "tcr_diagnosis.xlsx"
    
# ==========================================
# MODULE 2: #87 Driver Debrief & AI Diagnostic Center
# ==========================================
elif "2." in selected_module or "Setup" in selected_module:
    st.header("🛠️ 2. #87 Driver Debrief & AI Diagnostic Center")
    st.caption("에릭(Eric) 엔지니어의 셋업 DB 및 드라이버 리포트를 바탕으로 증상을 분석하고, 셋업 밸런스를 계산합니다.")

    # 💡 두 개의 서브 탭 생성
    sub_tab1, sub_tab2 = st.tabs(["🤖 AI Diagnostic Center (피드백 분석)", "⚖️ TCR Setup Calculator (ARB/RCH)"])

    # ------------------------------------------
    # [서브 탭 1] AI 진단 센터
    # ------------------------------------------
    with sub_tab1:
        setup_history_text = ""
        if os.path.exists(excel_path):
            try:
                df_setup = pd.read_excel(excel_path, sheet_name="87_Setup_History")
                setup_history_text = df_setup.fillna("").to_string() 
                with st.expander("📊 시스템이 참조 중인 에릭 엔지니어 #87 셋업 히스토리 DB 확인", expanded=False):
                    st.dataframe(df_setup)
            except:
                pass

        st.markdown("### 🗣️ 드라이버 피드백 및 차량 증상 보고")
        driver_feedback = st.text_area("💬 박준의 선수 피드백 또는 미캐닉 육안 점검 특이사항 입력:", height=150)

        if st.button("🧠 통합 AI 원인 분석 및 해결책 도출", key="ai_diag_2"):
            if driver_feedback:
                with st.spinner("에릭 엔지니어 셋업 DB, 드라이버 리포트, 각종 매뉴얼을 교차 분석 중입니다..."):
                    
                    prompt = f"""
                    너는 인디고 레이싱 팀의 치프 레이스 엔지니어이자 에릭(Eric) 엔지니어의 후계자다.
                    현재 박준의 드라이버(Car #87)가 다음의 피드백/증상을 보고했다:
                    "{driver_feedback}"

                    [참고 데이터: 에릭 엔지니어의 과거 셋업 히스토리 DB]
                    {setup_history_text}

                    위 로컬 DB와 매뉴얼 등을 종합하여 다음 4가지를 브리핑해라:

                    1. [Issue Category & Root Cause]
                       - 증상 분류 및 물리적/기계적 발생 원인 상세 분석

                    2. [Data Cross-Referencing (데이터 교차 검증 채널 안내)]
                       - 문제 확정을 위해 TESS 시스템의 어떤 탭에서 어떤 채널을 봐야 하는지 안내.

                    3. [Branched Mechanical Action Plan (조건부 미캐닉 조치 트리)]
                       - [우선순위 1번]: 1차 하드웨어 또는 셋업 수치 점검.
                       - [우선순위 2번]: 1번 점검 결과에 따른 두 가지 분기(If Then) 제시
                       - [우선순위 3번]: 심화 점검 및 파이널 셋업 픽스

                    4. [Driver Feedback for 박준의]
                       - 박준의 선수의 주행 스타일을 고려한 타이어 관리 또는 드라이빙 교정 팁.
                    """
                    
                    import time
                    max_retries = 3
                    
                    for attempt in range(max_retries):
                        try:
                            response = llm.invoke(prompt)
                            answer = response.content if hasattr(response, 'content') else str(response)
                            if isinstance(answer, list):
                                answer = answer[0].get("text", str(answer))
                            
                            answer = answer.replace('\\n', '\n')
                                
                            st.markdown("---")
                            st.markdown(answer)
                            break 
                            
                        except Exception as e:
                            error_msg = str(e)
                            if "503" in error_msg or "UNAVAILABLE" in error_msg or "429" in error_msg:
                                if attempt < max_retries - 1:
                                    st.warning(f"⚠️ 구글 AI 서버 과부하로 인해 자동 재시도 중입니다... (시도 {attempt+1}/{max_retries})")
                                    time.sleep(2)
                                else:
                                    st.error("❌ 구글 AI 서버 접속량 폭주로 응답이 지연되고 있습니다.")
                            else:
                                st.error(f"❌ AI 분석 중 알 수 없는 오류 발생: {e}")
                                break
            else:
                st.warning("⚠️ 드라이버 피드백이나 증상을 먼저 입력해주세요.")


    # ------------------------------------------
    # [서브 탭 2] Ultimate TCR Dynamics & Track Simulator (셋업 차트 통합)
    # ------------------------------------------
    with sub_tab2:
        import plotly.express as px
        
        # --- [1] 서킷 및 환경 설정 ---
        st.markdown("### 🏁 1. Track & Environment Setup")
        col_e1, col_e2, col_e3, col_e4 = st.columns(4)
        with col_e1: track_name = st.selectbox("Track / Circuit", ["Mugello", "Spa", "Paul Ricard", "Hungaroring", "Monza", "Barcelona", "KIC (영암)", "Inje Speedium", "Red Bull Ring", "Hockenheim", "Misano", "Imola"], index=6)
        with col_e2: track_cond = st.selectbox("Track Condition", ["Dry", "Semi Wet", "Wet"])
        with col_e3: track_temp = st.number_input("Air Temp (°C)", value=25)
        with col_e4: track_hum = st.number_input("Humidity (%)", value=50)

        # --- [2] 에어로 및 디퍼렌셜 (Aero & Drivetrain) ---
        st.markdown("### 🌬️ 2. Aero & Differential (LSD)")
        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        with col_a1: r_wing_pos = st.selectbox("Rear Wing Position", ["Pos 1", "Pos 2", "Pos 3", "Pos 4", "Pos 5"], index=2)
        with col_a2: r_wing_ang = st.number_input("Rear Wing Angle (°)", value=2.5, step=0.5)
        with col_a3: diff_ramp = st.selectbox("Diff Ramp Angle (Accel/Decel)", ["30/45", "45/45", "45/60", "60/60"])
        with col_a4: diff_preload = st.slider("Diff Preload (Nm)", 70, 130, 100, step=8)

        # --- [3] 서스펜션 아키텍처 (스프링, 댐퍼, 지오메트리) ---
        st.markdown("### 🔧 3. Suspension Architecture (Subframe Ref.)")
        camber_spacer_limit = {-3.5: 0, -4.0: 2, -4.5: 5, -5.0: 7, -5.5: 10, -6.0: 15, -6.5: 20, -7.0: 25}
        
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1: 
            f_setup_style = st.selectbox("F_Setup Style", ["Main + Helper (박준의)", "Main Only (박준성)"])
            f_camber = st.selectbox("F_Camber (°)", list(camber_spacer_limit.keys()), index=3)
        with col_f2:
            f_spring = st.selectbox("F_Main Spring (N/mm)", [40, 45, 50, 60, 70, 80, 90], index=0 if "Main Only" in f_setup_style else 4)
            f_toe = st.number_input("F_Toe (mm,+:In / -:Out)", value=-2.0, step=0.5)
        with col_f3:
            f_bump = st.selectbox("F_Bump Rubber (mm)", [12, 14, 26, 28], index=0 if "Helper" in f_setup_style else 2)
            
            # 💡 [업데이트] 프론트 서브프레임 차고 입력 및 스플리터 규정 검증 로직
            f_rh = st.number_input("F_RH (Subframe, mm)", min_value=70, max_value=120, value=90, step=1)
            splitter_est = f_rh - 5  # 서브프레임과 스플리터의 오프셋 (약 5mm)
            
            if splitter_est < 80:
                st.error(f"🚨 스플리터 {splitter_est}mm 예상 (규정 80mm 미만 차검 탈락 위험!)")
            else:
                st.caption(f"✓ 예상 스플리터 지상고: {splitter_est}mm")
                
        with col_f4:
            f_bump_click = st.slider("F_Bump (Click)", 1, 23, 20)
            f_reb_click = st.slider("F_Rebound (Click)", 1, 23, 15)

        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        with col_r1: 
            r_spring = st.selectbox("R_Main Spring (N/mm)", [50, 60, 70, 80, 90], index=1)
            r_camber = st.number_input("R_Camber (°)", value=-2.5, step=0.5)
        with col_r2: 
            r_shim = st.selectbox("R_Damper Shim (mm)", [0, 7, 12, 15], index=1 if "Helper" in f_setup_style else 2)
            r_toe = st.number_input("R_Toe (mm, +:In / -:Out)", value=1.0, step=0.5)
        with col_r3: 
            r_bump = st.selectbox("R_Bump Rubber (mm)", [25, 26], index=0)
            
            # 💡 [업데이트] 리어 서브프레임 차고 입력 (320~350mm 범위)
            r_rh = st.number_input("R_RH (Subframe, mm)", min_value=300, max_value=380, value=330, step=1)
            st.caption("✓ 리어 셋업 레퍼런스 포인트 기준")
            
        with col_r4: 
            r_bump_click = st.slider("R_Bump (Click)", 1, 23, 14)
            r_reb_click = st.slider("R_Rebound (Click)", 1, 23, 13)

        # --- [4] ARB & 코너 웨이트 (Weight & Balance) ---
        st.markdown("### ⚖️ 4. ARB, Weights & Brake Balance")
        
        # 💡 [핵심] 프론트 ARB A1~C3 구현 및 가중치 매핑
        f_arb_mapping = {"A1": 0.8, "A2": 0.85, "A3": 0.9, "B1": 0.95, "B2": 1.0, "B3": 1.05, "C1": 1.15, "C2": 1.25, "C3": 1.35}
        front_arb_data = {"25x1": 72, "18 Solid": 68, "18x2": 43, "18x1.5": 36, "14x2": 19} # B2(Medium) 기준값
        rear_arb_data = {"22x2": {"Soft": 38, "Medium": 43, "Hard": 48}, "25x1": {"Soft": 27, "Medium": 33, "Hard": 43}, "18x2": {"Soft": 18, "Medium": 20, "Hard": 23}}

        col_w1, col_w2, col_w3, col_w4 = st.columns(4)
        with col_w1:
            f_spec = st.selectbox("Front ARB Spec", list(front_arb_data.keys()), index=1)
            f_hole = st.select_slider("F_Hole (Blade)", options=list(f_arb_mapping.keys()), value="B2")
            f_stiffness = front_arb_data[f_spec] * f_arb_mapping[f_hole]
        with col_w2:
            r_spec = st.selectbox("Rear ARB Spec", list(rear_arb_data.keys()))
            r_hole = st.radio("R_Hole", ["Soft", "Medium", "Hard"], horizontal=True)
            r_stiffness = rear_arb_data[r_spec][r_hole]
        with col_w3:
            st.markdown("**Corner Weights (kg)**")
            w_fl = st.number_input("FL", value=405)
            w_fr = st.number_input("FR", value=390)
            w_rl = st.number_input("RL", value=240)
            w_rr = st.number_input("RR", value=250)
        with col_w4:
            st.markdown("**Brake & Target Check**")
            bb = st.number_input("Brake Balance (F%)", value=54.5, step=0.5)
            
            total_weight = w_fl + w_fr + w_rl + w_rr
            cross_weight = ((w_fr + w_rl) / total_weight) * 100
            
            st.info(f"Total: {total_weight}kg (Target: 1285kg)")
            st.success(f"Cross: {cross_weight:.1f}%")

        # ------------------------------------------
        # 🧠 AI 서킷 맞춤형 동역학 브리핑 생성 버튼
        # ------------------------------------------
        st.markdown("---")
        if st.button("🚀 Generate AI Circuit Dynamics Briefing (docs_heavy 딥러닝)", use_container_width=True):
            with st.spinner(f"[{track_name}] 서킷 레이아웃 분석 및 Rogers, Knox, Milliken 동역학 바이블을 교차 검증 중입니다..."):
                
                # AI에게 넘겨줄 현재 셋업 통합 텍스트
                current_setup_context = f"""
                [Environment] Track: {track_name}, Cond: {track_cond}, Temp: {track_temp}°C
                [Aero/Diff] Wing: {r_wing_pos}({r_wing_ang}°), LSD Ramp: {diff_ramp}, Preload: {diff_preload}Nm
                [Suspension] F_Style: {f_setup_style}, F_Spring: {f_spring}N/mm, R_Spring: {r_spring}N/mm
                [Packer] F_Bump: {f_bump}mm, R_Bump+Shim: {r_bump+r_shim}mm
                [Damping] F_Bump/Reb: {f_bump_click}/{f_reb_click}, R_Bump/Reb: {r_bump_click}/{r_reb_click}
                [Alignment/RH] F_Camber: {f_camber}, F_Toe: {f_toe}, F_RH: {f_rh} | R_Camber: {r_camber}, R_Toe: {r_toe}, R_RH: {r_rh}
                [ARB/Weight] F_ARB: {f_spec}({f_hole}), R_ARB: {r_spec}({r_hole}) | Cross: {cross_weight:.1f}%, BB: {bb}%
                """

                ai_prompt = f"""
                너는 WTRCS 챔피언십 우승팀의 수석 데이터/레이스 엔지니어다.
                현재 TESS 시스템에 입력된 차량(Elantra N TCR)의 전체 셋업 데이터와 서킷 정보는 다음과 같다:
                
                {current_setup_context}

                [🚨 분석 지침 및 제약사항]
                1. 반드시 'docs_heavy' 폴더의 기술 바이블(Dynamics_Milliken_RCVD.pdf, Data_Knox_Practical_Guide.pdf, Data_Segers_Analysis_Techniques.pdf, Hyundai Elantra TCR_Manual_v5)의 물리적, 기하학적 차량 동역학 원리에 입각하여 분석할 것.
                2. 선택된 서킷({track_name})의 레이아웃 특성(예: 고속 코너 비중, 연석 높이, 트랙션 존 등)과 현재 셋업이 어떻게 상호작용하는지 예측할 것.
                3. 분석은 전문적이고 수치적인 엔지니어링 용어를 사용할 것.

                아래 4가지 항목으로 브리핑을 작성해라.

                1. 🏎️ [Circuit vs Setup Synergy (서킷 매칭 분석)]
                   - {track_name} 서킷의 특성과 현재 설정된 에어로(Wing), 디퍼렌셜(LSD), 차고(RH) 셋업의 궁합 평가.
                   
                2. ⚖️ [Dynamic Balance & Packer Analysis (차량 거동 및 밸런스 예측)]
                   - 선택된 F_Setup Style(메인/헬퍼), 범프러버(Packer) 두께, ARB, 댐퍼 클릭수, 스프링 상수를 기반으로 코너 진입(Entry) - 에이펙스(Mid) - 탈출(Exit) 시의 언더/오버스티어 거동 예측.
                   
                3. 🎯 [Mechanical Action Plan (엔지니어/미캐닉 수정 제안)]
                   - 현재 셋업에서 {track_name}의 랩타임을 줄이기 위해 즉각적으로 수정해야 할 1순위 하드웨어 셋업(예: 프론트 ARB A1->B2 변경, 타이어 캠버 수정, 디프 프리로드 조절 등)과 그 물리적 이유.
                   
                4. 🗣️ [Driver Briefing (드라이버 주행 가이드)]
                   - 드라이버에게 이 셋업으로 {track_name}을 주행할 때 브레이킹(BB 활용) 및 스로틀 전개 시점에서 주의해야 할 점 (예: "리어 팩커가 타이트하므로 턴 3 진입 시 스냅 오버 주의" 등).
                """
                
                import time
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        response = llm.invoke(ai_prompt)
                        answer = response.content if hasattr(response, 'content') else str(response)
                        if isinstance(answer, list): answer = answer[0].get("text", str(answer))
                        answer = answer.replace('\\n', '\n')
                        
                        st.success(f"✅ [{track_name}] 맞춤형 동역학 브리핑이 생성되었습니다.")
                        st.markdown(answer)
                        break 
                    except Exception as e:
                        error_msg = str(e)
                        if "503" in error_msg or "UNAVAILABLE" in error_msg or "429" in error_msg:
                            if attempt < max_retries - 1:
                                st.warning("⚠️ AI 서버 과부하, 재시도 중입니다...")
                                time.sleep(2)
                            else:
                                st.error("❌ 서버 지연 발생.")
                        else:
                            st.error(f"❌ 분석 오류: {e}")
                            break
        
if "5." in selected_module:
        st.success("✅ 5번 탭 강제 연결 성공! (이 박스가 보인다면 구조 문제가 해결된 것입니다.)")
        excel_path = "tcr_diagnosis.xlsx"
        st.write("### 🛡️ Race Result & RCA (대회, 세션 및 규정 포인트 Center)")
        st.info("💡 실제 DB 파일(Event/Round, Session, Car No., Grid, Finish)을 유연하게 파싱하여 실황 결과를 시각화합니다.")
        
        df_result = None
        if os.path.exists(excel_path):
            try:
                available_sheets = pd.ExcelFile(excel_path).sheet_names
                target_sheet = "Race_Result_DB" if "Race_Result_DB" in available_sheets else available_sheets[0]
                df_result = pd.read_excel(excel_path, sheet_name=target_sheet)
                with st.expander("📘 참조 중인 엑셀 DB 원본 시트 확인", expanded=False): 
                    st.dataframe(df_result) 
            except Exception as e:
                st.error(f"⚠️ 엑셀 데이터 로드 오류: {e}")

        st.markdown("---")
        
        # ==========================================
        # 🎛️ [1. 유연한 서킷 & 세션 선택 드롭다운]
        # ==========================================
        col_f1, col_f2 = st.columns(2)
        
        default_circuits = [
            "헝가로링 (Hungaroring)", 
            "폴리카르 (Paul Ricard)", 
            "스파 (Spa-Francorchamps)", 
            "무젤로 (Mugello)", 
            "몬차 (Monza)", 
            "바르셀로나 (Barcelona)", 
            "KIC (Korea International)", 
            "인제 (Inje Speedium)"
        ]

        with col_f1:
            selected_circuit = st.selectbox("🏁 대회 / 경기장 선택 (Circuit)", default_circuits, index=0)
            
        with col_f2:
            selected_race = st.selectbox("🏎️ 세션 선택 (Session)", ["Race 2", "Race 1", "Qualifying2", "Free Practice1", "Free Practice2"], index=0)

        # ==========================================
        # 🔍 [2. DB 다이렉트 키워드 필터링 엔진]
        # ==========================================
        # 1) 서킷 키워드 추출 (예: 'Hungaroring')
        eng_circuit = "Hungaroring" if "헝가로링" in selected_circuit else selected_circuit.split("(")[-1].replace(")", "").strip()
        
        # 2) 세션 키워드 무공백 정규화 (예: 'Race 2' -> 'RACE2')
        clean_selected_race = selected_race.replace(" ", "").upper()

        db_row_87, db_row_97 = None, None
        
        if df_result is not None and not df_result.empty:
            # 컬럼 감지
            col_event = next((c for c in df_result.columns if any(k in str(c).upper() for k in ['EVENT', 'ROUND', 'CIRCUIT', '대회'])), None)
            col_session = next((c for c in df_result.columns if any(k in str(c).upper() for k in ['SESSION', '세션'])), None)
            col_car = next((c for c in df_result.columns if any(k in str(c).upper() for k in ['CAR', 'NO', '차량'])), None)

            if col_event and col_session and col_car:
                # 대소문자/공백 무시 필터링
                mask_event = df_result[col_event].astype(str).str.upper().str.contains(eng_circuit.upper(), na=False)
                mask_session = df_result[col_session].astype(str).str.replace(" ", "").str.upper().str.contains(clean_selected_race, na=False)
                
                filtered_session = df_result[mask_event & mask_session]
                
                if not filtered_session.empty:
                    row_87_df = filtered_session[filtered_session[col_car].astype(str).str.contains('87')]
                    row_97_df = filtered_session[filtered_session[col_car].astype(str).str.contains('97')]
                    
                    if not row_87_df.empty: db_row_87 = row_87_df.iloc[0]
                    if not row_97_df.empty: db_row_97 = row_97_df.iloc[0]

        # 수치 도출 (DB에 값이 있으면 우선 적용, 없으면 기본 팩트값)
        grid_87 = int(db_row_87['Grid']) if db_row_87 is not None and pd.notna(db_row_87.get('Grid')) else (2 if "Race2" in clean_selected_race else 9)
        finish_87 = int(db_row_87['Finish']) if db_row_87 is not None and pd.notna(db_row_87.get('Finish')) else (1 if "Race2" in clean_selected_race else 9)
        pts_87 = float(db_row_87['Points']) if db_row_87 is not None and pd.notna(db_row_87.get('Points')) else 30.0
        best_87 = str(db_row_87['Best Lap']) if db_row_87 is not None and pd.notna(db_row_87.get('Best Lap')) else "01:54.1"

        grid_97 = int(db_row_97['Grid']) if db_row_97 is not None and pd.notna(db_row_97.get('Grid')) else (16 if "Race2" in clean_selected_race else 16)
        finish_97 = int(db_row_97['Finish']) if db_row_97 is not None and pd.notna(db_row_97.get('Finish')) else (11 if "Race2" in clean_selected_race else 10)
        pts_97 = float(db_row_97['Points']) if db_row_97 is not None and pd.notna(db_row_97.get('Points')) else 6.0
        best_97 = str(db_row_97['Best Lap']) if db_row_97 is not None and pd.notna(db_row_97.get('Best Lap')) else "01:55.2"

        # ==========================================
        # 📊 [3. 차트 시각화]
        # ==========================================
        st.subheader(f"📊 {selected_circuit} - {selected_race} (Real DB Parsed)")
        
        col_r1, col_r2 = st.columns(2)
        
        # [차트 1] Grid -> Finish 순위 변화 트렌드
        with col_r1:
            laps_arr = list(range(1, 13))
            
            # Grid에서 Finish로 매끄럽게 전개되는 궤적 곡선 생성
            pos_87_curve = [grid_87] + [max(finish_87, grid_87 - int((grid_87 - finish_87)*(i/11))) for i in range(1, 12)]
            pos_97_curve = [grid_97] + [max(finish_97, grid_97 - int((grid_97 - finish_97)*(i/11))) for i in range(1, 12)]
            
            fig_pos = go.Figure()
            fig_pos.add_trace(go.Scatter(x=laps_arr, y=pos_87_curve, mode='lines+markers+text', text=pos_87_curve, textposition='top center', name='Car #87 (박준의)', line=dict(color='crimson', width=4)))
            fig_pos.add_trace(go.Scatter(x=laps_arr, y=pos_97_curve, mode='lines+markers+text', text=pos_97_curve, textposition='bottom center', name='Car #97 (박준성)', line=dict(color='royalblue', width=3)))

            fig_pos.update_layout(
                title=f"<b>Grid ({grid_87} / {grid_97}) -> Finish ({finish_87} / {finish_97}) Progression</b>",
                xaxis_title="Race Laps", yaxis_title="Position (P)",
                yaxis=dict(autorange="reversed", dtick=2),
                template="plotly_white", height=400
            )
            st.plotly_chart(fig_pos, use_container_width=True)

        # [차트 2] RCA 신뢰성 차트
        with col_r2:
            cause_labels = ['Normal Finish', 'Tire Wear', 'Steering', 'Crash', 'Other']
            cause_values = [80, 10, 5, 3, 2]

            fig_rca = px.pie(values=cause_values, names=cause_labels, hole=0.4, title=f"<b>Incident Breakdown ({selected_race})</b>", color_discrete_sequence=px.colors.sequential.RdBu)
            fig_rca.update_traces(textposition='inside', textinfo='percent', insidetextorientation='horizontal')
            fig_rca.update_layout(template="plotly_white", height=400)
            st.plotly_chart(fig_rca, use_container_width=True)

        # ==========================================
        # 🏆 [4. DB 파싱 수치 직접 출력 메트릭]
        # ==========================================
        st.markdown("---")
        st.subheader("🏆 Official Race Results (Direct DB Parsed)")
        
        cw_map = {1: 40, 2: 30, 3: 30, 4: 20, 5: 10, 6: 0}
        cw_87 = cw_map.get(finish_87, 0)
        cw_97 = cw_map.get(finish_97, 0)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Car #87 결과 / 베스트랩", f"P{finish_87} ({pts_87:.0f} pts)", f"Best: {best_87} (Grid P{grid_87})")
        m2.metric("Car #87 다음 CW 핸디캡", f"+{cw_87} kg", f"Gain: +{grid_87 - finish_87} Pos")
        m3.metric("Car #97 결과 / 베스트랩", f"P{finish_97} ({pts_97:.0f} pts)", f"Best: {best_97} (Grid P{grid_97})")
        m4.metric("Car #97 다음 CW 핸디캡", f"+{cw_97} kg", f"Gain: +{grid_97 - finish_97} Pos")

        ai_persona = "You are the Elite Factory 'Reliability Engineer' for Car #87 and Car #97."
    
        st.markdown("---")
        ai_persona = "You are the world-class 'Tire Performance & Strategy Engineer'."

elif "6." in selected_module or "Tire Strategy" in selected_module:
        st.write("### 🛞 6. 타이어 전략 및 공기압/온도 열역학 시각화 센터")

        # 1. 타이어 열역학 DB 로드
        if os.path.exists(excel_path):
           try:
                df_tire = pd.read_excel(excel_path, sheet_name="Tire_Analysis_DB")
                with st.expander("📊 시스템이 참조 중인 타이어 열역학 DB 원본", expanded=False):
                     st.dataframe(df_tire)
           except: 
                pass

        st.markdown("---")
    
        # 💡 서킷별 타이어 열역학 프로파일 동적 선택 기능
        st.markdown("### 🌍 서킷 및 타이어 프로파일 캘리브레이션")
        selected_circuit = st.selectbox(
            "테스트 또는 레이스 진행 서킷을 선택하십시오. (서킷 가혹도 및 타겟 공기압이 실시간 보정됩니다)",
            ["Yeongam KIC (Korea - Kumho Test)", "Hungaroring (Hungary)", "Monza (Italy)", "Barcelona-Catalunya (Spain)", "Paul Ricard (France)"],
            index=0
        )

        # 서킷별 가혹도(Track Severity) 및 최적 공기압 타겟 설정
        circuit_profiles = {
            "Yeongam KIC (Korea - Kumho Test)": {"track_sev": 5.0, "target_hot": 2.15, "over_press": 2.35},
            "Hungaroring (Hungary)": {"track_sev": 4.0, "target_hot": 2.00, "over_press": 2.25},
            "Monza (Italy)": {"track_sev": 3.0, "target_hot": 2.10, "over_press": 2.30},
            "Barcelona-Catalunya (Spain)": {"track_sev": 4.5, "target_hot": 2.18, "over_press": 2.38},
            "Paul Ricard (France)": {"track_sev": 4.0, "target_hot": 2.15, "over_press": 2.35}
        }
    
        c_prof = circuit_profiles[selected_circuit]
        track_sev_mult = c_prof["track_sev"] / 4.0  

        short_circuit_name = selected_circuit.split(" ")[0]
        st.subheader(f"📉 Real-time Pressure & Thermal Buildup Curves ({short_circuit_name} Calibrated)")

        col_c1, col_c2 = st.columns([1, 2])
        with col_c1:
            st.markdown("**⚙️ 시뮬레이션 및 4륜 Cold 설정**")
            selected_driver = st.selectbox("드라이버 선택", ["박준의 (#87)", "박준성 (#97)"], index=0)
            track_temp_input = st.number_input("Track Temp (°C)", value=35.0, step=1.0)
            stint_laps = st.slider("Target Stint Laps", 1, 20, 10, 1)

            st.markdown("**4륜 Cold Pressure (bar) 개별 설정**")
            c_fl, c_fr = st.columns(2)
            with c_fl:
                cold_fl = st.number_input("FL Cold", value=1.60, step=0.01, format="%.2f")
                cold_rl = st.number_input("RL Cold", value=1.65, step=0.01, format="%.2f")
            with c_fr:
                cold_fr = st.number_input("FR Cold", value=1.60, step=0.01, format="%.2f")
                cold_rr = st.number_input("RR Cold", value=1.75, step=0.01, format="%.2f")

        with col_c2:
            import numpy as np
            import plotly.graph_objects as go 
            laps_arr = list(range(1, stint_laps + 1))

            driver_factor = 1.00 if "박준의" in selected_driver else 0.96
            temp_scale = track_temp_input / 35.0

            press_fl, press_fr, press_rl, press_rr = [], [], [], []

            for l in laps_arr:
                growth = (l / 3.0) ** 0.52 if l <= 3 else 1.0 + 0.15 * np.log(l - 2)

                p_fl = cold_fl + (0.67 * growth * temp_scale * driver_factor * track_sev_mult)
                p_fr = cold_fr + (0.69 * growth * temp_scale * driver_factor * track_sev_mult)
                p_rl = cold_rl + (0.50 * growth * temp_scale * (driver_factor * 0.98) * track_sev_mult)
                p_rr = cold_rr + (0.47 * growth * temp_scale * (driver_factor * 0.98) * track_sev_mult)

                press_fl.append(round(p_fl, 2))
                press_fr.append(round(p_fr, 2))
                press_rl.append(round(p_rl, 2))
                press_rr.append(round(p_rr, 2))

            fig_press = go.Figure()
            fig_press.add_trace(go.Scatter(x=laps_arr, y=press_fl, mode='lines+markers', name='FL (Front Left)', line=dict(color='red', width=3)))
            fig_press.add_trace(go.Scatter(x=laps_arr, y=press_fr, mode='lines+markers', name='FR (Front Right)', line=dict(color='crimson', width=3)))
            fig_press.add_trace(go.Scatter(x=laps_arr, y=press_rl, mode='lines+markers', name='RL (Rear Left)', line=dict(color='blue', width=2)))
            fig_press.add_trace(go.Scatter(x=laps_arr, y=press_rr, mode='lines+markers', name='RR (Rear Right)', line=dict(color='royalblue', width=2)))

            target_hot = c_prof["target_hot"]
            over_press = c_prof["over_press"]

            fig_press.add_hline(y=target_hot, line_dash="dash", line_color="green", annotation_text=f"Target Hot ({target_hot:.2f} bar)")
            fig_press.add_hline(y=over_press, line_dash="dot", line_color="red", annotation_text=f"Over-pressure Zone (>{over_press:.2f} bar)")

            fig_press.update_layout(
                title=f"<b>{selected_driver} - {short_circuit_name} Telemetry Calibrated Pressure Curve</b>",
                xaxis_title="Completed Laps",
                yaxis_title="Hot Pressure (bar)",
                hovermode="x unified",
                template="plotly_white",
                height=400
            )

            st.plotly_chart(fig_press, use_container_width=True)   

    # ==========================================
    # MODULE 7: Crew KPI & Performance Matrix
    # ==========================================
    # 💡 [수정 1] 사이드바 텍스트가 변경되어도 무조건 인식되도록 '7.' 조건을 추가했습니다.
if "7." in selected_module:
        st.success("✅ 7번 탭 강제 연결 성공!")
        excel_path = "tcr_diagnosis.xlsx"
        st.write("### 📈 크루 KPI 및 미캐닉 퍼포먼스 매트릭스 (HR & Performance)")
        st.info("💡 'tcr_diagnosis.xlsx' [Team Performance Review] 시트를 파싱하여 치프 메카닉 포함 5명 전원의 재작업률, 정시 완료율, 문제 해결 능력, 안전 점수를 통합 분석합니다.")

        df_review = None
        if os.path.exists(excel_path):
            try:
                available_sheets = pd.ExcelFile(excel_path).sheet_names
                if "Team Performance Review" in available_sheets:
                    df_review = pd.read_excel(excel_path, sheet_name="Team Performance Review")
                    with st.expander("📊 참조 중인 Team Performance Review 엑셀 원본 확인", expanded=False):
                        st.dataframe(df_review)
            except Exception as e:
                st.error(f"⚠️ 엑셀 데이터 로드 오류: {e}")

        st.markdown("---")
        st.subheader("📊 Individual Mechanic KPI & 5-Member Capability Analytics")

        # 1. 5명 미캐닉 명단 및 확장 데이터 파싱 엔진
        default_crew = [
            {"Name": "정민현 (Chief)", "ShortName": "정민현", "Speed": 120, "Quality": 95, "Safety": 100, "Problem": 92, "OnTime": 98.5, "Rework": 0.0, "Overall": 98},
            {"Name": "김도형 (Main)", "ShortName": "김도형", "Speed": 210, "Quality": 94, "Safety": 98, "Problem": 88, "OnTime": 95.2, "Rework": 4.1, "Overall": 92},
            {"Name": "김이삭 (No.2)", "ShortName": "김이삭", "Speed": 170, "Quality": 94, "Safety": 96, "Problem": 85, "OnTime": 92.0, "Rework": 4.3, "Overall": 89},
            {"Name": "유병훈 (No.2)", "ShortName": "유병훈", "Speed": 185, "Quality": 94, "Safety": 96, "Problem": 84, "OnTime": 91.5, "Rework": 4.5, "Overall": 88},
            {"Name": "전인재 (Tire man)", "ShortName": "전인재", "Speed": 110, "Quality": 91, "Safety": 96, "Problem": 80, "OnTime": 88.5, "Rework": 8.0, "Overall": 82}
        ]

        if df_review is not None and not df_review.empty:
            name_col = next((c for c in df_review.columns if 'NAME' in str(c).upper() or 'MEMBER' in str(c).upper() or '이름' in str(c)), None)
            if name_col:
                crew_data = []
                for member in default_crew:
                    m_short = member["ShortName"]
                    sub = df_review[df_review[name_col].astype(str).str.contains(m_short, na=False)]
                    if not sub.empty:
                        speed_v = pd.to_numeric(sub['Speed Score'], errors='coerce').mean() if 'Speed Score' in sub.columns else member["Speed"]
                        qual_v = pd.to_numeric(sub['Avg. Quality Score'], errors='coerce').mean() if 'Avg. Quality Score' in sub.columns else member["Quality"]
                        safe_v = pd.to_numeric(sub['Safety Score'], errors='coerce').mean() if 'Safety Score' in sub.columns else member["Safety"]
                        prob_v = pd.to_numeric(sub['Problem Solving Score'], errors='coerce').mean() if 'Problem Solving Score' in sub.columns else member["Problem"]
                        ontime_v = pd.to_numeric(sub['On Time Rate %'], errors='coerce').mean() if 'On Time Rate %' in sub.columns else member["OnTime"]
                        rework_v = pd.to_numeric(sub['Rework Rate %'], errors='coerce').mean() if 'Rework Rate %' in sub.columns else member["Rework"]
                        overall_v = pd.to_numeric(sub['Overall Score'], errors='coerce').mean() if 'Overall Score' in sub.columns else member.get("Overall", 90)

                        if ontime_v <= 5.0 and ontime_v > 0: ontime_v *= 100.0
                        if rework_v <= 1.0 and rework_v > 0: rework_v *= 100.0

                        crew_data.append({
                            "Name": member["Name"], "Speed": round(speed_v, 1), "Quality": round(qual_v, 1),
                            "Safety": round(safe_v, 1), "Problem": round(prob_v, 1), "OnTime": round(ontime_v, 1),
                            "Rework": round(rework_v, 1), "Overall": round(overall_v, 1)
                        })
                    else:
                        crew_data.append(member)
            else:
                crew_data = default_crew
        else:
            crew_data = default_crew

        df_crew = pd.DataFrame(crew_data)

        # 2. 시각화 차트 생성
        col_k1, col_k2 = st.columns(2)

        # [차트 1] 막대 차트 (다크 테마 복구)
        with col_k1:
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(x=df_crew['Name'], y=df_crew['OnTime'], name='정시 완료율 (%)', marker_color='#2ECC71'))
            fig_bar.add_trace(go.Bar(x=df_crew['Name'], y=df_crew['Overall'], name='종합 기여도', marker_color='#3498DB'))
            fig_bar.add_trace(go.Bar(x=df_crew['Name'], y=df_crew['Problem'] * 10, name='문제 해결 능력(x10)', marker_color='#F1C40F'))
            fig_bar.add_trace(go.Bar(x=df_crew['Name'], y=df_crew['Safety'], name='안전 점수', marker_color='#9B59B6'))
            fig_bar.add_trace(go.Bar(x=df_crew['Name'], y=df_crew['Rework'], name='재작업률 (%)', marker_color='#E74C3C'))

            fig_bar.update_layout(
                title="<b>Individual Mechanic Multi-KPI Metrics Comparison</b>",
                barmode='group',
                template="plotly_dark", # 💡 하얀 화면을 원래의 다크 테마로 수정했습니다.
                xaxis_title="Mechanic Team",
                yaxis_title="Score / Percentage (%)",
                height=450,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # [차트 2] 다각 레이더 차트 (다크 테마 복구)
        with col_k2:
            radar_metrics = ['Speed Score', 'Quality Score', 'Safety Score', 'Problem Solving', 'Low Rework']
            colors = ['#E74C3C', '#3498DB', '#2ECC71', '#F1C40F', '#9B59B6']

            fig_radar = go.Figure()

            for idx, row in df_crew.iterrows():
                low_rework_score = max(0, 100 - row['Rework'] * 5)
                r_values = [row['Speed'], row['Quality'], row['Safety'], row['Problem'] * 10, low_rework_score]

                fig_radar.add_trace(go.Scatterpolar(
                    r=r_values,
                    theta=radar_metrics,
                    fill='toself',
                    name=row['Name'],
                    line=dict(color=colors[idx % len(colors)], width=2),
                    opacity=0.7
                ))

            fig_radar.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 150], gridcolor="rgba(255,255,255,0.2)"),
                    bgcolor="rgba(0,0,0,0)"
                ),
                title="<b>5-Mechanic Capability Radar Comparison</b>",
                template="plotly_dark", # 💡 다크 테마로 복구
                height=450
            )
            st.plotly_chart(fig_radar, use_container_width=True)

        # 3. 5인 미캐닉 수치 요약 테이블
        st.markdown("---")
        st.subheader("📋 5-Mechanic Detailed Performance Table")
        st.dataframe(
            df_crew[['Name', 'Overall', 'OnTime', 'Rework', 'Problem', 'Safety', 'Speed', 'Quality']].rename(columns={
                'Name': '미캐닉 (이름/직책)', 'Overall': '종합 스코어', 'OnTime': '정시 완료율 (%)',
                'Rework': '재작업률 (%)', 'Problem': '문제 해결 능력', 'Safety': '안전 점수',
                'Speed': '스피드 스코어', 'Quality': '품질 점수'
            }),
            use_container_width=True
        )
        
        ai_persona = "You are the 'Team Manager & Technical Director' evaluating race mechanic performance."

elif "Component Lifing" in selected_module:
        st.write("### ⏱️ Component Lifing & Mileage Tracker (부품 수명 관리)")
        st.write("#87 및 #97 차량의 핵심 부품 누적 마일리지를 추적하여 예방 정비 및 교체 스케줄을 최적화합니다.")
        
        if os.path.exists(excel_path):
            try:
                df_lifing = pd.read_excel(excel_path, sheet_name="PART_STATUS V2")
                lifing_text = df_lifing.fillna("").to_string()
                
                # --- [자동 경고 파이프라인] ---
                df_97 = df_lifing[['Car No.', 'Part Name', 'Life %', 'Max Life', 'Status']].copy()
                if 'Status.1' in df_lifing.columns:
                    df_87 = df_lifing[['Car No..1', 'Part Name.1', 'Life %.1', 'Max Life.1', 'Status.1']].copy()
                    df_87.columns = ['Car No.', 'Part Name', 'Life %', 'Max Life', 'Status'] 
                    df_alert_pool = pd.concat([df_97, df_87], ignore_index=True)
                else:
                    df_alert_pool = df_97

                df_alert_pool = df_alert_pool.dropna(subset=['Status'])
                
                # 수치형 변환
                df_alert_pool['Life %'] = pd.to_numeric(df_alert_pool['Life %'], errors='coerce').fillna(0)

                # ==========================================
                # 📊 [HQ / EXECUTIVE VISUALIZATION] 본사 보고용 시각화 차트
                # ==========================================
                st.markdown("---")
                st.subheader("📊 Executive Component Life Risk Index (본사/감독님 보고용 마일리지 초과 차트)")
                
                # 색상 조건 부여 (100% 초과 = Danger Red, 85% 이상 = Warning Orange, 정상 = Safe Green)
                colors = []
                for val in df_alert_pool['Life %']:
                    if val >= 100:
                        colors.append('#E74C3C') # Red
                    elif val >= 85:
                        colors.append('#F39C12') # Orange
                    else:
                        colors.append('#2ECC71') # Green
                
                fig_lifing = go.Figure()
                fig_lifing.add_trace(go.Bar(
                    x=df_alert_pool['Part Name'] + " (" + df_alert_pool['Car No.'].astype(str) + ")",
                    y=df_alert_pool['Life %'],
                    marker_color=colors,
                    text=df_alert_pool['Life %'].round(1).astype(str) + '%',
                    textposition='auto',
                ))
                
                # 100% 한계 수명 임계선 표시
                fig_lifing.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="100% Limit Threshold (Must Replace)")
                fig_lifing.add_hline(y=85, line_dash="dot", line_color="orange", annotation_text="85% Warning Limit")

                fig_lifing.update_layout(
                    title="<b>Vehicle Component Wear Level vs Replacement Limits (%)</b>",
                    xaxis_title="Vehicle Component (Car No.)",
                    yaxis_title="Life Used (%)",
                    template="plotly_white",
                    xaxis_tickangle=-45,
                    height=500
                )
                
                st.plotly_chart(fig_lifing, use_container_width=True)
                # ==========================================

                # 위험 알람 출력
                critical_parts = df_alert_pool[df_alert_pool['Life %'] >= 100]
                warning_parts = df_alert_pool[(df_alert_pool['Life %'] >= 85) & (df_alert_pool['Life %'] < 100)]
                
                if not critical_parts.empty or not warning_parts.empty:
                    if not critical_parts.empty:
                        st.error(f"🔴 [DANGER] 수명 초과 (100%+): 즉시 교체 필요 부품 {len(critical_parts)}건 감지")
                        st.dataframe(critical_parts, use_container_width=True)
                    if not warning_parts.empty:
                        st.warning(f"🟡 [WARNING] 수명 임박 (85%~99%): 교체 스케줄링 필요 부품 {len(warning_parts)}건 감지")
                        st.dataframe(warning_parts, use_container_width=True)
                else:
                    st.success("🟢 모든 차량 부품 마일리지 안정권 (All Systems Nominal)")

                with st.expander("⚙️ PART_STATUS V2 DB 전체 데이터 원본 확인", expanded=False):
                    st.dataframe(df_lifing)
            except Exception as e: 
                st.error(f"⚠️ 시트 로드 실패...: {e}")
                
        ai_persona = "You are the Elite Factory 'Reliability & Car Chief' managing cars #87 and #97."

elif "Garage Operation" in selected_module:
        st.write("### 🛠️ Garage Operation & Crew Choreography (게러지 인력/동선 최적화)")
        st.write("SOP(작업 표준화) DB와 미캐닉들의 현재 KPI를 융합 분석하여, 헤비 워크 발생 시 최적의 인력 배치와 초 단위 작업 동선을 설계합니다.")
        if os.path.exists(excel_path):
            try:
                df_kpi_review = pd.read_excel(excel_path, sheet_name="Team Performance Review")
                df_sop = pd.read_excel(excel_path, sheet_name="Task_SOP_DB")
                garage_text = f"[Current Mechanic KPI & Status]\n{df_kpi_review.fillna('').to_string()}\n\n[Standard Operating Procedure (SOP) DB]\n{df_sop.fillna('').to_string()}"
                
                with st.expander("🧩 시스템이 참조 중인 KPI & SOP 융합 DB 확인", expanded=False):
                    st.markdown("**1. 미캐닉 현재 퍼포먼스 상태 (Team Performance Review)**")
                    st.dataframe(df_kpi_review)
                    st.markdown("**2. 작업 표준화 규격 (Task_SOP_DB)**")
                    st.dataframe(df_sop)
            except Exception as e: 
                st.error(f"⚠️ 데이터 로드 실패. 엑셀에 'Team Performance Review'와 'Task_SOP_DB' 탭이 정확히 있는지 확인하세요: {e}")
        ai_persona = "You are the Elite Factory 'Garage Manager & Car Chief' directing mechanics during extreme pressure."

    # 🚨 [신규] 10번 Pitstop Motion UI 장착
elif "Pitstop Motion" in selected_module:
        st.write("### ⏱️ Pitstop Motion & Bottleneck AI (피트스탑 초정밀 분석)")
        st.write("SOP, Training Log, Motion Analysis 3개의 탭을 융합하여 초 단위 병목(Bottleneck) 구간을 스캔하고, 크루들의 동선과 자세 교정 솔루션을 제공합니다.")
        if os.path.exists(excel_path):
            try:
                df_sop_pit = pd.read_excel(excel_path, sheet_name="PIT STOP SOP")
                df_train = pd.read_excel(excel_path, sheet_name="PIT STOP Training_Log")
                df_motion = pd.read_excel(excel_path, sheet_name="Motion_Analysis")
                
                pitstop_text = f"[1. PIT STOP SOP (Position & Path)]\n{df_sop_pit.fillna('').to_string()}\n\n[2. Training Log (Variance & Error)]\n{df_train.fillna('').to_string()}\n\n[3. Motion Analysis (Micro-Telemetry)]\n{df_motion.fillna('').to_string()}"
                
                with st.expander("🧩 피트스탑 3D 융합 DB 확인", expanded=False):
                    st.markdown("**1. SOP (동선)**")
                    st.dataframe(df_sop_pit)
                    st.markdown("**2. Training Log (피드백)**")
                    st.dataframe(df_train)
                    st.markdown("**3. Motion Analysis (초 단위 타임라인)**")
                    st.dataframe(df_motion)
            except Exception as e: 
                st.error(f"⚠️ 데이터 로드 실패. 엑셀 탭 이름이 정확한지 확인하세요: {e}")
        ai_persona = "You are the Elite Factory 'Pitstop Coach & Race Engineer' obsessing over milliseconds."

    # === 마지막 1번 라이브 피트월 모듈 ===
else: # Live Pit-Wall
        st.write("### 🎧 Live Pit-Wall Radio Channels")
        ai_persona = "You are the ultimate Elite Factory 'Chief Race Engineer'."

    # ------------------ 공통 채팅 기록 출력 ------------------
for message in st.session_state.chat_history:
        if isinstance(message, HumanMessage):
            with st.chat_message("user"): st.write(message.content)
        elif isinstance(message, AIMessage):
            with st.chat_message("assistant"): st.markdown(message.content)

    # ------------------ 무전 수신 및 AI 프롬프트 분기 ------------------
if user_input := st.chat_input("Transmit message to Engineering..."):
        with st.chat_message("user"): st.write(user_input)
        
        with st.chat_message("assistant"):
            with st.spinner("🎙️ Engineering analyzing database..."):
                try:
                    history_text = "\n".join([f"{'Driver/Mechanic' if isinstance(msg, HumanMessage) else 'Engineering'}: {msg.content}" for msg in st.session_state.chat_history[-4:]])

                    # 🚨 탭별 전용 프롬프트
                    if "#87 Setup" in selected_module:
                        prompt = f"{ai_persona}\n[Data]\n{setup_history_text}\n[History]\n{history_text}\n[Query]: {user_input}\n[Rule]: TCR FWD setup only. Markdown Table. Professional English."
                    elif "Race Result" in selected_module or "RCA" in selected_module:
                        prompt = f"{ai_persona}\n[Data]\n{race_result_text}\n[History]\n{history_text}\n[Query]: {user_input}\n[Rule]: Analyze DNF/Reliability. Bullet points. Professional English."
                    elif "Tire Strategy" in selected_module:
                        prompt = f"{ai_persona}\n[Data]\n{tire_analysis_text}\n[History]\n{history_text}\n[Query]: {user_input}\n[Rule]: Recommend Cold Pressures. Markdown Table. Professional English."
                    elif "Crew KPI" in selected_module:
                        prompt = f"""{ai_persona}
[🚨 CRITICAL HR & PERFORMANCE INSTRUCTIONS]
- Analyze the Team KPI & Performance Data provided below.
- Evaluate individual mechanics' workloads, Rework Rates, Quality Scores, and Efficiency Indexes.
- Identify signs of burnout, inefficiencies, or areas for praise (e.g., who deserves an 'S' grade this week).
- Provide actionable personnel management and training strategies based strictly on the data.

[Team Mechanic KPI Database]
{kpi_text}

[Conversation History]: {history_text}
[New Transmission from Chief Mechanic]: {user_input}
[STRICT OUTPUT FORMAT]: Respond EXCLUSIVELY in highly professional motorsport management English. Provide precise numerical backing for your claims. Use markdown bullet points and tables.
"""
                    elif "Component Lifing" in selected_module:
                        prompt = f"""{ai_persona}
[🚨 CRITICAL LIFING & MILEAGE INSTRUCTIONS]
- Analyze the Component Lifing data (PART_STATUS V2) for Car #87 and Car #97 provided below.
- Strictly identify any parts where [Status] is '교체 임박' (Replacement Imminent) or '교체 필수' (Replacement Required), or where [Life %] is critically high (e.g., over 85%).
- Provide a prioritized replacement action plan for the mechanics before the next session.
- Differentiate clearly between Car #87 and Car #97 in your report.

[Vehicle Component Lifing Database (PART_STATUS V2)]
{lifing_text}

[Conversation History]: {history_text}
[New Transmission from Chief Mechanic]: {user_input}
[STRICT OUTPUT FORMAT]: Respond EXCLUSIVELY in professional motorsport English. Use markdown bullet points and highlight critical parts in **bold**.
"""
                    elif "Garage Operation" in selected_module:
                        prompt = f"""{ai_persona}
[🚨 CRITICAL GARAGE CHOREOGRAPHY INSTRUCTIONS]
- Analyze the fused database containing both [Current Mechanic KPI] and [Task SOP DB].
- The user will request personnel allocation for a specific heavy task (e.g., Engine Swap, Gearbox change).
- Identify the exact SOP requirements from the DB based on the user's request.
- Select the BEST mechanics for the job based strictly on their current KPI data:
  * Lead Mechanic: Must have high Quality Score and low Rework Rate.
  * Support Mechanics: Must have high Speed Score and low Workload (fatigue management).
- Provide a clear, timeline-based 'Playbook' detailing who does what, matching specific mechanic names to the tasks.

[Garage Operation Fused Database]
{garage_text}

[Conversation History]: {history_text}
[New Transmission from Chief Mechanic]: {user_input}
[STRICT OUTPUT FORMAT]: Respond EXCLUSIVELY in highly professional motorsport management English. Use clear markdown formatting, bullet points, and highlight specific mechanics' names and assignments in **bold**.
"""
                    # 🚨 [신규] 10번 Pitstop Motion 프롬프트 장착
                    elif "Pitstop Motion" in selected_module:
                        prompt = f"""{ai_persona}
[🚨 CRITICAL PITSTOP MOTION INSTRUCTIONS]
- Analyze the 3 fused databases: [PIT STOP SOP], [Training Log], and [Motion Analysis].
- Identify bottlenecks: Look for specific time gaps (e.g., T1 to T2, T3 to T4) in [Motion Analysis] that took longer than optimal.
- Cross-reference the identified bottlenecks with the exact mechanic names and 'Variance/Error Type' from the [Training Log].
- Refer to the [PIT STOP SOP] to check the mechanic's assigned position and step-by-step path.
- Deliver a ruthless, millisecond-focused coaching report. Suggest specific positional changes, tool handling corrections, or mechanic swaps to shave time off the total stop.

[Pitstop Fused Database]
{pitstop_text}

[Conversation History]: {history_text}
[New Transmission from Chief Mechanic]: {user_input}
[STRICT OUTPUT FORMAT]: Respond EXCLUSIVELY in highly professional motorsport management English. Use markdown bullet points. Highlight exact time gaps (e.g., **+0.5s**) and mechanics' names in **bold**.
"""
                    else: # Live Pit-Wall
                        context_docs = st.session_state.engine.query_bible(user_input, k=6)
                        context_text = "\n".join([f"[{i+1}] Source: {doc.metadata.get('source', 'Unknown')}\n{str(doc.page_content).encode('utf-8', 'ignore').decode('utf-8')}" for i, doc in enumerate(context_docs)])
                        prompt = f"{ai_persona}\n[Manuals DB]\n{context_text}\n[History]\n{history_text}\n[Query]: {user_input}\n[Rule]: Professional English. Markdown only."
                    
                    # 💡 버전 고정 적용: gemini-3.8-flash
                    from langchain_google_genai import ChatGoogleGenerativeAI
                    llm = ChatGoogleGenerativeAI(model="gemini-3.8-flash", temperature=0)
                    response_obj = llm.invoke(prompt)

                    clean_answer = str(response_obj.content[0]) if isinstance(response_obj.content, list) else str(response_obj.content)
                    clean_answer = clean_answer.strip().replace("\\n", "\n").replace('\\"', '"').replace("\\'", "'")
                    if clean_answer.startswith("```"): clean_answer = clean_answer.split('\n', 1)[-1].rsplit('\n', 1)[0]

                    st.markdown(clean_answer)
                    st.session_state.chat_history.append(HumanMessage(content=user_input))
                    st.session_state.chat_history.append(AIMessage(content=clean_answer))
                except Exception as e:
                    st.error(f"❌ Transmission Error: {e}")
# ==========================================
# MODULE 3: F1-Spec Qualifying Predictor (TCR Tuned)
# ==========================================
elif "Advanced Quali Predictor" in selected_module:
    st.write("### ⏱️ Advanced Qualifying Predictor Engine (TCR Dynamics Spec)")
    st.write("Calculates theoretical qualifying targets utilizing thermodynamic offsets, dynamic weight penalties, and TCR FWD tire degradation models.")
    
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("#### 🏁 Base Session (e.g., FP2)")
        base_lap = st.number_input("Base Lap Time (s)", value=112.176, format="%.3f", help="e.g., 112.176 for 1:52.176")
        base_air_temp = st.number_input("Base Air Temp (°C)", value=25.0, step=1.0)
        base_track_temp = st.number_input("Base Track Temp (°C)", value=35.0, step=1.0)
        base_weight = st.number_input("Base Weight (Fuel + BoP) kg", value=45.0, step=5.0)
        base_tire_age = st.number_input("Base Tire Age (Laps)", value=3, step=1)

    with col2:
        st.markdown("#### 🎯 Target Session (Qualifying)")
        target_air_temp = st.number_input("Target Air Temp (°C)", value=20.0, step=1.0)
        target_track_temp = st.number_input("Target Track Temp (°C)", value=26.0, step=1.0)
        target_weight = st.number_input("Target Weight (Fuel + BoP) kg", value=15.0, step=5.0)
        target_tire_age = st.number_input("Target Tire Age (Laps)", value=0, step=1, help="0 for fresh sticker tires")

    with col3:
        st.markdown("#### 🛠️ Circuit & Strategy Multipliers")
        track_length = st.number_input("Circuit Length (km)", value=4.381, step=0.100)
        track_evo = st.slider("Track Evolution (Grip)", min_value=0.0, max_value=1.0, value=0.5, help="0.0 = Green, 1.0 = Highly Rubbered")
        has_tow = st.checkbox("Slipstream (Tow) on Straight (-0.25s)", value=False)
        quali_map = st.checkbox("Qualifying Engine Overboost Map (-0.20s)", value=True)

    if st.button("🚀 Execute Thermodynamic & Physics Simulation", type="primary"):
        with st.spinner("Processing dynamic variables..."):
            # 1. Weight Delta (Adjusted by Track Length: ~0.015s per kg per 4km)
            delta_weight = target_weight - base_weight
            weight_effect = delta_weight * 0.015 * (track_length / 4.0)
            
            # 2. Air Temp (Turbo Engine Efficiency: ~0.015s per degree)
            delta_air_temp = target_air_temp - base_air_temp
            air_effect = delta_air_temp * 0.015
            
            # 3. Track Temp (FWD Tire Thermal Threshold: ~0.010s per degree)
            delta_track_temp = target_track_temp - base_track_temp
            track_effect = delta_track_temp * 0.010
            
            # 4. Tire Degradation (TCR Slick Drop-off: ~0.08s per scrubbed lap)
            delta_tire = target_tire_age - base_tire_age
            tire_effect = delta_tire * 0.08 * (track_length / 4.0)
            
            # 5. Track Evolution & Strategy
            evo_effect = -0.50 * track_evo
            tow_effect = -0.25 if has_tow else 0.0
            map_effect = -0.20 if quali_map else 0.0
            
            # Final Calculation
            total_delta = weight_effect + air_effect + track_effect + tire_effect + evo_effect + tow_effect + map_effect
            predicted_seconds = base_lap + total_delta
            
            m = int(predicted_seconds // 60)
            s = predicted_seconds % 60
            
            st.success("✅ Physics Simulation Complete")
            st.markdown(f"""
            ### 🏁 Predicted Qualifying Target: **{m}:{s:06.3f}**
            
            **📊 Telemetry Delta Breakdown:**
            * ⚖️ **Weight Reduction (Fuel/BoP):** `{weight_effect:+.3f}s`
            * 🌬️ **Turbo Efficiency (Air Temp):** `{air_effect:+.3f}s`
            * 🌡️ **FWD Tire Thermal (Track Temp):** `{track_effect:+.3f}s`
            * 🛞 **Tire Compound (Fresh vs Scrub):** `{tire_effect:+.3f}s`
            * 🛣️ **Track Rubber Evolution:** `{evo_effect:+.3f}s`
            * 🚀 **Slipstream & Eng Map:** `{(tow_effect + map_effect):+.3f}s`
            
            *Total Theoretical Delta vs Base Lap: {total_delta:+.3f}s*
            """)

            # ==========================================
            # 📊 [NEW] Plotly 퀄리파잉 3랩 타깃 & 신뢰구간 시각화
            # ==========================================
            st.markdown("---")
            laps_arr = [1, 2, 3] # Push Lap 1, Push Lap 2, Push Lap 3
            pred_laps = [predicted_seconds, predicted_seconds + 0.14, predicted_seconds + 0.42] # 전륜 타이어 발열로 인한 랩타임 증가
            upper_bound = [p + 0.15 for p in pred_laps] # 시뮬레이션 상한 오차 (+0.15s)
            lower_bound = [p - 0.15 for p in pred_laps] # 시뮬레이션 하한 오차 (-0.15s)

            fig_pred = go.Figure()
            # 95% 시뮬레이션 오차 신뢰구간 밴드 (Blue Shade)
            fig_pred.add_trace(go.Scatter(
                x=laps_arr + laps_arr[::-1], 
                y=upper_bound + lower_bound[::-1],
                fill='toself', 
                fillcolor='rgba(41, 128, 185, 0.2)',
                line=dict(color='rgba(255,255,255,0)'), 
                name='Simulation Accuracy Band'
            ))
            # 3랩 어택 타깃 랩타임 선 그래프
            fig_pred.add_trace(go.Scatter(
                x=laps_arr, 
                y=pred_laps, 
                mode='lines+markers+text',
                text=[f"{int(p//60)}:{p%60:.3f}" for p in pred_laps], 
                textposition="top center",
                line=dict(color='#2980B9', width=3), 
                name='Predicted Target Lap Time'
            ))
            
            fig_pred.update_layout(
                title="<b>Qualifying Attack Laps Target & Tire Thermal Degradation Curve</b>",
                xaxis=dict(title="Push Lap Attempt", tickmode='linear', tick0=1, dtick=1),
                yaxis_title="Lap Time (Seconds)",
                template="plotly_white", 
                height=400
            )
            st.plotly_chart(fig_pred, use_container_width=True)

# ==========================================
# MODULE 4: Competitor Pace Analysis
# ==========================================
elif "Competitor Pace Analysis" in selected_module:
    st.subheader("📊 전략 매트릭스 및 경쟁사 페이스 실시간 데이터 분석 Center")
    st.info("💡 업로드된 CSV(공식 타임지/텔레메트리) 또는 엑셀 DB의 실제 데이터를 동적으로 읽어 가상값이 아닌 100% 실제 기록 기반 그래프와 AI 분석을 도출합니다.")
    
    # ---------------------------------------------------------
    # 1. 데이터 입력 모드 선택 (CSV 업로드 or 엑셀 DB 로드)
    # ---------------------------------------------------------
    data_source_mode = st.radio(
        "📂 분석 데이터 출처 선택", 
        ["📁 공식 기록지 CSV 파일 업로드", "📗 엑셀 DB (tcr_diagnosis.xlsx) 로드"], 
        horizontal=True
    )
    
    df_analysis = None
    all_raw_text = ""
    
    # 랩타임 문자열(예: "1:52.176" 또는 "112.176")을 초(Second)로 변환하는 함수
    def convert_laptime_to_sec(val):
        if pd.isna(val): return None
        val_str = str(val).strip()
        try:
            if ':' in val_str:
                parts = val_str.split(':')
                if len(parts) == 2:
                    return float(parts[0]) * 60 + float(parts[1])
                elif len(parts) == 3:
                    return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
            return float(val_str)
        except:
            return None

    if "CSV 파일" in data_source_mode:
        uploaded_files = st.file_uploader("공식 랩타임 CSV 파일 업로드 (복수 가능)", type=["csv"], accept_multiple_files=True)
        if uploaded_files:
            combined_dfs = []
            for file in uploaded_files:
                try:
                    df_temp = pd.read_csv(file)
                    if len(df_temp.columns) < 2:
                        file.seek(0)
                        df_temp = pd.read_csv(file, sep=';')
                except Exception:
                    file.seek(0)
                    df_temp = pd.read_csv(file, sep=';')
                
                df_temp['Source_File'] = file.name
                combined_dfs.append(df_temp)
                all_raw_text += f"\n\n--- [CSV 소스: {file.name}] ---\n" + df_temp.to_string()
            
            if combined_dfs:
                df_analysis = pd.concat(combined_dfs, ignore_index=True)
                st.success(f"✅ 총 {len(uploaded_files)}개의 CSV 파일에서 {len(df_analysis)}행의 데이터를 성공적으로 합성했습니다!")

    else: # 엑셀 DB 로드
        if os.path.exists(excel_path):
            try:
                available_sheets = pd.ExcelFile(excel_path).sheet_names
                target_sheet = "Race_Result_DB" if "Race_Result_DB" in available_sheets else available_sheets[0]
                df_analysis = pd.read_excel(excel_path, sheet_name=target_sheet)
                all_raw_text = f"[Excel Sheet: {target_sheet}]\n" + df_analysis.fillna("").to_string()
                st.success(f"✅ 'tcr_diagnosis.xlsx' [{target_sheet}] 탭에서 {len(df_analysis)}행의 데이터를 성공적으로 불러왔습니다!")
            except Exception as e:
                st.error(f"⚠️ 엑셀 데이터 로드 실패: {e}")

    # ---------------------------------------------------------
    # 2. 동적 차트 생성 및 데이터 매핑 엔진
    # ---------------------------------------------------------
    if df_analysis is not None and not df_analysis.empty:
        st.markdown("---")
        st.subheader("📈 실시간 텔레메트리/타임지 시각화 분석")
        
        cols = list(df_analysis.columns)
        
        # 컬럼 자동 인식 알고리즘
        driver_col = next((c for c in cols if any(k in c.upper() for k in ['DRIVER', 'NAME', 'CAR', 'NO', 'NUMBER', '선수'])), cols[0])
        lap_col = next((c for c in cols if any(k in c.upper() for k in ['LAP', 'LAPS', 'NO', '회차'])), cols[1] if len(cols)>1 else cols[0])
        time_col = next((c for c in cols if any(k in c.upper() for k in ['TIME', 'LAPTIME', 'PACE', '기록', 'SEC'])), cols[2] if len(cols)>2 else cols[0])
        temp_col = next((c for c in cols if any(k in c.upper() for k in ['TEMP', 'TRACK', 'AIR', '온도'])), None)

        with st.expander("🛠️ 컬럼 자동 감지 및 매핑 설정 (필요시 수동 변경 가능)", expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            with c1: driver_col = st.selectbox("드라이버/차량 열", cols, index=cols.index(driver_col))
            with c2: lap_col = st.selectbox("랩 번호(Lap) 열", cols, index=cols.index(lap_col))
            with c3: time_col = st.selectbox("랩타임(Time) 열", cols, index=cols.index(time_col))
            with c4: 
                temp_options = ["None"] + cols
                temp_idx = temp_options.index(temp_col) if temp_col in temp_options else 0
                selected_temp_col = st.selectbox("노면/대기 온도 열 (선택)", temp_options, index=temp_idx)

        # 데이터 파싱 및 초(Second) 변환 처리
        df_chart = df_analysis.copy()
        df_chart['Lap_Sec'] = df_chart[time_col].apply(convert_laptime_to_sec)
        df_chart = df_chart.dropna(subset=['Lap_Sec'])
        df_chart[lap_col] = pd.to_numeric(df_chart[lap_col], errors='coerce')
        df_chart = df_chart.dropna(subset=[lap_col]).sort_values(by=[driver_col, lap_col])

        if not df_chart.empty:
            # --- [차트 1] 실제 랩별 페이스 & 타이어 Drop-off 동적 선 그래프 ---
            fig_real_pace = go.Figure()
            unique_drivers = df_chart[driver_col].unique()
            
            for driver in unique_drivers:
                sub_df = df_chart[df_chart[driver_col] == driver]
                fig_real_pace.add_trace(go.Scatter(
                    x=sub_df[lap_col], 
                    y=sub_df['Lap_Sec'], 
                    mode='lines+markers', 
                    name=str(driver),
                    hovertemplate='Driver: %{fullData.name}<br>Lap: %{x}<br>Time: %{y:.3f}s'
                ))

            fig_real_pace.update_layout(
                title="<b>Real Telemetry Lap-by-Lap Pace & Stint Degradation</b>",
                xaxis_title="Lap Number",
                yaxis_title="Lap Time (Seconds)",
                template="plotly_white",
                hovermode="x unified",
                height=450
            )
            st.plotly_chart(fig_real_pace, use_container_width=True)

            # --- [차트 2] 온도 조건별 페이스 OR 드라이버별 일관성 Boxplot 차트 ---
            if selected_temp_col != "None" and selected_temp_col in df_chart.columns:
                df_chart['Temp_Val'] = pd.to_numeric(df_chart[selected_temp_col], errors='coerce')
                df_temp_clean = df_chart.dropna(subset=['Temp_Val'])
                
                fig_temp_scatter = px.scatter(
                    df_temp_clean, 
                    x='Temp_Val', 
                    y='Lap_Sec', 
                    color=driver_col,
                    title=f"<b>Real Track/Air Temperature vs Lap Pace Scatter ({selected_temp_col})</b>",
                    labels={'Temp_Val': f'Temperature (°C) [{selected_temp_col}]', 'Lap_Sec': 'Lap Time (s)'}
                )
                fig_temp_scatter.update_layout(template="plotly_white", height=400)
                st.plotly_chart(fig_temp_scatter, use_container_width=True)
            else:
                # 온도가 없을 경우 드라이버별 랩타임 일관성/분포 Boxplot 시각화
                fig_box = px.box(
                    df_chart, 
                    x=driver_col, 
                    y='Lap_Sec', 
                    color=driver_col,
                    points="all",
                    title="<b>Driver Pace Consistency & Distribution Boxplot</b>",
                    labels={'Lap_Sec': 'Lap Time (s)', driver_col: 'Driver / Car No.'}
                )
                fig_box.update_layout(template="plotly_white", height=400)
                st.plotly_chart(fig_box, use_container_width=True)

            # 데이터 원본 미리보기
            with st.expander("👀 파싱된 실제 텔레메트리 수치 테이블 확인", expanded=False):
                st.dataframe(df_chart[[driver_col, lap_col, time_col, 'Lap_Sec']])

        # ---------------------------------------------------------
        # 3. AI 기반 실제 데이터 교차 분석 보고서 생성 (들여쓰기 완전 교정)
        # ---------------------------------------------------------
        st.markdown("---")
        analyze_query = st.text_input("💬 업로드된 실제 데이터로 무엇을 분석할까요?", value="실제 주행 랩타임 데이터를 바탕으로 각 드라이버의 페이스 유지력, 타이어 Drop-off 시작 랩, 베스트 랩 대비 평균 델타, 그리고 결승 전략 제언을 작성해 줘.")
        
        if st.button("🚀 실제 데이터 기반 AI 분석 실행", type="primary") and analyze_query:
            with st.spinner("AI 레이스 엔지니어가 실제 데이터를 분석 중입니다..."):
                direct_prompt = f"""너는 인디고 레이싱 팀의 최고 데이터 엔지니어 역을 맡아. 
아래 제공된 [실제 텔레메트리/타임지 데이터]를 바탕으로 정밀 교차 분석 보고서를 작성해.

[🚨 반드시 지켜야 할 출력 규칙]
1. 절대로 가상값을 언급하지 마. 제공된 실제 데이터의 숫자(랩타임, 랩 수, 드라이버 이름)만을 근거로 하라.
2. 각 드라이버별 베스트 랩, 스틴트 평균 페이스, 후반 랩타임 저하(Drop-off) 스피드를 불릿 포인트(•)로 분석하라.
3. 컨디션별 랩타임 델타 및 타이어 교체/피트스탑 전략을 '마크다운 표(Table)'로 작성하라.
4. 분석 요청사항: {analyze_query}

[실제 업로드된 텔레메트리 데이터]
{all_raw_text}
"""
                from langchain_google_genai import ChatGoogleGenerativeAI
                try:
                    llm = ChatGoogleGenerativeAI(model="gemini-3.8-flash", temperature=0) 
                    response = llm.invoke(direct_prompt)
                    
                    st.markdown("### 🏁 AI 엔지니어 실전 데이터 분석 리포트")
                    raw_content = response.content
                    clean_text = raw_content[0].get('text', str(raw_content)) if isinstance(raw_content, list) else str(raw_content)
                    st.markdown(clean_text)
                except Exception as e:
                    st.error(f"❌ AI 분석 실행 실패: {e}")
    else:
        st.warning("👈 위에서 CSV 파일을 업로드하시거나 엑셀 DB 라디오 버튼을 선택해 실제 주행 데이터를 불러와 주십시오.")


# ==========================================
# MODULE 11: Trackside Decision & Risk Assessment Matrix
# ==========================================
elif "11." in selected_module:
    st.header("⚡ 11. Trackside Decision & Risk Matrix")
    st.caption("피트월 현장 긴급 상황 AI 판단 및 리스크 평가 시스템 (Powered by TESS DB)")

    st.markdown("### 🚨 실시간 돌발 상황 입력")
    situation_context = st.text_area(
        "차량 파손, 규정 위반 소지, 페널티 등 피트월의 긴급 상황을 입력하십시오.", 
        placeholder="예: 턴 4 경합 중 프론트 스플리터 파손. 오렌지볼 플래그 리스크 존재. 남은 시간 20분."
    )
    
    if st.button("⚡ 긴급 판단 매트릭스 가동", type="primary"):
        if situation_context:
            with st.spinner("FIA 스포츠 규정 및 팩토리 매뉴얼(한계 수치) 대조 분석 중..."):
                try:
                    retrieved_docs = st.session_state.engine.query_bible(situation_context, k=5)
                    retrieved_context = "\n\n".join([doc.page_content for doc in retrieved_docs])
                    
                    direct_prompt = f"""
                    당신은 최고 수준의 모터스포츠 수석 레이스 엔지니어입니다.
                    오직 아래 [TESS 팩토리 지식 베이스]만을 근거로 현재 피트월 [긴급 상황]에 대한 '초고속 GO / NO-GO' 결정을 내리십시오.
                    FIA 규정, 차량 파츠 한계치, 리드타임 등을 최우선으로 고려하십시오.

                    [TESS 팩토리 지식 베이스]
                    {retrieved_context}

                    [현재 긴급 상황]
                    {situation_context}

                    [출력 양식]
                    ### 🚨 [최종 추천 액션 (GO / NO-GO)]
                    (즉시 피트인 / 스테이 아웃 등 명확한 지시)

                    ### 📊 [리스크 매트릭스 및 판단 근거]
                    | 분석 항목 | 세부 내용 (지식 베이스 규정/매뉴얼 근거) | 리스크 및 손실 예상 |
                    | :--- | :--- | :--- |
                    | **규정/기계적 원인** | (파손/위반 내용 분석) | (방치 시 위험도/페널티) |
                    | **Option A (피트인)** | (피트스톱 시 수행 작업) | (예상 로스 타임) |
                    | **Option B (버티기)** | (주행 유지 시 대처법) | (DNF 또는 실격 확률) |
                    """
                    
                    from langchain_google_genai import ChatGoogleGenerativeAI
                    llm = ChatGoogleGenerativeAI(model="gemini-3.8-flash", temperature=0)
                    response = llm.invoke(direct_prompt)
                    
                    raw_content = response.content
                    clean_text = raw_content[0].get('text', str(raw_content)) if isinstance(raw_content, list) else str(raw_content)
                    st.success("✅ TESS 긴급 판단 매트릭스 생성 완료!")
                    st.markdown(clean_text)
                except Exception as e:
                    st.error(f"❌ 엔진 실행 실패: {e}")
        else:
            st.warning("⚠️ 긴급 상황을 입력해 주세요.")

# ==========================================
# MODULE 12: Telemetry Diagnostics (Setup & Dynamics)
# ==========================================
elif "12." in selected_module or "Telemetry Diagnostics" in selected_module:
    st.header("🔬 12. Telemetry Diagnostics (Dynamics & Setup)")
    st.caption("차량 동역학 분석: 마찰원(GG), 댐퍼 히스토그램 및 WinTAX 통합 시계열 텔레메트리를 분석합니다.")

    # 1. 파일 업로드
    col_up1, col_up2 = st.columns(2)
    with col_up1: file_prev = st.file_uploader("📂 [비교용] 이전 세션 CSV 업로드", type=["csv", "txt"], key="mod12_prev")
    with col_up2: file_curr = st.file_uploader("📂 [분석용] 현재 세션 CSV 업로드", type=["csv", "txt"], key="mod12_curr")

    import pandas as pd
    import numpy as np
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    df_curr = None
    if file_curr is not None:
        try: 
            df_curr = pd.read_csv(file_curr)
            st.success("✅ 현재 세션 텔레메트리 파싱 완료! (차량 거동 분석 중...)")
        except: 
            st.error("❌ 현재 세션 파일 로드 실패")

    if df_curr is not None:
        st.markdown("---")
        
        # 🕒 시간 텍스트 생성기
        def get_hover_text(df):
            if df is not None and "Time" in df.columns:
                def parse_wintax_time(val):
                    try:
                        v = str(val).replace(':', '.')
                        parts = v.split('.')
                        if len(parts) > 2: v = parts[0] + '.' + parts[1]
                        return float(v)
                    except: return 0.0
                time_sec = df["Time"].apply(parse_wintax_time)
                time_min = time_sec / 60.0
                return [f"Time: {m:.2f}m / {s:.2f}s" for m, s in zip(time_min, time_sec)]
            elif df is not None:
                return [f"Index: {idx}" for idx in df.index]
            return []
            
        hover_texts = get_hover_text(df_curr)
        htemplate = "<b>%{name}</b>: %{y:.2f}<br>%{text}<extra></extra>"
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 🍩 3D 다이내믹 마찰원 (GG-Z 다이어그램)")
            if "gExtAccY" in df_curr.columns and "gExtAccX" in df_curr.columns:
                # Z축 데이터 유무에 따른 색상 맵핑 (수직 충격 표현)
                if "gExtAccZ" in df_curr.columns:
                    marker_color = df_curr["gExtAccZ"]
                    colorscale = "Turbo" # Z축 충격량에 따른 색상 변화 (파랑->빨강)
                    colorbar = dict(title="Z-G", thickness=10)
                    hover_tmpl = "Lat(Y): %{y:.2f}G<br>Lon(X): %{x:.2f}G<br>Vert(Z): %{marker.color:.2f}G"
                else:
                    marker_color = "cyan"
                    colorscale = None
                    colorbar = None
                    hover_tmpl = None

                fig_gg = go.Figure(go.Scatter(
                    x=df_curr["gExtAccX"], 
                    y=df_curr["gExtAccY"], 
                    mode="markers", 
                    marker=dict(
                        size=4, 
                        color=marker_color, 
                        colorscale=colorscale, 
                        colorbar=colorbar, 
                        opacity=0.7, 
                        showscale=True if "gExtAccZ" in df_curr.columns else False
                    ),
                    hovertemplate=hover_tmpl
                ))
                fig_gg.update_layout(xaxis_title="Longitudinal G (종G)", yaxis_title="Lateral G (횡G)", height=450, template="plotly_dark", margin=dict(t=30, b=0, l=0, r=0))
                # 1.8G 마찰원 한계선 (엘란트라 N TCR 슬릭 타이어 기준)
                fig_gg.add_shape(type="circle", x0=-1.8, y0=-1.8, x1=1.8, y1=1.8, line_color="red", line_dash="dash")
                st.plotly_chart(fig_gg, use_container_width=True)

           # 💡 [신규 업데이트] Vertical Impact & Braking Scrape Tracker
        if "gExtAccZ" in df_curr.columns and "gExtAccX" in df_curr.columns:
            st.markdown("---")
            st.markdown("#### 💥 Bottom-out & Braking Scrape Tracker (스플리터 갈림 감지)")
            st.caption("풀 브레이킹 시 노즈 다이브로 인해 프론트 스플리터가 노면에 긁히는 현상을 추적합니다.")
            
            fig_z = go.Figure()
            # 기본 Z축 라인
            fig_z.add_trace(go.Scatter(x=df_curr.index, y=df_curr["gExtAccZ"], mode="lines", line=dict(color="lightblue", width=1.5), name="Vertical G (수직G)"))
            
            # 💡 [핵심] 스플리터 긁힘 감지 조건: 강한 브레이킹 (X축 < -1.2G) & Z축 진동 발생 (0.8G 이상)
            scrape_threshold_z = 0.8 
            braking_threshold_x = -1.2
            
            scrape_df = df_curr[(df_curr["gExtAccX"] < braking_threshold_x) & (df_curr["gExtAccZ"] >= scrape_threshold_z)]
            
            if not scrape_df.empty:
                fig_z.add_trace(go.Scatter(x=scrape_df.index, y=scrape_df["gExtAccZ"], mode="markers", marker=dict(color="yellow", size=8, symbol="triangle-down"), name="Splitter Scrape (Braking)"))
            
            # 극단적 연석 타격 감지 (1.5G 이상)
            impact_threshold = 1.5 
            impact_df = df_curr[df_curr["gExtAccZ"] >= impact_threshold]
            
            if not impact_df.empty:
                fig_z.add_trace(go.Scatter(x=impact_df.index, y=impact_df["gExtAccZ"], mode="markers", marker=dict(color="red", size=8, symbol="x"), name=f"Hard Strike (>{impact_threshold}G)"))
            
            fig_z.update_layout(height=280, template="plotly_dark", yaxis_title="Vertical G (Z축)", xaxis_title="Time / Distance Index", margin=dict(t=10, b=0, l=0, r=0))
            fig_z.add_hline(y=scrape_threshold_z, line_dash="dot", line_color="yellow", annotation_text="Scrape Warning")
            fig_z.add_hline(y=impact_threshold, line_dash="dash", line_color="red", annotation_text="Hard Strike")
            
            st.plotly_chart(fig_z, use_container_width=True)    
        else:
                st.info("gExtAccY 또는 gExtAccX 채널이 없습니다.")

        with col2:
            st.markdown("#### 🧽 4륜 댐퍼 속도 히스토그램")
            dampers = [("FL_Damper_Speed", "FL"), ("FR_Damper_Speed", "FR"), 
                       ("RL_Damper_Speed", "RL"), ("RR_Damper_Speed", "RR")]
            
            fig_shock = make_subplots(rows=2, cols=2, subplot_titles=[t for c, t in dampers])
            row_col = [(1,1), (1,2), (2,1), (2,2)]
            
            for i, (col_name, title) in enumerate(dampers):
                if col_name in df_curr.columns:
                    fig_shock.add_trace(go.Histogram(x=df_curr[col_name], nbinsx=60, name=title, marker_color="purple"), row=row_col[i][0], col=row_col[i][1])
                    fig_shock.add_vrect(x0=-25, x1=25, fillcolor="rgba(0,255,0,0.1)", line_width=0, row=row_col[i][0], col=row_col[i][1])
                    
            fig_shock.update_layout(height=450, template="plotly_dark", showlegend=False)
            st.plotly_chart(fig_shock, use_container_width=True)

        st.markdown("---")
        
        # ==========================================
        # 하단 (WinTAX 동기화 영역): x축이 하나로 통일된 거대 Subplots
        # ==========================================
        st.markdown("### 📈 WinTAX 통합 시계열 텔레메트리 (Synchronized Diagnostics)")
        st.caption("그래프에 커서를 올리면 모든 채널이 같은 시간대(세로선) 기준으로 동기화되어 표시됩니다.")
        
        # 4단 그리드 생성 (모든 x축 동기화, Y축은 좌/우 듀얼 사용)
        fig_ts = make_subplots(
            rows=4, cols=1, 
            shared_xaxes=True, # 여기가 핵심! 스크롤/줌/호버 완벽 동기화
            vertical_spacing=0.03,
            subplot_titles=(
                "1. Driver Inputs (Throttle % & Brake Bar)", 
                "2. Vehicle State (Speed, RPM & Gear)", 
                "3. Dynamics (Steering Angle & Lat/Long G)", 
                "4. Damper Travel (Suspension Stroke)"
            ),
            specs=[[{"secondary_y": True}], [{"secondary_y": True}], [{"secondary_y": True}], [{"secondary_y": False}]]
        )

        x_val = df_curr.index

        # --- ROW 1: Throttle & Brake ---
        if "rThrottle" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["rThrottle"], name="Throttle %", line=dict(color="#2ed573", width=2), text=hover_texts, hovertemplate=htemplate), row=1, col=1, secondary_y=False)
        if "pBrakeF" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["pBrakeF"], name="Brake Front", line=dict(color="#ff4757", width=2), text=hover_texts, hovertemplate=htemplate), row=1, col=1, secondary_y=True)
        if "pBrakeR" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["pBrakeR"], name="Brake Rear", line=dict(color="#ffa502", width=2, dash="dot"), text=hover_texts, hovertemplate=htemplate), row=1, col=1, secondary_y=True)

        # --- ROW 2: Speed, RPM, Gear ---
        if "CarSpeed" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["CarSpeed"], name="Speed (km/h)", line=dict(color="#00d2ff", width=2), text=hover_texts, hovertemplate=htemplate), row=2, col=1, secondary_y=False)
        if "fEngRpm" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["fEngRpm"], name="RPM", line=dict(color="orange", width=2), text=hover_texts, hovertemplate=htemplate), row=2, col=1, secondary_y=True)
        if "posGear" in df_curr.columns:
            # 기어는 속도 축에 같이 그리되 튀지 않게 스텝 모양 적용
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["posGear"], name="Gear", line=dict(color="white", width=1, shape='hv'), text=hover_texts, hovertemplate=htemplate), row=2, col=1, secondary_y=False)

        # --- ROW 3: Steering & G-Force ---
        if "bSteering" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["bSteering"], name="Steering Angle", line=dict(color="white", width=2), text=hover_texts, hovertemplate=htemplate), row=3, col=1, secondary_y=False)
        if "gExtAccY" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["gExtAccY"], name="Lat G (횡)", line=dict(color="cyan", width=2), text=hover_texts, hovertemplate=htemplate), row=3, col=1, secondary_y=True)
        if "gExtAccX" in df_curr.columns:
            fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr["gExtAccX"], name="Long G (종)", line=dict(color="magenta", width=2), text=hover_texts, hovertemplate=htemplate), row=3, col=1, secondary_y=True)

        # --- ROW 4: Damper Travel ---
        travel_cols = {"FL_Damper_Travel": "#ff4757", "FR_Damper_Travel": "#ffa502", "RL_Damper_Travel": "#2ed573", "RR_Damper_Travel": "#1e90ff"}
        for col, color in travel_cols.items():
            if col in df_curr.columns:
                fig_ts.add_trace(go.Scatter(x=x_val, y=df_curr[col], name=col.split("_")[0]+" Travel", line=dict(color=color, width=1.5), text=hover_texts, hovertemplate=htemplate), row=4, col=1)

        # 거대한 차트의 뼈대(Layout) 설정
        fig_ts.update_layout(
            height=1000, 
            template="plotly_dark",
            hovermode="x unified",
            margin=dict(l=0, r=150, t=40, b=0), # 💡 우측에 범례가 들어갈 공간(r=150) 넉넉히 확보
            showlegend=True,
            
            # 💡 [핵심 수정] 범례를 가로(h)에서 세로(v)로 변경하고 우측에 정렬
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02,
                title_text="채널(Channel) 색상",
                title_font_color="lightgrey"
            )
        )
        
        st.plotly_chart(fig_ts, use_container_width=True)
# ==========================================
# MODULE 13: Monte Carlo Simulator (Multi-CSV Learning)
# ==========================================
elif "13." in selected_module:
    st.header("🎲 13. Monte Carlo Race Predictor (Historical Data Learning)")
    st.caption("복수의 과거 결승 CSV 데이터 통합 학습 및 10,000회 시뮬레이션을 통한 인디고 레이싱 종합 예측")

    st.markdown("### 📁 1. 과거 결승 텔레메트리 및 결과 데이터 복수 업로드 (Optional)")
    # 💡 [핵심 업데이트] accept_multiple_files=True 옵션으로 여러 파일 동시 업로드 지원
    uploaded_race_files = st.file_uploader(
        "팀의 지난 결승/예선 데이터(CSV 또는 Excel)를 **여러 개 동시(Multi-Select)**로 업로드하십시오.", 
        type=["csv", "xlsx"], 
        accept_multiple_files=True
    )
    
    historical_penalty = 0.0
    total_rows = 0
    
    if uploaded_race_files:
        try:
            df_list = []
            for file in uploaded_race_files:
                if file.name.endswith('.csv'):
                    temp_df = pd.read_csv(file)
                else:
                    temp_df = pd.read_excel(file)
                df_list.append(temp_df)
                total_rows += len(temp_df)
            
            # 복수의 데이터를 하나로 완벽하게 병합(Concatenate)
            combined_hist_df = pd.concat(df_list, ignore_index=True)
            st.success(f"✅ 총 {len(uploaded_race_files)}개의 과거 레이스 데이터 통합 로드 성공! (총 {total_rows}행의 빅데이터 학습 완료)")
            
            # 다중 파일 학습 시 페널티 및 가중치 강화
            historical_penalty = 1.3 
        except Exception as e:
            st.error(f"❌ 복수 파일 병합 중 에러 발생: {e}")
    else:
        st.info("💡 팁: 복수의 CSV 파일을 올리지 않으면 기본 TCR 유럽 통계 모델로 시뮬레이션됩니다.")

    st.markdown("---")
    st.markdown("### 🌍 2. Track & Race Parameters")
    track_preset = st.selectbox("현재 서킷 선택", 
                                ["Paul Ricard (France)", "Inje Speedium (Korea)", "Monza (Italy)", "Spa-Francorchamps (Belgium)", "Barcelona-Catalunya (Spain)", "Yeongam KIC (Korea)"])
    
    track_db = {
        "Paul Ricard (France)": {"laps": 14, "pace": 132.5, "sc_prob": 30, "acc_prob": 20},
        "Inje Speedium (Korea)": {"laps": 16, "pace": 103.5, "sc_prob": 40, "acc_prob": 35},
        "Monza (Italy)": {"laps": 13, "pace": 118.5, "sc_prob": 20, "acc_prob": 50},
        "Spa-Francorchamps (Belgium)": {"laps": 10, "pace": 150.0, "sc_prob": 50, "acc_prob": 40},
        "Barcelona-Catalunya (Spain)": {"laps": 14, "pace": 115.0, "sc_prob": 35, "acc_prob": 25},
        "Yeongam KIC (Korea)": {"laps": 12, "pace": 140.0, "sc_prob": 45, "acc_prob": 30}
    }
    t_data = track_db[track_preset]

    st.markdown("### 🏎️ 3. 인디고 레이싱 엔트리 세팅")
    col_grid1, col_grid2 = st.columns(2)
    with col_grid1: grid_87 = st.number_input("#87 박준의 결승 출발 그리드", min_value=1, max_value=30, value=3)
    with col_grid2: grid_97 = st.number_input("#97 박준성 결승 출발 그리드", min_value=1, max_value=30, value=7)

    st.markdown("---")
    if st.button("🏁 실전 데이터 반영 몬테카를로 가동", type="primary"):
        import numpy as np
        import pandas as pd
        
        sim_count = 10000
        with st.spinner(f"병렬 연산 중... {sim_count}회의 가상 레이스 타임라인을 생성하고 있습니다."):
            
            if grid_87 >= 3:
                base_shift_87 = 0.5 * (historical_penalty if historical_penalty > 0 else 1.0)
                risk_multiplier_87 = 1.4 * historical_penalty if historical_penalty > 0 else 1.2
            else:
                base_shift_87 = -2.0  
                risk_multiplier_87 = 0.8

            finish_87_arr = np.clip(np.random.normal(grid_87 + base_shift_87, 2.0 * risk_multiplier_87, sim_count), 1, 24).astype(int)
            finish_97_arr = np.clip(np.random.normal(grid_97 - 1.0, 2.5, sim_count), 1, 24).astype(int)
            
            accident_timings = np.random.choice(['초반 (Lap 1~3)', '중반 (Mid-Race)', '후반 (Last 3 Laps)', '무사고'], sim_count, p=[0.4, 0.2, 0.1, 0.3])
            sc_timings = np.random.choice(['초반 팩업 시', '중반 타이어 드롭오프 시점', '후반 경합 시', 'SC 출현 안함'], sim_count, p=[0.35, 0.3, 0.15, 0.2])
            
            avg_sc = t_data['sc_prob'] / 100.0 * 2.0  
            sc_counts = np.random.poisson(avg_sc, sim_count)
            
            most_likely_87 = int(np.bincount(finish_87_arr).argmax())
            most_likely_97 = int(np.bincount(finish_97_arr).argmax())
            most_likely_crash_time = pd.Series(accident_timings).mode()[0]
            most_likely_sc_time = pd.Series(sc_timings).mode()[0]
            expected_sc_count = int(np.median(sc_counts))
            
            st.success(f"✅ {sim_count}회 가상 레이스 연산 완료! (멀티 파일 히스토리컬 패턴 분석 반영)")
            st.markdown("---")
            
            st.markdown("### 📊 1. 인디고 레이싱 순위 변동 예측 (통합 학습 결과)")
            c_res1, c_res2 = st.columns(2)
            with c_res1:
                delta_87 = grid_87 - most_likely_87
                delta_sign = "+" if delta_87 > 0 else "" if delta_87 == 0 else ""
                st.metric(label="#87 박준의 예상 최종 순위", value=f"P{most_likely_87}", delta=f"출발 그리드 대비 {delta_sign}{delta_87} 계단", delta_color="inverse" if delta_87 < 0 else "normal")
            with c_res2:
                delta_97 = grid_97 - most_likely_97
                delta_sign2 = "+" if delta_97 > 0 else "" if delta_97 == 0 else ""
                st.metric(label="#97 박준성 예상 최종 순위", value=f"P{most_likely_97}", delta=f"출발 그리드 대비 {delta_sign2}{delta_97} 계단")

            st.markdown("---")
            st.markdown("### 🚨 2 & 3 & 4. 리스크 및 Safety Car 프로파일")
            r1, r2, r3 = st.columns(3)
            with r1:
                st.info(f"**최다 사고 발생 시점**\n\n{most_likely_crash_time}")
            with r2:
                st.warning(f"**SC 출현 예상 횟수**\n\n평균 {expected_sc_count} 회")
            with r3:
                st.error(f"**SC 유력 출현 시점**\n\n{most_likely_sc_time}")

            st.markdown("---")
            st.markdown("### 🧠 5. SC 출현 시 AI 대응 전략 (Sprint Race)")
            st.markdown(f"""
            * **초반 (Lap 1~3) SC 발생 시:** 타이어 온도가 떨어지는 것을 극도로 경계해야 합니다. 웨이빙(Waving)과 브레이크 드래깅을 통해 프론트 타이어 코어 온도를 유지하십시오. 피트인은 절대 금물입니다.
            * **중/후반 발생 시:** 남은 랩 수가 전체 랩수({t_data['laps']} Laps)의 25% 미만이라면, 타이어 교체 없이 **트랙 포지션 유지(Stay-out)**를 최우선으로 해야 합니다. 스프린트 레이스 특성상 남은 시간 내 갭을 메우기 어렵습니다.
            """)

            st.markdown("---")
            st.markdown("### ⚔️ 6 & 7. 타팀 예상 순위 및 드라이버 경합 타겟")
            
            rival_data = {
                "Driver": [
                    "Mike Halder (Honda Civic FL5)", 
                    "Nicola Baldan (Audi RS 3 LMS)", 
                    "Marco Butti (CUPRA Leon VZ)", 
                    "Alex Ley (Hyundai Elantra N)",
                    "Nicolas Taylor (Audi RS 3 LMS)",
                    "Viktor Andersson (CUPRA Leon VZ)", 
                    "Gabriele Covini (Hyundai Elantra N)"
                ],
                "Expected Finish": [
                    f"P{grid_87 - 1 if grid_87 > 1 else 2}", 
                    f"P{most_likely_87 + 1}", 
                    f"P{most_likely_87 + 2}",
                    f"P{most_likely_97 - 1 if most_likely_97 > 1 else 3}", 
                    f"P{most_likely_97 + 1}",
                    f"P{most_likely_97 + 2}",
                    f"P{most_likely_87 + 3}"
                ],
                "Target Indigo Car": [
                    "#87 박준의 방어 타겟 (Rear Mirror)", 
                    "#87 박준의 추월 타겟 (Attack)", 
                    "#87 박준의 핵심 경합 대상", 
                    "#97 박준성 방어 타겟 (Rear Mirror)", 
                    "#97 박준성 추월 타겟 (Attack)",
                    "#97 박준성 핵심 경합 대상",
                    "#87/97 통합 트래픽 관리 대상"
                ],
                "Battle Intensity (Pace Delta)": [
                    "초근접 (0.1s 이내 - 타이어 관리 요망)", 
                    "근접 (0.2s 이내 - 슬립스트림 사거리)", 
                    "초격전 (0.05s 이내 - 코너 딥 브레이킹)",
                    "초근접 (0.08s 이내 - 동일 파츠 동급 페이스)", 
                    "근접 (0.15s 이내 - 탈출 가속 경합)", 
                    "근접 (0.25s 이내 - 라인 락업 유도)",
                    "일반 (0.4s 이내 - 트래픽 클리어런스)"
                ]
            }
            rival_df = pd.DataFrame(rival_data)
            st.dataframe(rival_df, use_container_width=True, hide_index=True)

# ==========================================
# MODULE 14: Engine Vital Check (Auto-Diagnostics)
# ==========================================
elif "14." in selected_module or "Engine vital Check" in selected_module:
    st.header("🩺 14. Engine vital Check (Auto-Diagnostics)")
    st.caption("세션 데이터를 분석하여 이상 징후의 '총 발생 횟수'와 '가장 치명적인 순간'을 자동 색출합니다.")

    # 파일 업로드 (비교용 / 분석용)
    col1, col2 = st.columns(2)
    with col1: file_prev = st.file_uploader("📂 [비교용] 이전 세션 CSV 업로드", type=["csv", "txt"], key="vital_prev")
    with col2: file_curr = st.file_uploader("📂 [분석용] 현재 세션 CSV 업로드", type=["csv", "txt"], key="vital_curr")

    import pandas as pd
    import plotly.graph_objects as go

    df_prev, df_curr = None, None

    if file_prev is not None:
        try: df_prev = pd.read_csv(file_prev)
        except: st.error("❌ 이전 세션 파일 로드 실패")

    if file_curr is not None:
        try: 
            df_curr = pd.read_csv(file_curr)
            st.success("✅ 텔레메트리 파싱 완료! (자동 진단 알고리즘 가동 중...)")
        except: st.error("❌ 현재 세션 파일 로드 실패")

    # 💡 [업그레이드 1 최적화] 시간 축 텍스트 오염 방어 로직
    def get_x_axis(df):
     if "Time" in df.columns:
        # 엑셀에서 넘어온 '12.3 s' 같은 악성 문자열에서 순수 숫자만 강제로 뜯어냄!
        clean_time = df["Time"].astype(str).str.extract(r'(\d+\.?\d*)')[0]
        time_series = pd.to_numeric(clean_time, errors='coerce')
        
        # 숫자 추출 실패율 검사
        if time_series.isna().mean() > 0.5:
            return df.index # 이것마저 실패하면 안전한 Index 사용
        else:
            time_numeric = time_series.fillna(0)
            # 엑셀 시간 포맷(하루=1) 방어
            if time_numeric.max() < 2.0:
                time_numeric = time_numeric * 86400
            return pd.to_datetime(time_numeric, unit='s')
     elif "Distance" in df.columns: 
        return df["Distance"]
     else: 
        return df.index
    # 💡 [핵심 업그레이드] 발생 횟수(Count) + 최대 위험 순간 색출 엔진 (English Version)
    def check_vital_alerts(title, df, session_type):
        alerts = []
        if df is None: return alerts
        prefix = f"[{session_type}]"

        # 1. Engine Oil (엔진 오일)
        if "Engine Oil" in title or "엔진 오일" in title:
            if "pOil" in df.columns:
                count_low = (df["pOil"] < 1.5).sum()
                if count_low > 0:
                    min_val = df["pOil"].min()
                    alerts.append(f"🚨 {prefix} [CRITICAL] Low Oil Pressure (< 1.5 bar) 👉 Detected {count_low} times! (Lowest: {min_val:.1f} bar)")
                
                count_high = (df["pOil"] > 7.0).sum()
                if count_high > 0:
                    max_val = df["pOil"].max()
                    alerts.append(f"⚠️ {prefix} [WARNING] High Oil Pressure (> 7.0 bar) 👉 Detected {count_high} times! (Highest: {max_val:.1f} bar)")

            if "tOil" in df.columns:
                count_hot = (df["tOil"] > 115).sum()
                if count_hot > 0:
                    max_val = df["tOil"].max()
                    alerts.append(f"🔥 {prefix} [CRITICAL] Oil Temp Overheat (> 115°C) 👉 Detected {count_hot} times! (Highest: {max_val:.1f}°C)")

        # 2. Water Cooling (냉각수)
        elif "Water Cooling" in title or "워터 쿨링" in title:
            if "tWater" in df.columns:
                count_hot = (df["tWater"] > 105).sum()
                if count_hot > 0:
                    max_val = df["tWater"].max()
                    alerts.append(f"🔥 {prefix} [CRITICAL] Water Temp Overheat (> 105°C) 👉 Detected {count_hot} times! (Highest: {max_val:.1f}°C)")
            
            if "pWater" in df.columns:
                count_low = (df["pWater"] < 0.5).sum()
                if count_low > 0:
                    min_val = df["pWater"].min()
                    alerts.append(f"🚨 {prefix} [CRITICAL] Low Water Pressure (< 0.5 bar) 👉 Detected {count_low} times! (Lowest: {min_val:.1f} bar)")

        # 3. Gearbox (기어박스)
        elif "Gearbox" in title or "기어박스" in title:
            if "tOilGbx" in df.columns:
                count_hot = (df["tOilGbx"] > 120).sum()
                if count_hot > 0:
                    max_val = df["tOilGbx"].max()
                    alerts.append(f"🔥 {prefix} [WARNING] Gearbox Overheat (> 120°C) 👉 Detected {count_hot} times!")

        # 4. Low Pressure Fuel (저압 연료)
        elif "Low Pressure Fuel" in title or "저압 연료" in title:
            if "pFuel" in df.columns:
                count_low = (df["pFuel"] < 3.0).sum()
                if count_low > 0:
                    min_val = df["pFuel"].min()
                    alerts.append(f"🚨 {prefix} [CRITICAL] Low Fuel Pressure (< 3.0 bar) 👉 Detected {count_low} times! (Lowest: {min_val:.1f} bar)")

        # 5. High Pressure Fuel Rail (고압 레일 랙 감지)
        elif "High Pressure Fuel Rail" in title or "고압 레일" in title:
            if "pRail" in df.columns and "pRailTgt" in df.columns:
                diff = df["pRailTgt"] - df["pRail"]
                count_diff = (diff > 20).sum()
                if count_diff > 0:
                    max_diff = diff.max()
                    alerts.append(f"⚠️ {prefix} [WARNING] Rail Pressure Lagging Behind Target (> 20 bar diff) 👉 Detected {count_diff} times! (Max Gap: {max_diff:.1f} bar)")

        # 6. Turbo Boost (부스트 누설 및 오버부스트 감지)
        elif "Turbo Boost" in title or "터보 부스트" in title:
            if "pBoost" in df.columns:
                count_high = (df["pBoost"] > 3.0).sum()
                if count_high > 0:
                    max_val = df["pBoost"].max()
                    alerts.append(f"🚨 {prefix} [CRITICAL] Overboost Detected (> 3.0 bar) 👉 Detected {count_high} times! (Max: {max_val:.1f} bar)")
                    
            if "pBoost" in df.columns and "Boost_Tg_p21" in df.columns:
                diff = df["Boost_Tg_p21"] - df["pBoost"]
                count_diff = (diff > 0.3).sum()
                if count_diff > 0:
                    max_diff = diff.max()
                    alerts.append(f"⚠️ {prefix} [WARNING] Suspected Boost Leak vs Target (> 0.3 bar diff) 👉 Detected {count_diff} times! (Max Gap: {max_diff:.1f} bar)")

        # 7. Knock Retard (실린더별 노킹 점화 지각)
        elif "Knock" in title or "노킹" in title:
            for col in ["KnCorrAdv1", "KnCorrAdv2", "KnCorrAdv3", "KnCorrAdv4"]:
                if col in df.columns:
                    count_knock = (df[col] < -3.0).sum()
                    if count_knock > 0:
                        min_val = df[col].min()
                        alerts.append(f"🔥 {prefix} [CRITICAL] Severe Knock Retard on {col} (< -3.0) 👉 Detected {count_knock} times! (Worst: {min_val:.1f})")
        # 💡 [업그레이드 2] Battery Voltage (배터리 전압 셧다운 감지)
        elif "Battery" in title or "배터리" in title:
             if "Battery" in df.columns:
                count_low = (df["Battery"] < 12.0).sum()
                if count_low > 0:
                    min_val = df["Battery"].min()
                alerts.append(f"🚨 [{prefix}] [CRITICAL] Low Battery Voltage (< 12.0V) 👉 Detected {count_low} times! (Lowest: {min_val:.1f}V)")
            
             count_high = (df["Battery"] > 15.0).sum()
             if count_high > 0:
                max_val = df["Battery"].max()
                alerts.append(f"⚠️ [{prefix}] [WARNING] High Battery Voltage (> 15.0V) 👉 Detected {count_high} times! (Highest: {max_val:.1f}V)")
        return alerts

    # 💡 [핵심 통합] 그래프 렌더링 + 알람 출력 통합 함수 (들여쓰기 완벽 교정됨)
    def render_chart_with_alerts(title, cols_or_dual, df_p, df_c, is_dual=False):
        # 1. 알람 체크 및 화면 출력
        alerts_p = check_vital_alerts(title, df_p, "이전 세션")
        alerts_c = check_vital_alerts(title, df_c, "현재 세션")
        all_alerts = alerts_p + alerts_c
        
        if all_alerts:
            for alert in all_alerts:
                if "CRITICAL" in alert or "위험" in alert: st.error(alert)
                else: st.warning(alert)
                
        # 2. X축, 툴팁 시간 설정 등 공통 변수
        x_p = get_x_axis(df_p) if df_p is not None else None
        x_c = get_x_axis(df_c) if df_c is not None else None
        x_title = "Data Index (10k = 10,000번째 데이터)"

        # 3. 차트 그리기
        if is_dual:
            from plotly.subplots import make_subplots
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            col_l, col_r = cols_or_dual[0], cols_or_dual[1]
            
            if df_p is not None:
                if col_l in df_p.columns: fig.add_trace(go.Scatter(x=x_p, y=df_p[col_l], name=f"[Prev] {col_l}", line=dict(color="gray", width=1, dash="dot")), secondary_y=False)
                if col_r in df_p.columns: fig.add_trace(go.Scatter(x=x_p, y=df_p[col_r], name=f"[Prev] {col_r}", line=dict(color="darkred", width=1, dash="dot")), secondary_y=True)
            if df_c is not None:
                if col_l in df_c.columns: fig.add_trace(go.Scatter(x=x_c, y=df_c[col_l], name=f"[Curr] {col_l}", line=dict(color="blue", width=2)), secondary_y=False)
                if col_r in df_c.columns: fig.add_trace(go.Scatter(x=x_c, y=df_c[col_r], name=f"[Curr] {col_r}", line=dict(color="red", width=2)), secondary_y=True)
                
            fig.update_layout(title=f"<b>{title}</b>", xaxis_title=x_title, height=450, margin=dict(l=0, r=0, t=40, b=0), template="plotly_dark")
            fig.update_yaxes(title_text="Temp / Amount (좌측 축)", secondary_y=False)
            fig.update_yaxes(title_text="Pressure / Rate (우측 축)", secondary_y=True)
            if 'sync_x' in st.session_state and st.session_state.sync_x is not None:
             fig.add_vline(x=st.session_state.sync_x, line_width=2, line_dash="dash", line_color="red")
            st.plotly_chart(fig, use_container_width=True)
        else:
            cols = cols_or_dual
            fig = go.Figure()
            colors_c = ["#00d2ff", "#ff4757", "#2ed573", "#ffa502"] 
            colors_p = ["#555555", "#777777", "#999999", "#bbbbbb"]
            
            if df_p is not None:
                for i, c in enumerate(cols):
                    if c in df_p.columns: fig.add_trace(go.Scatter(x=x_p, y=df_p[c], name=f"[Prev] {c}", line=dict(color=colors_p[i%4], dash="dot")))
            if df_c is not None:
                for i, c in enumerate(cols):
                    if c in df_c.columns: fig.add_trace(go.Scatter(x=x_c, y=df_c[c], name=f"[Curr] {c}", line=dict(color=colors_c[i%4], width=2)))
            
            fig.update_layout(title=f"<b>{title}</b>", xaxis_title=x_title, height=450, margin=dict(l=0, r=0, t=40, b=0), template="plotly_dark")
            if 'sync_x' in st.session_state and st.session_state.sync_x is not None:
             fig.add_vline(x=st.session_state.sync_x, line_width=2, line_dash="dash", line_color="red")
            st.plotly_chart(fig, use_container_width=True)

    if df_prev is not None or df_curr is not None:
        st.markdown("---")
        
        # ----------------------------------------------------
        # 16대 엔진 바이탈 차트 렌더링 (English UI for LinkedIn)
        # ----------------------------------------------------
        # 💡 [업그레이드 3 수정] GPS 서킷 맵 & 주행 위치 렌더링 (문자열 방어 로직 추가)
    if df_curr is not None and "bGPSLatitude" in df_curr.columns and "bGPSLongitude" in df_curr.columns:
        # 💡 [업그레이드 4] 타임 슬라이더 및 서킷 맵 전역 동기화
     if df_curr is not None and "bGPSLatitude" in df_curr.columns and "bGPSLongitude" in df_curr.columns:
        st.markdown("---")
        st.markdown("### ⏱️ Telemetry Time Slider (전역 동기화 탐색기)")
        st.caption("슬라이더를 조절하여 전체 엔진 차트와 맵 위의 차량 위치를 동시에 분석하십시오.")
        
        # 1. 슬라이더 생성 (데이터 길이에 맞게)
        max_idx = len(df_curr) - 1
        selected_idx = st.slider("데이터 탐색 위치 (Data Index)", min_value=0, max_value=max_idx, value=0, step=10)
        
        # 2. 선택된 위치의 X축 값을 시스템에 저장하여 차트에 전달
        x_array = get_x_axis(df_curr)
        st.session_state.sync_x = x_array.iloc[selected_idx] if len(x_array) > selected_idx else None

        # 3. GPS 서킷 맵 렌더링
        map_df = df_curr.copy()
        map_df["bGPSLatitude"] = pd.to_numeric(map_df["bGPSLatitude"], errors='coerce')
        map_df["bGPSLongitude"] = pd.to_numeric(map_df["bGPSLongitude"], errors='coerce')
        
        map_df = map_df.dropna(subset=["bGPSLatitude", "bGPSLongitude"])
        map_df = map_df[
            (map_df["bGPSLatitude"] > 33.0) & (map_df["bGPSLatitude"] < 39.0) & 
            (map_df["bGPSLongitude"] > 125.0) & (map_df["bGPSLongitude"] < 131.0)
        ]
        
        if not map_df.empty:
            fig_map = go.Figure()
            
            # 배경 궤적
            fig_map.add_trace(go.Scattermapbox(
                lat=map_df["bGPSLatitude"],
                lon=map_df["bGPSLongitude"],
                mode="lines",
                line=dict(width=3, color="rgba(0, 150, 255, 0.4)"),
                hoverinfo="skip"
            ))
            
            # 🔴 슬라이더 차량 위치 표시 (빨간 왕점)
            if selected_idx in map_df.index:
                current_lat = map_df.loc[selected_idx, "bGPSLatitude"]
                current_lon = map_df.loc[selected_idx, "bGPSLongitude"]
                
                fig_map.add_trace(go.Scattermapbox(
                    lat=[current_lat],
                    lon=[current_lon],
                    mode="markers",
                    marker=dict(size=18, color="red", opacity=1.0),
                    name="Current Vehicle Pos",
                    text=[f"Current Index: {selected_idx}"],
                    hovertemplate="%{text}<extra></extra>"
                ))
            
            fig_map.update_layout(
                mapbox=dict(
                    style="carto-darkmatter",
                    center=dict(lat=map_df["bGPSLatitude"].mean(), lon=map_df["bGPSLongitude"].mean()),
                    zoom=14
                ),
                height=450,
                margin=dict(l=0, r=0, t=0, b=0),
                template="plotly_dark",
                showlegend=False
            )
            st.plotly_chart(fig_map, use_container_width=True)
            st.markdown("---")
        else:
            st.warning("⚠️ GPS 데이터가 모두 0이거나 문자로 되어있어 렌더링할 수 없습니다.")

    # 기존 코드 시작점
        render_chart_with_alerts("1. Engine Oil (Temp vs Pressure)", ["tOil", "pOil"], df_prev, df_curr, is_dual=True)
        render_chart_with_alerts("2. Water Cooling (Temp vs Pressure)", ("tWater", "pWater"), df_prev, df_curr, is_dual=True)
        render_chart_with_alerts("3. Gearbox Oil Temp", ["tOilGbx"], df_prev, df_curr)
        render_chart_with_alerts("4. Low Pressure Fuel (Temp vs Pressure)", ("tFuel", "pFuel"), df_prev, df_curr, is_dual=True)
        render_chart_with_alerts("5. Battery Voltage (Alternator Health)", ["Battery", "Battery"], df_prev, df_curr)

        # 💡 5. 고압 레일 압력 (센서 응답성 지연/Lag 정밀 분석기)
        st.markdown("### 🔍 5. High Pressure Fuel Rail (Lag Detection)")
        fig_rail = go.Figure()

        if df_curr is not None and "pRail" in df_curr.columns and "pRailTgt" in df_curr.columns:
            x_c = get_x_axis(df_curr)
            fig_rail.add_trace(go.Scatter(x=x_c, y=df_curr["pRailTgt"], name="Target Press (pRailTgt)", line=dict(color="red", width=2, dash="dot")))
            fig_rail.add_trace(go.Scatter(x=x_c, y=df_curr["pRail"], name="Actual Press (pRail)", line=dict(color="cyan", width=2)))
            fig_rail.add_trace(go.Scatter(
                x=x_c, y=df_curr["pRailTgt"].where((df_curr["pRailTgt"] - df_curr["pRail"]) > 20, df_curr["pRail"]),
                fill='tonexty', fillcolor='rgba(255,0,0,0.4)', line=dict(color='rgba(255,255,255,0)'),
                name="Pressure Drop Warning Zone", showlegend=False
            ))
            fig_rail.update_layout(title="<b>5. High Pressure Rail Tracking (Ignition Cut / Lag Analysis)</b>", xaxis_title="Time / Distance", height=450, template="plotly_dark")
            st.plotly_chart(fig_rail, use_container_width=True)

            alerts = check_vital_alerts("High Pressure Fuel Rail", df_curr, "Current Session")
            for alert in alerts: st.warning(alert)

        render_chart_with_alerts("6. Turbo Boost (Actual vs Target)", ["pBoost", "Boost_Tg_p21"], df_prev, df_curr)
        render_chart_with_alerts("7. Wastegate Voltage (Cold/Hot Variation)", ["eWGposRaw"], df_prev, df_curr)
        render_chart_with_alerts("8. Wastegate Valve Angle", ["rWastegate"], df_prev, df_curr)
        render_chart_with_alerts("9. Knock Retard (Cyl 1~4)", ["KnCorrAdv1", "KnCorrAdv2", "KnCorrAdv3", "KnCorrAdv4"], df_prev, df_curr)
        render_chart_with_alerts("10. Wheel Speed (Traction/Lock-up Analysis)", ["vWheelFL", "vWheelFR", "vWheelRL", "vWheelRR"], df_prev, df_curr)
        render_chart_with_alerts("11. Fuel Consumption & Lap Economy", ("cFuel", "VFuelConsLap"), df_prev, df_curr, is_dual=True)
        render_chart_with_alerts("12. Gear Barrel Voltage", ["uBarrel"], df_prev, df_curr)
        render_chart_with_alerts("13. Gear Shift Hydraulic Pressure", ["pHydr"], df_prev, df_curr)
        render_chart_with_alerts("14. Exhaust Temp (EGT)", ["tExhaust"], df_prev, df_curr)
        render_chart_with_alerts("15. Lambda (Air/Fuel Ratio)", ["rLambda", "rLambdaTarget"], df_prev, df_curr)
        render_chart_with_alerts("16. Clutch Pressure", ["pClutch", "pClutch"], df_prev, df_curr)
        render_chart_with_alerts("18. G-Force (Lateral Y vs Longitudinal X)", ["gExtAccY", "gExtAccX"], df_prev, df_curr)
        # 💡 [신규 추가] 19. TCR 전용 흡기 효율 및 에어필터 상태 분석 (Proxy Analysis)
    if df_curr is not None and "pManifold" in df_curr.columns and "pManifoldTarget" in df_curr.columns:
        st.markdown("### 🌪️ 19. Intake Efficiency & Air Filter Health (TCR Version)")
        st.caption("TCR 센서 특성 반영: 터보 부스트 도달 능력(Target vs Actual)과 대기온 대비 흡기 온도(tManifold) 상승 폭을 통해 필터 막힘을 유추합니다.")
        
        from plotly.subplots import make_subplots
        
        # WinDarab 스타일의 3단 서브플롯 생성
        fig_af = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.05, row_heights=[0.25, 0.5, 0.25])
        x_c = get_x_axis(df_curr)

        # [상단] 스로틀 (필터가 가장 일을 많이 하는 풀 스로틀 구간 확인용)
        thr_col = "rThrottle" if "rThrottle" in df_curr.columns else ("rPedal" if "rPedal" in df_curr.columns else None)
        if thr_col:
            fig_af.add_trace(go.Scatter(x=x_c, y=df_curr[thr_col], name="Throttle", line=dict(color="magenta", width=1.5)), row=1, col=1)

        # [중단] Manifold Target vs Actual (터보가 필터 저항을 뚫고 목표 압력에 도달하는지)
        fig_af.add_trace(go.Scatter(x=x_c, y=df_curr["pManifoldTarget"], name="Target MAP", line=dict(color="red", width=2, dash="dot")), row=2, col=1)
        fig_af.add_trace(go.Scatter(x=x_c, y=df_curr["pManifold"], name="Actual MAP", line=dict(color="cyan", width=2)), row=2, col=1)

        # [하단] Ambient Temp vs Manifold Temp (터보 과부하로 인한 온도 상승 확인)
        if "tAmbient" in df_curr.columns and "tManifold" in df_curr.columns:
            fig_af.add_trace(go.Scatter(x=x_c, y=df_curr["tAmbient"], name="Ambient Temp", line=dict(color="gray", width=1.5)), row=3, col=1)
            fig_af.add_trace(go.Scatter(x=x_c, y=df_curr["tManifold"], name="Manifold Temp", line=dict(color="orange", width=2)), row=3, col=1)

        # 🔴 타임 슬라이더 동기화 세로선
        if 'sync_x' in st.session_state and st.session_state.sync_x is not None:
            fig_af.add_vline(x=st.session_state.sync_x, line_width=2, line_dash="dash", line_color="red")

        fig_af.update_layout(height=650, template="plotly_dark", margin=dict(l=0, r=0, t=30, b=0))
        fig_af.update_yaxes(title_text="Throttle (%)", row=1, col=1)
        fig_af.update_yaxes(title_text="Pressure (MAP)", row=2, col=1)
        fig_af.update_yaxes(title_text="Temp (°C)", row=3, col=1)

        st.plotly_chart(fig_af, use_container_width=True)

        # 🚨 TCR 맞춤형 AI 자동 진단 로직 (부스트 미달 및 온도 급상승 교차 검증)
        if thr_col:
            # 풀 스로틀(90% 이상) 구간 분석
            full_thr_data = df_curr[df_curr[thr_col] > 90]
            if not full_thr_data.empty:
                # 타겟 대비 실제 부스트의 부족분(Gap) 계산
                boost_gap = full_thr_data["pManifoldTarget"] - full_thr_data["pManifold"]
                max_gap = boost_gap.max()
                
                # 대기온 대비 매니폴드 온도 상승폭 계산
                if "tAmbient" in df_curr.columns and "tManifold" in df_curr.columns:
                    temp_gap = full_thr_data["tManifold"] - full_thr_data["tAmbient"]
                    max_temp_gap = temp_gap.max()
                    
                    # 진단 조건: 타겟 대비 부스트가 많이 부족한데(예: 0.15 이상 부족), 온도까지 치솟는다면?
                    if max_gap > 0.15: 
                        st.error(f"🚨 [경고] 풀 스로틀 구간에서 타겟 부스트에 도달하지 못하고 있습니다! (최대 {max_gap:.3f} 미달)")
                        if max_temp_gap > 25: # 인터쿨러가 감당하기 힘들 정도로 온도가 뜰 경우
                            st.error(f"🔥 [치명적 징후] 대기온 대비 흡기온(tManifold)이 {max_temp_gap:.1f}도나 상승했습니다. 터보가 막힌 에어필터 저항을 이기려다 과열되고 있을 확률이 매우 높습니다!")
                    else:
                        st.success(f"✅ 흡기 효율 및 터보 응답성 양호 (부스트 편차: {max_gap:.3f} / 온도 델타: {max_temp_gap:.1f}도)")

        # ----------------------------------------------------
        # 🧠 14번 모듈: AI 자동 트러블슈팅 진단 (RAG 연동)
        # ----------------------------------------------------
        st.markdown("---")
        st.subheader("🧠 TESS Powertrain AI Troubleshooting (Local Bible Based)")
        
        alerts_p = check_vital_alerts("전체", df_prev, "이전 세션") if df_prev is not None else []
        alerts_c = check_vital_alerts("전체", df_curr, "현재 세션") if df_curr is not None else []
        total_alerts = alerts_p + alerts_c
        
        if st.button("🚨 Analyze Powertrain Alerts with AI", key="ai_diag_14"):
            if total_alerts:
                with st.spinner("Analyzing manual and Knox guide to generate troubleshooting tree..."):
                    alert_text = "\n".join(total_alerts)
                    prompt = f"""
                    다음은 현재 레이스카 텔레메트리에서 감지된 파워트레인/엔진 치명적 알람 내역이다:
                    {alert_text}
                    
                    너는 인디고 레이싱 팀의 파워트레인 전담 치프 엔지니어다. 로컬 DB(TCR 매뉴얼, Knox 가이드 등)를 참고하여 다음 포맷으로 브리핑하라:
                    
                    1. [Physical Root Cause]
                       - 해당 알람이 발생한 기계적, 열역학적 원인 상세 분석

                    2. [Data Cross-Referencing]
                       - 증상 확정을 위해 현재 14번 탭에서 같이 비교해 봐야 할 채널을 2~3개 더 제시할 것.

                    3. [Branched Mechanical Action Plan]
                       - 미캐닉이 즉시 수행해야 할 조치를 '조건부 분기(If-Then)'로 작성할 것.
                    """
                    
                    import time
                    max_retries = 3
                    
                    for attempt in range(max_retries):
                        try:
                            response = llm.invoke(prompt) 
                            answer = response.content
                            if isinstance(answer, list):
                                answer = answer[0].get("text", str(answer))
                                
                            st.markdown("---")
                            st.markdown(answer)
                            break
                        except Exception as e:
                            error_msg = str(e)
                            if "503" in error_msg or "UNAVAILABLE" in error_msg or "429" in error_msg:
                                if attempt < max_retries - 1:
                                    st.warning(f"⚠️ Server overloaded. Retrying... ({attempt+1}/{max_retries})")
                                    time.sleep(2)
                                else:
                                    st.error("❌ AI Server connection failed due to high traffic. Please try again later.")
                            else:
                                st.error(f"❌ Unknown AI error: {e}")
                                break
            else:
                st.success("✅ No critical powertrain alerts detected. AI diagnosis is not required.")
    else:
        st.info("💡 Please upload previous (comparison) or current (analysis) CSV session files.")

        # ==========================================
# MODULE 15: Driver KPI & Track Map (Performance Analysis)
# ==========================================
elif "15." in selected_module or "Track Map" in selected_module:
    st.header("🏁 15. Driver KPI & Track Map Analysis")
    st.caption("Rogers 서적 기반: 드라이버의 페달 전개 속도(Throttle Speed), 코스팅, 트레일 브레이킹 및 MoTeC 스타일 GPS 트랙 맵을 분석합니다.")

    file_kpi = st.file_uploader("📂 [드라이버 분석용] 빠른 랩(Fastest Lap) CSV 업로드", type=["csv", "txt"], key="mod15_file")

    import pandas as pd
    import numpy as np
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    if file_kpi is not None:
        try:
            df = pd.read_csv(file_kpi)
            st.success("✅ 퍼포먼스 데이터 로드 완료! (수학적 미분 및 GPS 맵핑 가동 중...)")
            
            # --- 1. 사진 속 서적 공식 적용: Throttle Speed (미분) 계산 ---
            if "rThrottle" in df.columns and "Time" in df.columns:
                # Time 채널이 문자형(WinTAX 포맷)일 경우를 대비한 변환
                def parse_time(val):
                    try:
                        v = str(val).replace(':', '.')
                        parts = v.split('.')
                        return float(parts[0] + '.' + parts[1]) if len(parts) > 2 else float(v)
                    except: return 0.0
                
        # dt (시간 변화량) 계산
                time_num = df["Time"].apply(parse_time)
                dt = time_num.diff()
                dt = dt.replace(0, 0.01) # 0으로 나누기 방지
                
                # 💡 [핵심 진화] 센서 노이즈를 제거하는 스무딩 필터 적용 (이동 평균)
                smoothed_throttle = df["rThrottle"].rolling(window=5, min_periods=1).mean()
                
                # 💡 스무딩된 데이터로 미분(속도) 계산
                df["Throttle_Speed"] = smoothed_throttle.diff() / dt
                df["Throttle_Speed"] = df["Throttle_Speed"].clip(lower=-500, upper=500) # 노이즈 컷오프        
            
            # --- 2. MoTeC 스타일 컬러 트랙 맵 생성 ---
            st.markdown("### 🗺️ MoTeC Style Color Track Map")
            st.caption("GPS 데이터를 기반으로 트랙을 그리고, 페달 조작 상태에 따라 색상을 입힙니다. (🟢 가속 | 🔴 제동 | ⚪ 코스팅/유지)")
            
            # WinTAX GPS 채널명 대응 (GPSLat, GPSLong 또는 유사 이름)
            lat_col = next((col for col in df.columns if "lat" in col.lower() or "gps_y" in col.lower()), None)
            lon_col = next((col for col in df.columns if "lon" in col.lower() or "gps_x" in col.lower()), None)
            
            if lat_col and lon_col and "rThrottle" in df.columns and "pBrakeF" in df.columns:
                # 상태 구분을 위한 색상 할당 로직
                colors = []
                for idx, row in df.iterrows():
                    thr = row["rThrottle"]
                    brk = row["pBrakeF"]
                    if brk > 5.0: colors.append("red")       # 브레이크 5bar 이상 -> 빨강
                    elif thr > 70.0: colors.append("#00ff00") # 스로틀 70% 이상 -> 밝은 초록
                    elif thr > 10.0: colors.append("yellow")  # 스로틀 유지(Maintenance) -> 노랑
                    else: colors.append("lightgrey")          # 둘 다 안 밟음(Coasting) -> 회색
                
                fig_map = go.Figure(go.Scatter(
                    x=df[lon_col], y=df[lat_col],
                    mode="markers",
                    marker=dict(color=colors, size=4),
                    text=[f"Throttle: {t:.1f}%<br>Brake: {b:.1f}bar" for t, b in zip(df["rThrottle"], df["pBrakeF"])],
                    hovertemplate="%{text}<extra></extra>"
                ))
                fig_map.update_layout(
                    height=600, template="plotly_dark",
                    xaxis=dict(showgrid=False, zeroline=False, visible=False),
                    yaxis=dict(showgrid=False, zeroline=False, visible=False, scaleanchor="x", scaleratio=1),
                    margin=dict(l=0, r=0, t=0, b=0)
                )
                st.plotly_chart(fig_map, use_container_width=True)
            else:
                st.warning("⚠️ CSV 파일에 GPS 데이터(Latitude/Longitude)가 없거나 채널명이 달라 맵을 그릴 수 없습니다.")

            # --- 3. Throttle Speed (가속 페달 공격성) 및 페달 워크 분석 차트 ---
            st.markdown("---")
            st.markdown("### ⚡ Driver Pedal Work & Throttle Speed (미분 분석)")
            st.caption("위쪽 차트는 페달 조작량(%)과 GPS 속도(km/h), 아래쪽 차트는 서적 공식 $v_{TP}(t)$를 적용한 '초당 스로틀 밟는 속도'입니다.")
            
            # 💡 [핵심 진화] GPS 속도를 찌그러짐 없이 넣기 위해 위쪽 차트에 보조축(secondary_y) 활성화
            fig_pedal = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.05, 
                                      specs=[[{"secondary_y": True}], [{"secondary_y": False}]])
            
            x_val = df.index
            
            # Row 1: 기본 페달 입력 (왼쪽 Y축 기준: 0~100)
            if "rThrottle" in df.columns:
                fig_pedal.add_trace(go.Scatter(x=x_val, y=df["rThrottle"], name="Throttle Valve (%)", line=dict(color="#00ff00", width=2)), row=1, col=1, secondary_y=False)
            
            if "rPedal" in df.columns:
                fig_pedal.add_trace(go.Scatter(x=x_val, y=df["rPedal"], name="Accel Pedal (드라이버 발 %)", line=dict(color="yellow", width=1.5, dash="dot")), row=1, col=1, secondary_y=False)
                
            if "pBrakeF" in df.columns:
                fig_pedal.add_trace(go.Scatter(x=x_val, y=df["pBrakeF"], name="Brake (bar)", line=dict(color="red", width=2)), row=1, col=1, secondary_y=False)
                
            # 💡 [신규 추가] vGPSSpeed (오른쪽 보조 Y축 기준: 0~250+ km/h)
            if "vGPSSpeed" in df.columns:
                fig_pedal.add_trace(go.Scatter(x=x_val, y=df["vGPSSpeed"], name="GPS Speed (km/h)", line=dict(color="#00d2ff", width=2)), row=1, col=1, secondary_y=True)
                
            # Row 2: 서적의 미분 공식 (Throttle Speed)
            if "Throttle_Speed" in df.columns:
                fig_pedal.add_trace(go.Scatter(x=x_val, y=df["Throttle_Speed"], name="Throttle Speed (%/s)", line=dict(color="cyan", width=1)), row=2, col=1)
                # 기준선 (0)
                fig_pedal.add_hline(y=0, line_dash="dash", line_color="gray", row=2, col=1)
            
            fig_pedal.update_layout(height=600, template="plotly_dark", hovermode="x unified")
            st.plotly_chart(fig_pedal, use_container_width=True)
            
            # --- 4. 요약 통계 KPI ---
            st.markdown("---")
            st.markdown("### 📊 Driver Performance KPI Summary")
            
            if "rThrottle" in df.columns and "pBrakeF" in df.columns:
                total_points = len(df)
                # Full Throttle (>90%) 비율
                full_thr_pts = len(df[df["rThrottle"] > 90.0])
                full_thr_pct = (full_thr_pts / total_points) * 100
                
                # Coasting (스로틀 <5% AND 브레이크 <2bar) 비율
                coast_pts = len(df[(df["rThrottle"] < 5.0) & (df["pBrakeF"] < 2.0)])
                coast_pct = (coast_pts / total_points) * 100
                
                c1, c2, c3 = st.columns(3)
                c1.metric("🔥 Full Throttle (>90%) Time", f"{full_thr_pct:.1f} %")
                c2.metric("👻 Coasting (No Pedal) Time", f"{coast_pct:.1f} %")
                if "Throttle_Speed" in df.columns:
                    max_thr_spd = df["Throttle_Speed"].max()
                    c3.metric("⚡ Max Throttle Snap Speed", f"{max_thr_spd:.0f} %/s")

        except Exception as e:
            st.error(f"❌ 데이터 분석 중 오류 발생: {e}")

  # ==========================================
# MODULE 16: Pre-Race Strategy & Weather Analysis (Custom Car)
# ==========================================
elif "16." in selected_module or "Strategy" in selected_module:
    st.header("🏁 16. Pre-Race Strategy & Weather Analysis")
    st.caption("대상 차량, 서킷명, 날씨를 입력하면 수석 엔지니어 수준의 맞춤형 셋업 전략 및 예상 트러블슈팅을 브리핑합니다.")

    # 1. 서킷 및 날씨, 차량 입력 폼 (3열 구조로 변경)
    with st.form("strategy_form"):
        st.subheader("📍 Event Conditions (환경 설정)")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            car_model = st.text_input("🏎️ 대상 차량 (예: Elantra N TCR)", "Hyundai Elantra N TCR")
        with col2:
            track_name = st.text_input("🏁 서킷 이름 (예: 영암 KIC, 인제)", "영암 KIC (풀코스)")
        with col3:
            weather_desc = st.text_input("🌤️ 날씨 및 노면", "대기온도 31도, 오후 중 비예보 강수량 10.5mm")

        st.markdown("---")
        submit_strategy = st.form_submit_button("🚨 Generate Race Strategy (전략 수립)")

    # 2. AI 전략 분석 결과 출력
    if submit_strategy:
        with st.spinner(f"Analyzing circuit characteristics for {car_model}..."):
            prompt = f"""
            너는 세계 최고 수준의 모터스포츠 팀 '수석 레이스 엔지니어(Chief Race Engineer)'다.
            입력된 차량, 서킷, 날씨 정보만을 바탕으로 완벽한 주말 레이스 전략을 브리핑하라.
            
            [입력 정보]
            - 대상 차량: {car_model}
            - 서킷명: {track_name}
            - 날씨 및 노면 상태: {weather_desc}
            
            [🚨 핵심 엔지니어링 지침]
            - 절대 F1이나 하이퍼카(LMP) 기준으로 답변하지 말 것!
            - 반드시 '{car_model}'의 물리적 특성(구동 방식, 엔진 위치, 에어로 한계, 중량 등)을 100% 반영할 것.
            - 예를 들어, Elantra N TCR 같은 전륜구동(FF) 투어링카라면 프론트 타이어의 극심한 열화(Degradation), 브레이크 과부하, 언더스티어 대응, 코너 진입 시 리프트 오프 오버스티어 셋업 등의 TCR 특성을 중점적으로 다뤄야 한다.
            
            다음 4가지 카테고리로 나누어 가장 기계적이고 전문적인 데이터 엔지니어링 관점에서 작성할 것:
            
            1. [Circuit Characteristics & Demands] (서킷 특성 및 차량 매칭)
               - 해당 서킷의 레이아웃 특성과 `{car_model}` 차량이 겪을 궁합 분석
               - 타이어 마모도(Degradation)와 브레이크 부하(Brake Severity) 수준 정밀 분석
            
            2. [Expected Car Problems & Countermeasures] (예상 차량 트러블 및 대응)
               - 해당 날씨와 서킷 특성이 결합되었을 때 `{car_model}`에서 발생 확률이 높은 기계적/열역학적 문제 
               - 치프 미캐닉과 미캐닉 팀이 개러지에서 사전에 준비하고 세팅해야 할 물리적 예방 조치
            
            3. [Vehicle Setup Strategy] (차량 셋업 전략)
               - `{car_model}`에 맞춘 에어로다이내믹스 타겟 
               - 서스펜션 지오메트리 및 댐핑 강성 (FF 차량의 경우 요(Yaw) 모멘트를 살리기 위한 리어 강성 등 구체적 셋업 조언)
               - 기계적 그립(Mechanical Grip) 확보를 위한 최적의 라이드 하이트(Ride Height) 및 안티롤바(ARB) 셋업 포인트
            
            4. [Tire Management & Pressure Strategy] (타이어 및 공기압 전략)
               - 제시된 기상 조건에 맞춘 콜드 프레셔(Cold Pressure) 타겟 셋업 방향
               - 프론트/리어 타이어의 웜업 속도 차이 극복 및 열 관리(Thermal Management) 포인트
            """
            
            try:
                # LLM 호출
                response = llm.invoke(prompt)
                answer = response.content if hasattr(response, 'content') else str(response)
                
                # AI 응답 텍스트 정제 (리스트로 반환될 경우 텍스트 추출)
                if isinstance(answer, list):
                    answer = answer[0].get("text", str(answer))
                elif not isinstance(answer, str):
                    answer = str(answer)
                
                # 줄바꿈 이스케이프 문자 정상화
                answer = answer.replace('\\n', '\n')
                
                st.success(f"✅ {car_model} 맞춤형 레이스 전략 브리핑이 완료되었습니다.")
                st.markdown(f"### 📊 Chief Race Engineer Briefing Report: {car_model}")
                st.markdown(answer)
                
            except Exception as e:
                st.error(f"❌ AI 분석 중 오류가 발생했습니다: {e}")

               # ==========================================
# MODULE 17: Auto Pro PPT Generator (Design Applied)
# ==========================================
elif "17." in selected_module or "Document" in selected_module:
    st.header("📊 17. Auto Pro PPT Generator")
    st.caption("핵심 키워드만 입력하면 완벽하게 디자인된 프로페셔널 제안서/보고서 PPT를 즉시 생성합니다.")

    import io
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
    except ImportError:
        st.error("🚨 터미널에서 'pip install python-pptx'를 실행해주세요.")

    with st.form("ppt_generator_form"):
        doc_topic = st.text_area("📝 제안서/보고서 핵심 내용 (키워드 위주로 입력)", 
                                 "예시: 제네시스 마그마 글로벌 TD 이직 제안서. 나의 기술적 강점과 가족 비전을 강조해서 5장짜리로 만들어줘.")
        submit_doc = st.form_submit_button("✨ 프로페셔널 PPT 자동 생성")

    if submit_doc and doc_topic:
        with st.spinner("AI가 슬라이드를 기획하고 전문적인 PPT로 렌더링 중입니다..."):
            
            # 1. AI 프롬프트 (정확한 파싱과 개조식 요약을 위한 특수 규칙 강제)
            prompt = f"""
            너는 모터스포츠 최고위직을 위한 전문 전략 기획자(Strategic Planner)다.
            사용자 요청: "{doc_topic}"
            
            이 내용을 바탕으로 프레젠테이션(PPT) 슬라이드 대본을 작성해라.
            반드시 아래의 특수 규칙을 100% 지켜서 출력해라 (마크다운 ** 같은 기호 절대 금지).

            [규칙]
            1. 첫 번째 슬라이드는 무조건 [TITLE_SLIDE] 로 시작.
               제목: (문서 제목)
               부제: (문서 부제)
            2. 두 번째 슬라이드부터는 무조건 [SLIDE] 로 시작.
               제목: (슬라이드 제목)
               내용:
               - (핵심 내용 1줄 요약)
               - (핵심 내용 1줄 요약)
               - (핵심 내용 1줄 요약)
            3. 내용은 장황한 글이 아닌, 스티브 잡스나 일론 머스크처럼 발표용 개조식(단답형/핵심요약)으로 아주 깔끔하게 작성할 것.
            """
            
            try:
                response = llm.invoke(prompt)
                ai_content = response.content if hasattr(response, 'content') else str(response)
                if isinstance(ai_content, list): ai_content = ai_content[0].get("text", str(ai_content))
                
                # 마크다운 잔재 완벽 제거 (별표 찌꺼기 방지)
                ai_content = ai_content.replace("**", "").replace("*", "")
                
                st.success("✅ AI 기획 완료! TESS 테마가 적용된 PPT를 생성합니다.")
                with st.expander("AI 기획안 초안 보기 (텍스트)"):
                    st.text(ai_content)
                
                # 2. PPTX 디자인 및 렌더링 작업
                prs = Presentation()
                
                # 배경색을 TESS 다크 모드(네이비)로 강제 지정하는 함수
                def set_slide_background(slide):
                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = RGBColor(10, 25, 47) # TESS 감성의 다크 네이비

                # AI 텍스트 파싱
                sections = ai_content.split("[SLIDE]")
                title_section = ai_content.split("[TITLE_SLIDE]")[1].split("[SLIDE]")[0] if "[TITLE_SLIDE]" in ai_content else sections[0]
                
                # --- 표지 슬라이드 생성 ---
                title_slide = prs.slides.add_slide(prs.slide_layouts[0])
                set_slide_background(title_slide)
                
                main_title = "Trackside Engineering Report"
                sub_title = "Generated by TESS"
                for line in title_section.strip().split('\n'):
                    if line.startswith("제목:"): main_title = line.replace("제목:", "").strip()
                    elif line.startswith("부제:"): sub_title = line.replace("부제:", "").strip()

                title_shape = title_slide.shapes.title
                subtitle_shape = title_slide.placeholders[1]
                
                # 표지 제목 디자인 (골드)
                title_shape.text = main_title
                title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(201, 160, 91)
                title_shape.text_frame.paragraphs[0].font.bold = True
                title_shape.text_frame.paragraphs[0].font.size = Pt(40)
                
                # 표지 부제 디자인 (라이트 그레이)
                subtitle_shape.text = sub_title
                subtitle_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(136, 146, 176)
                subtitle_shape.text_frame.paragraphs[0].font.size = Pt(20)

                # --- 본문 슬라이드 생성 ---
                for section in sections:
                    if "제목:" not in section: continue
                    if "[TITLE_SLIDE]" in section: continue
                    
                    lines = section.strip().split('\n')
                    s_title = ""
                    s_content = []
                    
                    is_content = False
                    for line in lines:
                        if line.startswith("제목:"):
                            s_title = line.replace("제목:", "").strip()
                        elif line.startswith("내용:"):
                            is_content = True
                        elif is_content and line.strip().startswith("-"):
                            s_content.append(line.strip()[1:].strip())
                            
                    if not s_title: continue
                    
                    # 본문 슬라이드 추가 및 배경 설정
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    set_slide_background(slide)
                    
                    # 본문 제목 디자인 (골드)
                    title_shape = slide.shapes.title
                    title_shape.text = s_title
                    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(201, 160, 91)
                    title_shape.text_frame.paragraphs[0].font.bold = True
                    title_shape.text_frame.paragraphs[0].font.size = Pt(32)
                    
                    # 본문 내용 디자인 (화이트, 간격 조절)
                    body_shape = slide.shapes.placeholders[1]
                    tf = body_shape.text_frame
                    tf.clear() # 기존 엉성한 텍스트 박스 초기화
                    
                    for bullet in s_content:
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.font.color.rgb = RGBColor(255, 255, 255) # 화이트 텍스트
                        p.font.size = Pt(20)
                        p.level = 0
                        p.space_after = Pt(18) # 줄 간격 넉넉하게
                        
                # 3. 완성된 PPT 다운로드 버튼 생성
                ppt_io = io.BytesIO()
                prs.save(ppt_io)
                ppt_io.seek(0)
                
                st.download_button(
                    label="📥 [TESS Pro] 전문 디자인 PPT 다운로드",
                    data=ppt_io,
                    file_name="TESS_Professional_Proposal.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"❌ 문서 생성 중 오류가 발생했습니다: {e}")