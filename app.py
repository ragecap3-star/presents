import streamlit as st
import pandas as pd
import random
import time
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="경품 추첨 프로그램",
    page_icon="🎁",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for dark modern theme matching the original app
st.markdown("""
<style>
    .stApp {
        background-color: #10131a;
        color: white;
    }
    .big-title {
        font-size: 2.5rem;
        font-weight: 700;
        color: white;
        margin-bottom: 0px;
    }
    .prize-title {
        font-size: 3rem;
        font-weight: 800;
        color: #facc15;
        text-align: center;
        margin: 20px 0;
    }
    .status-text {
        font-size: 1.2rem;
        color: #94a3b8;
        text-align: center;
    }
    .winner-display {
        font-size: 4rem;
        font-weight: 900;
        color: #ffffff;
        text-align: center;
        background: #1e293b;
        padding: 40px;
        border-radius: 20px;
        border: 2px solid #334155;
        margin: 30px 0;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
    }
    .winner-card {
        background-color: #222222;
        border: 1px solid #333333;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state variables
if "participants" not in st.session_state:
    st.session_state.participants = []
if "remaining" not in st.session_state:
    st.session_state.remaining = []
if "winners" not in st.session_state:
    st.session_state.winners = []
if "prize_name" not in st.session_state:
    st.session_state.prize_name = "경품"
if "winner_count" not in st.session_state:
    st.session_state.winner_count = 1
if "draw_group_no" not in st.session_state:
    st.session_state.draw_group_no = 0
if "current_winner_display" not in st.session_state:
    st.session_state.current_winner_display = None

# Sidebar Controls
st.sidebar.markdown("## ⚙️ 추첨 설정 및 제어")

uploaded_file = st.sidebar.file_uploader("참가자 엑셀 업로드 (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        cols = [str(c).strip().lower() for c in df.columns]
        
        # Find name column
        name_idx = None
        for key in ["이름", "성명", "name", "성명(한글)", "회원명"]:
            if key in cols:
                name_idx = list(df.columns)[cols.index(key)]
                break
        if name_idx is None:
            name_idx = df.columns[1] if len(df.columns) >= 2 else df.columns[0]
            
        # Find ID column
        id_idx = None
        for key in ["면허번호", "id", "번호", "회원"]:
            if key in cols and list(df.columns)[cols.index(key)] != name_idx:
                id_idx = list(df.columns)[cols.index(key)]
                break
                
        participants = []
        for _, row in df.iterrows():
            name = row[name_idx]
            if pd.isna(name) or str(name).strip() == "":
                continue
            identifier = ""
            if id_idx is not None and not pd.isna(row[id_idx]):
                identifier = str(row[id_idx]).strip()
            participants.append({"name": str(name).strip(), "id": identifier})
            
        if participants and st.sidebar.button("참가자 적용하기"):
            st.session_state.participants = participants
            st.session_state.remaining = participants.copy()
            st.session_state.winners = []
            st.session_state.draw_group_no = 0
            st.sidebar.success(f"{len(participants)}명 불러오기 완료!")
    except Exception as e:
        st.sidebar.error(f"파일 읽기 오류: {e}")

st.sidebar.markdown("---")
st.session_state.prize_name = st.sidebar.text_input("경품명", value=st.session_state.prize_name)

max_rem = len(st.session_state.remaining) if st.session_state.remaining else 1
st.session_state.winner_count = st.sidebar.number_input(
    "추첨 인원 수", min_value=1, max_value=max(1, max_rem), value=min(st.session_state.winner_count, max(1, max_rem))
)

st.sidebar.markdown("---")

# Main Header
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown('<p class="big-title">🎁 경 품 추 첨 프 로 그 램</p>', unsafe_allow_html=True)
with col_h2:
    total_p = len(st.session_state.participants)
    rem_p = len(st.session_state.remaining)
    st.metric("참가자 현황", f"남은 {rem_p}명 / 전체 {total_p}명")

st.markdown(f'<div class="prize-title">{st.session_state.prize_name}</div>', unsafe_allow_html=True)

# Main Draw Actions
col_btn1, col_btn2, col_btn3 = st.columns(3)

with col_btn1:
    start_clicked = st.button("★ 추첨 START ★", use_container_width=True, type="primary")
with col_btn2:
    redraw_clicked = st.button("재추첨", use_container_width=True)
with col_btn3:
    reset_clicked = st.button("초기화", use_container_width=True)

if reset_clicked:
    st.session_state.remaining = st.session_state.participants.copy()
    st.session_state.winners = []
    st.session_state.draw_group_no = 0
    st.session_state.current_winner_display = None
    st.rerun()

# Drawing Logic
if start_clicked or redraw_clicked:
    is_redraw = redraw_clicked
    if not st.session_state.remaining:
        st.warning("추첨할 참가자가 없습니다. 엑셀 파일을 먼저 불러오주세요.")
    elif st.session_state.winner_count > len(st.session_state.remaining):
        st.warning("남은 참가자 수보다 당첨 인원이 많습니다.")
    else:
        st.session_state.draw_group_no += 1
        group_no = st.session_state.draw_group_no
        p_label = f"{st.session_state.prize_name} (재추첨)" if is_redraw else st.session_state.prize_name
        
        placeholder = st.empty()
        status_placeholder = st.empty()
        
        newly_drawn = []
        for i in range(int(st.session_state.winner_count)):
            if not st.session_state.remaining:
                break
            # Rolling animation effect
            for step in range(15):
                temp = random.choice(st.session_state.remaining)
                placeholder.markdown(f'<div class="winner-display" style="color: #94a3b8; font-size: 3rem;">🎲 {temp["name"]} {temp["id"]}</div>', unsafe_allow_html=True)
                time.sleep(0.04)
                
            winner = random.choice(st.session_state.remaining)
            st.session_state.remaining.remove(winner)
            
            st.session_state.winners.append({
                "round": len(st.session_state.winners) + 1,
                "draw_group": group_no,
                "draw_number": i + 1,
                "prize": p_label,
                "name": winner["name"],
                "id": winner["id"],
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            newly_drawn.append(winner)
            
            placeholder.markdown(f'<div class="winner-display">🎉 {winner["name"]} {winner["id"]} 🎉</div>', unsafe_allow_html=True)
            status_placeholder.markdown(f'<p class="status-text">축하합니다! 당첨되었습니다.</p>', unsafe_allow_html=True)
            time.sleep(1.5)
            
        st.success(f"이번 추첨이 완료되었습니다! ({len(newly_drawn)}명 당첨)")
        st.rerun()

# Display current/recent winner or status
if st.session_state.winners:
    latest = st.session_state.winners[-1]
    st.markdown(f"""
    <div style="background-color: #1e293b; padding: 30px; border-radius: 15px; text-align: center; border: 2px solid #facc15; margin-top: 20px;">
        <h3 style="color: #facc15; margin-bottom: 10px;">✨ 최신 당첨자</h3>
        <h1 style="font-size: 3.5rem; color: white; margin: 0;">{latest["name"]} <span style="font-size: 2rem; color: #cbd5e1;">{latest["id"]}</span></h1>
        <p style="color: #94a3b8; margin-top: 10px;">경품: {latest["prize"]} | 추첨 시각: {latest["time"]}</p>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div style="background-color: #1e293b; padding: 40px; border-radius: 15px; text-align: center; border: 1px dashed #475569; margin-top: 20px;">
        <h3 style="color: #94a3b8;">엑셀 파일을 업로드하고 [추첨 START] 버튼을 눌러주세요!</h3>
    </div>
    """, unsafe_allow_html=True)

# All Winners Section / Results Export
st.markdown("---")
st.subheader("🏆 전체 당첨자 목록 및 결과 저장")

if st.session_state.winners:
    df_winners = pd.DataFrame(st.session_state.winners)
    display_df = df_winners[["round", "prize", "name", "id", "time"]]
    display_df.columns = ["순번", "경품", "이름", "회원번호", "추첨일시"]
    st.dataframe(display_df, use_container_width=True)
    
    # Excel download
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False, sheet_name='당첨자목록')
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 당첨 결과 엑셀(XLSX)로 저장",
        data=excel_data,
        file_name=f"경품추첨결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("아직 당첨자가 없습니다.")
