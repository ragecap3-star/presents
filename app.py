import streamlit as st
import pandas as pd
import random
import time
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="경품 추첨 프로그램", page_icon="🎁", layout="wide")

st.markdown(r'''
<style>
.stApp { background:#10131a; color:white; }
.prize { text-align:center; color:#facc15; font-size:32px; font-weight:700; }
.status { text-align:center; color:#94a3b8; font-size:20px; }
.name { text-align:center; color:white; font-size:clamp(70px,10vw,150px);
font-weight:900; padding:50px 10px; min-height:220px; }
.winnerflash { text-align:center; color:white; font-size:clamp(90px,12vw,180px);
font-weight:900; padding:30px 10px; animation:flash .7s ease-out; }
@keyframes flash {
0% {transform:scale(.6);opacity:0;text-shadow:0 0 0 #fff}
35% {transform:scale(1.2);opacity:1;text-shadow:0 0 45px #fff,0 0 80px #facc15}
60% {transform:scale(1.05);text-shadow:0 0 25px #fff}
100% {transform:scale(1);text-shadow:none}
}
.card {background:#222;border:2px solid #444;border-radius:12px;padding:15px;text-align:center;margin:5px}
.cardnum {color:#aaa;font-size:20px;font-weight:bold}
.cardname {color:#fff;font-size:42px;font-weight:900;margin:15px 0}
.cardid {color:#bbb;font-size:18px}
</style>
''', unsafe_allow_html=True)

defaults = {
    "participants": [], "remaining": [], "winners": [],
    "prize_name": "경품", "winner_count": 1, "draw_group_no": 0,
    "display_name": "준비", "status": "엑셀 파일을 불러와 주세요.",
    "show_winners": False, "redraw_mode": False
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_excel(participants):
    st.session_state.participants = participants
    st.session_state.remaining = participants.copy()
    st.session_state.winners = []
    st.session_state.draw_group_no = 0
    st.session_state.display_name = "추첨 준비 완료"
    st.session_state.status = f"{len(participants)}명을 불러왔습니다."

def make_excel():
    df = pd.DataFrame([{
        "순번": w["round"], "경품": w["prize"], "이름": w["name"],
        "회원번호": w["id"], "추첨일시": w["time"]
    } for w in st.session_state.winners])
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="당첨자")
    return bio.getvalue()

def draw(count, redraw=False):
    area = st.empty()
    status_area = st.empty()
    st.session_state.draw_group_no += 1
    group_no = st.session_state.draw_group_no

    for step in range(34):
        temp = random.choice(st.session_state.remaining)
        area.markdown(
            f'<div class="name">{temp["name"]}</div>',
            unsafe_allow_html=True
        )
        time.sleep(0.035 + step * 0.003)

    for draw_index in range(count):
        winner = random.choice(st.session_state.remaining)
        st.session_state.remaining.remove(winner)

        area.markdown(
            f'<div class="winnerflash">🎉 {winner["name"]} 🎉</div>',
            unsafe_allow_html=True
        )
        status_area.markdown(
            f'<div class="status">★ 축하합니다! {st.session_state.prize_name}'
            f'{" (재추첨)" if redraw else ""} 당첨 ★</div>',
            unsafe_allow_html=True
        )
        time.sleep(0.75)

        st.session_state.winners.append({
            "round": len(st.session_state.winners) + 1,
            "draw_group": group_no,
            "draw_number": draw_index + 1,
            "prize": f'{st.session_state.prize_name} (재추첨)' if redraw else st.session_state.prize_name,
            "name": winner["name"],
            "id": winner["id"],
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

        if draw_index < count - 1:
            time.sleep(2.25)

    st.session_state.display_name = "추첨 완료"
    st.session_state.status = (
        f'추첨 완료 — 현재까지 총 {len(st.session_state.winners)}명 당첨'
    )

st.markdown("# 🎁 경 품 추 첨")
st.markdown(
    f'<div class="prize">{st.session_state.prize_name}</div>',
    unsafe_allow_html=True
)
st.markdown(
    f'<div class="status">남은 참가자 {len(st.session_state.remaining)}명</div>',
    unsafe_allow_html=True
)

st.markdown(
    f'<div class="name">{st.session_state.display_name}</div>',
    unsafe_allow_html=True
)
st.markdown(
    f'<div class="status">{st.session_state.status}</div>',
    unsafe_allow_html=True
)

st.divider()

with st.sidebar:
    st.header("⚙️ 설정")

    uploaded = st.file_uploader("엑셀 참가자 파일", type=["xlsx"])
    if uploaded and st.button("엑셀 불러오기", use_container_width=True):
        try:
            df = pd.read_excel(uploaded)
            if df.empty:
                st.error("엑셀 파일이 비어 있습니다.")
            else:
                headers = [str(x).strip().lower() for x in df.columns]

                name_idx = next(
                    (headers.index(k) for k in
                     ("이름", "성명", "name", "성명(한글)", "회원명")
                     if k in headers),
                    1 if len(headers) >= 2 else 0
                )

                id_idx = next(
                    (headers.index(k) for k in
                     ("면허번호", "id", "번호", "회원")
                     if k in headers),
                    None
                )

                if id_idx == name_idx:
                    id_idx = None

                participants = []
                for _, row in df.iterrows():
                    name = row.iloc[name_idx]
                    if pd.isna(name) or str(name).strip() == "":
                        continue

                    ident = ""
                    if id_idx is not None and not pd.isna(row.iloc[id_idx]):
                        ident = str(row.iloc[id_idx]).strip()

                    participants.append({
                        "name": str(name).strip(),
                        "id": ident
                    })

                if not participants:
                    st.error("참가자를 찾을 수 없습니다.")
                else:
                    reset_excel(participants)
                    st.rerun()

        except Exception as e:
            st.error(f"불러오기 오류: {e}")

    prize_input = st.text_input(
        "경품명",
        value=st.session_state.prize_name
    )

    if st.button("경품명 적용", use_container_width=True):
        if prize_input.strip():
            st.session_state.prize_name = prize_input.strip()
            st.rerun()

    if st.session_state.remaining:
        st.session_state.winner_count = st.number_input(
            "이번 추첨 인원",
            min_value=1,
            max_value=len(st.session_state.remaining),
            value=min(
                st.session_state.winner_count,
                len(st.session_state.remaining)
            ),
            step=1
        )

    if st.button("★ 추첨 START ★", type="primary", use_container_width=True):
        if st.session_state.remaining:
            draw(int(st.session_state.winner_count))
            st.rerun()
        else:
            st.warning("추첨할 참가자가 없습니다.")

    if st.button("🔄 재추첨", use_container_width=True):
        st.session_state.redraw_mode = True
        st.rerun()

    if st.session_state.redraw_mode and st.session_state.remaining:
        n = st.number_input(
            "재추첨 인원",
            min_value=1,
            max_value=len(st.session_state.remaining),
            value=1,
            step=1,
            key="redraw_n"
        )
        if st.button("재추첨 실행", use_container_width=True):
            draw(int(n), redraw=True)
            st.session_state.redraw_mode = False
            st.rerun()

    if st.button("🏆 전체 당첨자", use_container_width=True):
        st.session_state.show_winners = True
        st.rerun()

    if st.session_state.winners:
        st.download_button(
            "📥 결과 Excel 저장",
            data=make_excel(),
            file_name="당첨자_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if st.session_state.show_winners:
    st.markdown("## 🏆 전체 당첨자")

    groups = {}
    for w in st.session_state.winners:
        groups.setdefault(
            w.get("draw_group", w["round"]), []
        ).append(w)

    keys = sorted(groups.keys(), reverse=True)

    selected = st.selectbox(
        "추첨 그룹",
        keys,
        format_func=lambda x: f"{x}번 추첨"
    )

    winners = groups[selected]

    st.markdown(
        f'<div class="prize">{winners[0]["prize"]}</div>',
        unsafe_allow_html=True
    )

    cols = st.columns(min(len(winners), 6))

    for i, w in enumerate(winners):
        with cols[i % len(cols)]:
            st.markdown(
                f'''
                <div class="card">
                  <div class="cardnum">{w.get("draw_number", i+1)}번</div>
                  <div class="cardname">{w["name"]}</div>
                  <div class="cardid">{w["id"]}</div>
                </div>
                ''',
                unsafe_allow_html=True
            )

    if st.button("닫기"):
        st.session_state.show_winners = False
        st.rerun()
