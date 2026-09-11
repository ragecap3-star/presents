# -*- coding: utf-8 -*-
import random
import threading
import time
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
try:
    from openpyxl import load_workbook, Workbook
except ImportError:
    raise SystemExit("openpyxl이 필요합니다. 'pip install openpyxl'을 실행하세요.")


class LotteryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("경품 추첨 프로그램")
        self.root.geometry("1280x720")
        self.root.minsize(900, 600)
        self.root.configure(bg="#10131a")

        self.participants = []
        self.remaining = []
        self.winners = []

        self.prize_name = "경품"
        self.winner_count = 1
        self.is_drawing = False

        # 한 번의 추첨 작업(START 또는 재추첨)을 구분하는 번호
        self.draw_group_no = 0

        self.build_ui()

    def build_ui(self):
        top = tk.Frame(self.root, bg="#10131a")
        top.pack(fill="x", padx=30, pady=(25, 10))

        tk.Label(
            top, text="🎁 경 품 추 첨",
            font=("Malgun Gothic", 30, "bold"),
            fg="white", bg="#10131a"
        ).pack(side="left")

        self.count_label = tk.Label(
            top, text="참가자 0명",
            font=("Malgun Gothic", 15),
            fg="#cbd5e1", bg="#10131a"
        )
        self.count_label.pack(side="right", pady=8)

        center = tk.Frame(self.root, bg="#10131a")
        center.pack(fill="both", expand=True, padx=50, pady=10)

        self.prize_label = tk.Label(
            center, text="경품을 입력하세요",
            font=("Malgun Gothic", 36, "bold"),
            fg="#facc15", bg="#10131a"
        )
        self.prize_label.pack(pady=(20, 10))

        self.status_label = tk.Label(
            center, text="엑셀 파일을 불러와 주세요.",
            font=("Malgun Gothic", 18),
            fg="#94a3b8", bg="#10131a"
        )
        self.status_label.pack(pady=5)

        self.name_label = tk.Label(
            center, text="준비",
            font=("Malgun Gothic", 128, "bold"),
            fg="white", bg="#10131a"
        )
        self.name_label.pack(expand=True)

        self.result_label = tk.Label(
            center, text="",
            font=("Malgun Gothic", 20),
            fg="#facc15", bg="#10131a"
        )
        self.result_label.pack(pady=5)

        controls = tk.Frame(self.root, bg="#171b24")
        controls.pack(fill="x", padx=30, pady=20)

        self.make_button(controls, "엑셀 불러오기", self.load_excel).pack(side="left", padx=6, pady=12)
        self.make_button(controls, "경품명 설정", self.set_prize).pack(side="left", padx=6, pady=12)
        self.make_button(controls, "당첨 인원", self.set_winner_count).pack(side="left", padx=6, pady=12)

        self.start_button = self.make_button(
            controls, "★ 추첨 START ★", self.start_draw, primary=True
        )
        self.start_button.pack(side="left", padx=12, pady=12)

        # 요청하신 버튼을 하단 메뉴에 확실하게 추가
        self.all_winner_button = self.make_button(
            controls, "🏆 전체 당첨자", self.show_all_winners
        )
        self.all_winner_button.pack(side="left", padx=6, pady=12)

        self.make_button(
            controls, "재추첨", self.redraw
        ).pack(side="left", padx=6, pady=12)

        self.make_button(
            controls, "결과 저장", self.save_results
        ).pack(side="left", padx=6, pady=12)

        self.make_button(
            controls, "전체화면 F11", self.toggle_fullscreen
        ).pack(side="right", padx=6, pady=12)

        self.fullscreen = False
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())
        self.root.bind("<Escape>", lambda e: self.exit_fullscreen())

    def make_button(self, parent, text, command, primary=False):
        return tk.Button(
            parent, text=text, command=command,
            font=("Malgun Gothic", 13, "bold"),
            padx=14, pady=10,
            bg="#2563eb" if primary else "#263244",
            fg="white",
            activebackground="#3b82f6",
            activeforeground="white",
            relief="flat",
            cursor="hand2"
        )

    def load_excel(self):
        path = filedialog.askopenfilename(
            title="참가자 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")]
        )
        if not path:
            return

        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))

            if not rows:
                raise ValueError("엑셀 파일이 비어 있습니다.")

            headers = [str(x).strip().lower() if x is not None else "" for x in rows[0]]

            # 이름 열 자동 탐색
            name_idx = None
            for key in ("이름", "성명", "name", "성명(한글)", "회원명"):
                if key in headers:
                    name_idx = headers.index(key)
                    break
            if name_idx is None:
                name_idx = 1 if len(headers) >= 2 else 0

            # 회원번호/번호 열 자동 탐색
            id_idx = None
            for key in ("면허번호", "면허번호", "id", "번호", "회원", "면허번호"):
                if key in headers:
                    id_idx = headers.index(key)
                    break
            if id_idx == name_idx:
                id_idx = None

            participants = []
            for row in rows[1:]:
                if not row or name_idx >= len(row):
                    continue
                name = row[name_idx]
                if name is None or str(name).strip() == "":
                    continue

                identifier = ""
                if id_idx is not None and id_idx < len(row) and row[id_idx] is not None:
                    identifier = str(row[id_idx]).strip()

                participants.append({"name": str(name).strip(), "id": identifier})

            if not participants:
                raise ValueError("참가자를 찾을 수 없습니다.")

            self.participants = participants
            self.remaining = self.participants.copy()
            self.winners = []
            self.draw_group_no = 0

            self.count_label.config(text=f"참가자 {len(participants)}명")
            self.name_label.config(text="추첨 준비 완료")
            self.status_label.config(
                text=f"'{Path(path).name}'에서 {len(participants)}명을 불러왔습니다."
            )
            self.result_label.config(text="")

        except Exception as e:
            messagebox.showerror("불러오기 오류", str(e))

    def set_prize(self):
        value = simpledialog.askstring(
            "경품명", "경품명을 입력하세요:", initialvalue=self.prize_name
        )
        if value and value.strip():
            self.prize_name = value.strip()
            self.prize_label.config(text=self.prize_name)

    def set_winner_count(self):
        max_count = len(self.remaining)
        if max_count < 1:
            messagebox.showinfo("당첨 인원", "추첨 가능한 참가자가 없습니다.")
            return

        value = simpledialog.askinteger(
            "당첨 인원",
            f"이번 추첨에서 뽑을 인원\n현재 추첨 가능: {max_count}명",
            initialvalue=min(self.winner_count, max_count),
            minvalue=1,
            maxvalue=max_count
        )
        if value:
            self.winner_count = value
            self.result_label.config(text=f"이번 추첨: {value}명")

    def start_draw(self):
        if self.is_drawing:
            return

        if not self.remaining:
            messagebox.showwarning("알림", "추첨할 참가자가 없습니다.")
            return

        if self.winner_count > len(self.remaining):
            messagebox.showwarning(
                "알림",
                f"남은 참가자가 {len(self.remaining)}명입니다."
            )
            return

        self.is_drawing = True
        self.start_button.config(state="disabled")

        # 이번 START 추첨 전체에 하나의 그룹 번호 부여
        self.draw_group_no += 1
        group_no = self.draw_group_no

        threading.Thread(
            target=self.draw_multiple,
            args=(group_no,),
            daemon=True
        ).start()

    def draw_multiple(self, group_no):
        try:
            for draw_index in range(self.winner_count):
                winner = self.draw_one()

                if winner:
                    self.winners.append({
                        # 전체 당첨 순번
                        "round": len(self.winners) + 1,

                        # 이번 추첨 그룹
                        "draw_group": group_no,

                        # 이번 추첨에서 몇 번째인지
                        "draw_number": draw_index + 1,

                        "prize": self.prize_name,
                        "name": winner["name"],
                        "id": winner["id"],
                        "time": datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    })

                time.sleep(3)

            self.root.after(0, self.finish_draw)

        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "오류",
                    f"추첨 중 오류가 발생했습니다.\n{e}"
                )
            )
            self.root.after(0, self.finish_draw)

    def draw_one(self):
        # 현재 remaining 안에서만 롤링
        for step in range(34):
            if not self.remaining:
                return None

            temp = random.choice(self.remaining)
            self.root.after(0, lambda t=temp: self.name_label.config(text=t["name"]))
            time.sleep(0.035 + step * 0.003)

        winner = random.choice(self.remaining)
        self.remaining.remove(winner)

        self.root.after(0, lambda w=winner: self.show_winner(w))
        return winner

    # 원본 파일의 show_winner() 함수만 아래 코드로 교체하세요.

    def show_winner(self, winner):
        # 당첨 확정 순간 크게 + 밝게
        self.name_label.config(
            text=f"🎉 {winner['name']} {winner['id']} 🎉",
            font=("Malgun Gothic", 170, "bold"),
            fg="white"
        )

        self.result_label.config(
            text=f"★ 축하합니다! {self.prize_name} 당첨 ★"
        )

        self.status_label.config(
            text="당첨자 확정"
        )

        # 0.08초 후 살짝 밝은 빛 번짐
        self.root.after(
            80,
            lambda: self.name_label.config(
                font=("Malgun Gothic", 180, "bold"),
                fg="#FFFFCC"
            )
        )

        # 0.18초 후 원래 밝기로 돌아오면서 크기도 약간 축소
        self.root.after(
            280,
            lambda: self.name_label.config(
                font=("Malgun Gothic", 128, "bold"),
                fg="white"
            )
        )

        # 원래 크기로 돌아온 후 3초 유지
        self.root.after(
            280,
            lambda: self.status_label.config(
                text="당첨자 확정 — 다음 추첨까지 3초간 유지합니다."
            )
        )

    def finish_draw(self):
        self.is_drawing = False
        self.start_button.config(state="normal")
        self.count_label.config(text=f"남은 참가자 {len(self.remaining)}명")
        self.status_label.config(
            text=f"추첨 완료 — 현재까지 총 {len(self.winners)}명 당첨"
        )

    def redraw(self):
        if self.is_drawing:
            return

        if not self.remaining:
            messagebox.showwarning(
                "알림",
                "재추첨할 수 있는 참가자가 없습니다."
            )
            return

        count = simpledialog.askinteger(
            "재추첨",
            "재추첨할 인원 수를 입력하세요.",
            minvalue=1,
            maxvalue=len(self.remaining)
        )

        if not count:
            return

        self.is_drawing = True

        # 재추첨도 하나의 독립적인 추첨 그룹으로 처리
        self.draw_group_no += 1
        group_no = self.draw_group_no

        threading.Thread(
            target=self.draw_redraw,
            args=(group_no, count),
            daemon=True
        ).start()

    def draw_redraw(self, group_no, count):
        try:
            for draw_index in range(count):
                winner = self.draw_one()

                if winner:
                    self.winners.append({
                        # 전체 당첨 순번
                        "round": len(self.winners) + 1,

                        # 재추첨 그룹
                        "draw_group": group_no,

                        # 이번 재추첨에서 몇 번째인지
                        "draw_number": draw_index + 1,

                        "prize": f"{self.prize_name} (재추첨)",
                        "name": winner["name"],
                        "id": winner["id"],
                        "time": datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    })

                time.sleep(3)

            self.root.after(0, self.finish_draw)

        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "오류",
                    f"재추첨 중 오류가 발생했습니다.\n{e}"
                )
            )
            self.root.after(0, self.finish_draw)

    def show_all_winners(self):
        if not self.winners:
            messagebox.showinfo(
                "전체 당첨자",
                "아직 당첨자가 없습니다."
            )
            return

        # =================================================
        # 1. 추첨 그룹별로 당첨자 분류
        # =================================================
        groups = {}

        for winner in self.winners:
            group_no = winner.get("draw_group")

            # 예전 데이터와의 호환
            if group_no is None:
                group_no = winner.get("round", 1)

            if group_no not in groups:
                groups[group_no] = []

            groups[group_no].append(winner)

        # 추첨 그룹 순서
        group_keys = sorted(groups.keys())

        # 최신 추첨부터 보여주기
        current_page = [len(group_keys) - 1]

        # =================================================
        # 전체화면 창
        # =================================================
        win = tk.Toplevel(self.root)
        win.title("🏆 전체 당첨자")
        win.configure(bg="#111111")

        # 전체화면
        win.attributes("-fullscreen", True)

        # =================================================
        # 상단 제목
        # =================================================
        title_label = tk.Label(
            win,
            text="🏆 전체 당첨자",
            font=("맑은 고딕", 28, "bold"),
            fg="white",
            bg="#111111"
        )
        title_label.pack(pady=(12, 2))

        # 페이지 표시
        page_label = tk.Label(
            win,
            text="",
            font=("맑은 고딕", 14, "bold"),
            fg="#aaaaaa",
            bg="#111111"
        )
        page_label.pack(pady=(0, 4))

        # =================================================
        # 당첨자 표시 영역
        # =================================================
        content_frame = tk.Frame(
            win,
            bg="#111111"
        )
        content_frame.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=(2, 2)
        )

        # =================================================
        # 하단 버튼
        # =================================================
        button_frame = tk.Frame(
            win,
            bg="#111111"
        )
        button_frame.pack(
            fill="x",
            side="bottom",
            pady=(5, 15)
        )

        # -------------------------------------------------
        # 페이지 표시 함수
        # -------------------------------------------------
        def show_page():

            # 기존 화면 삭제
            for widget in content_frame.winfo_children():
                widget.destroy()

            group_no = group_keys[current_page[0]]
            winners = groups[group_no]

            count = len(winners)

            # =================================================
            # 페이지 정보
            # =================================================
            page_label.config(
                text=f"{current_page[0] + 1} / {len(group_keys)}"
            )

            # =================================================
            # 경품명
            # =================================================
            first_winner = winners[0]

            prize_name = first_winner.get(
                "prize",
                self.prize_name
            )

            prize_label = tk.Label(
                content_frame,
                text=prize_name,
                font=("맑은 고딕", 20, "bold"),
                fg="#ffd700",
                bg="#111111"
            )
            prize_label.pack(
                pady=(0, 5)
            )

            # =================================================
            # 당첨자 수에 따른 자동 레이아웃
            # =================================================

            if count == 1:
                columns = 1
                name_font_size = 200
                number_font_size = 24
                id_font_size = 150

            elif count == 2:
                columns = 2
                name_font_size = 120
                number_font_size = 20
                id_font_size = 100

            elif count == 3:
                columns = 3
                name_font_size = 110
                number_font_size = 18
                id_font_size = 90

            elif count == 4:
                columns = 2
                name_font_size = 100
                number_font_size = 17
                id_font_size = 80

            elif count <= 6:
                columns = 3
                name_font_size = 90
                number_font_size = 16
                id_font_size = 70

            elif count <= 8:
                columns = 4
                name_font_size = 80
                number_font_size = 15
                id_font_size = 65

            elif count <= 10:
                columns = 5
                name_font_size = 80
                number_font_size = 14
                id_font_size = 65

            elif count <= 12:
                columns = 4
                name_font_size = 75
                number_font_size = 14
                id_font_size = 65

            elif count <= 15:
                columns = 5
                name_font_size = 70
                number_font_size = 13
                id_font_size = 60

            elif count <= 20:
                columns = 5
                name_font_size = 60
                number_font_size = 13
                id_font_size = 45

            elif count <= 24:
                columns = 6
                name_font_size = 54
                number_font_size = 12
                id_font_size = 45

            elif count <= 30:
                columns = 6
                name_font_size = 52
                number_font_size = 12
                id_font_size = 42

            elif count <= 40:
                columns = 8
                name_font_size = 23
                number_font_size = 11
                id_font_size = 13

            else:
                columns = 10
                name_font_size = 20
                number_font_size = 10
                id_font_size = 9

            rows = (count + columns - 1) // columns

            # =================================================
            # 당첨자 Grid
            # =================================================
            grid_frame = tk.Frame(
                content_frame,
                bg="#111111"
            )
            grid_frame.pack(
                fill="both",
                expand=True
            )

            # 모든 열을 동일한 크기로
            for col in range(columns):
                grid_frame.columnconfigure(
                    col,
                    weight=1
                )

            # 모든 행을 동일한 크기로
            for row in range(rows):
                grid_frame.rowconfigure(
                    row,
                    weight=1
                )

            # =================================================
            # 당첨자 표시
            # =================================================
            for index, winner in enumerate(winners):

                row = index // columns
                col = index % columns

                # -------------------------------------------------
                # 카드
                # -------------------------------------------------
                card = tk.Frame(
                    grid_frame,
                    bg="#222222",
                    bd=2,
                    relief="ridge"
                )

                card.grid(
                    row=row,
                    column=col,
                    padx=4,
                    pady=4,
                    sticky="nsew"
                )

                # -------------------------------------------------
                # 이번 추첨의 당첨번호
                # -------------------------------------------------
                draw_number = winner.get(
                    "draw_number",
                    index + 1
                )

                number_label = tk.Label(
                    card,
                    text=f"{draw_number}번",
                    font=(
                        "맑은 고딕",
                        number_font_size,
                        "bold"
                    ),
                    fg="#aaaaaa",
                    bg="#222222"
                )

                number_label.pack(
                    pady=(5, 0)
                )

                # -------------------------------------------------
                # 이름
                # -------------------------------------------------
                name_label = tk.Label(
                    card,
                    text=winner.get("name", ""),
                    font=(
                        "맑은 고딕",
                        name_font_size,
                        "bold"
                    ),
                    fg="white",
                    bg="#222222",
                    anchor="center"
                )

                name_label.pack(
                    fill="both",
                    expand=True,
                    padx=3,
                    pady=1
                )

                # -------------------------------------------------
                # 회원번호
                # -------------------------------------------------
                member_id = winner.get("id", "")

                id_label = tk.Label(
                    card,
                    text=str(member_id),
                    font=(
                        "맑은 고딕",
                        id_font_size
                    ),
                    fg="#bbbbbb",
                    bg="#222222"
                )

                id_label.pack(
                    pady=(0, 5)
                )

            # =================================================
            # 이전 / 다음 버튼 상태
            # =================================================
            if current_page[0] <= 0:
                prev_button.config(
                    state="disabled"
                )
            else:
                prev_button.config(
                    state="normal"
                )

            if current_page[0] >= len(group_keys) - 1:
                next_button.config(
                    state="disabled"
                )
            else:
                next_button.config(
                    state="normal"
                )

        # =================================================
        # 이전 페이지
        # =================================================
        def previous_page():
            if current_page[0] > 0:
                current_page[0] -= 1
                show_page()

        # =================================================
        # 다음 페이지
        # =================================================
        def next_page():
            if current_page[0] < len(group_keys) - 1:
                current_page[0] += 1
                show_page()

        # =================================================
        # 창 닫기
        # =================================================
        def close_window():
            win.destroy()

        # =================================================
        # 버튼
        # =================================================
        prev_button = tk.Button(
            button_frame,
            text="◀ 이전",
            font=("맑은 고딕", 16, "bold"),
            width=9,
            height=1,
            command=previous_page
        )

        prev_button.pack(
            side="left",
            padx=15
        )

        next_button = tk.Button(
            button_frame,
            text="다음 ▶",
            font=("맑은 고딕", 16, "bold"),
            width=9,
            height=1,
            command=next_page
        )

        next_button.pack(
            side="left",
            padx=15
        )

        close_button = tk.Button(
            button_frame,
            text="닫기",
            font=("맑은 고딕", 16, "bold"),
            width=9,
            height=1,
            command=close_window
        )

        close_button.pack(
            side="right",
            padx=15
        )

        # =================================================
        # 키보드 조작
        # =================================================

        # 왼쪽 화살표
        win.bind(
            "<Left>",
            lambda event: previous_page()
        )

        # 오른쪽 화살표
        win.bind(
            "<Right>",
            lambda event: next_page()
        )

        # ESC
        win.bind(
            "<Escape>",
            lambda event: close_window()
        )

        # =================================================
        # 첫 화면 표시
        # =================================================
        show_page()
   


    def save_results(self):
        if not self.winners:
            messagebox.showinfo("저장", "아직 당첨자가 없습니다.")
            return

        path = filedialog.asksaveasfilename(
            title="당첨 결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "당첨자"
            ws.append(["순번", "경품", "이름", "회원번호", "추첨일시"])

            for w in self.winners:
                ws.append([
                    w["round"], w["prize"], w["name"], w["id"], w["time"]
                ])

            wb.save(path)
            messagebox.showinfo("저장 완료", f"당첨자 결과를 저장했습니다.\n{path}")
        except Exception as e:
            messagebox.showerror("저장 오류", str(e))

    def toggle_fullscreen(self):
        self.fullscreen = not self.fullscreen
        self.root.attributes("-fullscreen", self.fullscreen)

    def exit_fullscreen(self):
        self.fullscreen = False
        self.root.attributes("-fullscreen", False)


if __name__ == "__main__":
    root = tk.Tk()
    app = LotteryApp(root)
    root.mainloop()
