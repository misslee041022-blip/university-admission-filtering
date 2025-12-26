import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import os

# [New] 3번째 라이브러리: 엑셀 서식 작성을 위한 엔진 (pip install xlsxwriter)
import xlsxwriter


class UniversityFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("대입 최저학력기준 자동 필터링 시스템 (Final Ver.)")
        self.root.geometry("1100x900")

        self.df = None
        self.initial_results = None
        self.final_results = None

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=25)

        self.create_widgets()

    def create_widgets(self):
        # 1. 파일 로드
        file_frame = ttk.LabelFrame(self.root, text="1. 데이터 로드", padding=10)
        file_frame.pack(side="top", fill="x", padx=10, pady=5)

        self.lbl_file_status = ttk.Label(
            file_frame, text="엑셀 파일(.xlsx)을 불러와주세요.", foreground="red"
        )
        self.lbl_file_status.pack(side="left", padx=5)
        btn_load = ttk.Button(file_frame, text="엑셀 파일 열기", command=self.load_file)
        btn_load.pack(side="right")

        # 2. 성적 입력
        input_frame = ttk.LabelFrame(
            self.root, text="2. 수능 성적 입력 (등급)", padding=10
        )
        input_frame.pack(side="top", fill="x", padx=10, pady=5)

        self.entries = {}
        # [1줄] 국어, 수학, 영어
        ttk.Label(input_frame, text="국어:").grid(
            row=0, column=0, padx=5, pady=5, sticky="e"
        )
        self.entries["kor"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["kor"].grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(input_frame, text="수학:").grid(
            row=0, column=2, padx=5, pady=5, sticky="e"
        )
        self.entries["math"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["math"].grid(row=0, column=3, padx=2, pady=5, sticky="w")

        self.math_type = ttk.Combobox(
            input_frame, values=["미적_기하", "확통"], width=8, state="readonly"
        )
        self.math_type.current(0)
        self.math_type.grid(row=0, column=4, padx=2, pady=5, sticky="w")

        ttk.Label(input_frame, text="영어:").grid(
            row=0, column=5, padx=5, pady=5, sticky="e"
        )
        self.entries["eng"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["eng"].grid(row=0, column=6, padx=5, pady=5, sticky="w")

        # [2줄] 한국사, 탐구1, 탐구2
        ttk.Label(input_frame, text="한국사:").grid(
            row=1, column=0, padx=5, pady=5, sticky="e"
        )
        self.entries["his"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["his"].grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(input_frame, text="탐구1:").grid(
            row=1, column=2, padx=5, pady=5, sticky="e"
        )
        self.tam1_type = ttk.Combobox(
            input_frame, values=["과탐", "사탐"], width=5, state="readonly"
        )
        self.tam1_type.current(0)
        self.tam1_type.grid(row=1, column=3, padx=2, pady=5, sticky="w")
        self.entries["tam1"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["tam1"].grid(row=1, column=4, padx=2, pady=5, sticky="w")

        ttk.Label(input_frame, text="탐구2:").grid(
            row=1, column=5, padx=5, pady=5, sticky="e"
        )
        self.tam2_type = ttk.Combobox(
            input_frame, values=["과탐", "사탐"], width=5, state="readonly"
        )
        self.tam2_type.current(0)
        self.tam2_type.grid(row=1, column=6, padx=2, pady=5, sticky="w")
        self.entries["tam2"] = ttk.Entry(input_frame, width=5, justify="center")
        self.entries["tam2"].grid(row=1, column=7, padx=2, pady=5, sticky="w")

        btn_run = ttk.Button(
            input_frame,
            text="최저 충족 여부 분석 시작 🚀",
            command=self.run_primary_filter,
        )
        btn_run.grid(row=2, column=0, columnspan=8, pady=15, sticky="ew")

        # 3. 상세 필터링
        filter_frame = ttk.LabelFrame(
            self.root, text="3. 상세 조건 검색 (동적 필터링)", padding=10
        )
        filter_frame.pack(side="top", fill="x", padx=10, pady=5)

        self.var_limit = tk.StringVar(value="전체")
        self.var_cate = tk.StringVar(value="전체")
        self.var_univ = tk.StringVar(value="전체")
        self.var_type = tk.StringVar(value="전체")

        ttk.Label(filter_frame, text="① 최저유무:").pack(side="left", padx=5)
        self.cb_limit = ttk.Combobox(
            filter_frame,
            textvariable=self.var_limit,
            values=["전체", "최저있음", "최저없음"],
            state="readonly",
            width=8,
        )
        self.cb_limit.pack(side="left", padx=5)
        self.cb_limit.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="② 계열:").pack(side="left", padx=5)
        self.cb_cate = ttk.Combobox(
            filter_frame, textvariable=self.var_cate, state="readonly", width=10
        )
        self.cb_cate.pack(side="left", padx=5)
        self.cb_cate.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="③ 학교:").pack(side="left", padx=5)
        self.cb_univ = ttk.Combobox(
            filter_frame, textvariable=self.var_univ, state="readonly", width=12
        )
        self.cb_univ.pack(side="left", padx=5)
        self.cb_univ.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="④ 전형:").pack(side="left", padx=5)
        self.cb_type = ttk.Combobox(
            filter_frame, textvariable=self.var_type, state="readonly", width=12
        )
        self.cb_type.pack(side="left", padx=5)
        self.cb_type.bind("<<ComboboxSelected>>", self.on_filter_change)

        btn_reset = ttk.Button(
            filter_frame, text="필터 초기화", command=self.reset_detail_filter
        )
        btn_reset.pack(side="right", padx=10)

        # 하단 버튼 프레임
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(side="bottom", fill="x")

        self.lbl_count = ttk.Label(
            bottom_frame, text="총 0개 학과 검색됨", font=("bold", 12)
        )
        self.lbl_count.pack(side="left")

        btn_sim = ttk.Button(
            bottom_frame,
            text="📈 종합 등급 시뮬레이터 (멀티)",
            command=self.open_simulation_dialog,
        )
        btn_sim.pack(side="right", padx=5)

        # 저장 버튼 (기능 업그레이드됨)
        btn_save = ttk.Button(
            bottom_frame,
            text="결과 저장 (Excel 리포트)",
            command=self.save_excel_report,
        )
        btn_save.pack(side="right", padx=5)

        # 4. 결과 출력
        result_frame = ttk.LabelFrame(
            self.root, text="4. 분석 결과 (최저 충족 학과)", padding=10
        )
        result_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)

        columns = [
            "대학명",
            "계열",
            "모집단위",
            "전형구분",
            "최저기준",
            "50컷",
            "70컷",
            "URL",
        ]
        self.tree = ttk.Treeview(
            result_frame, columns=columns, show="headings", selectmode="browse"
        )

        col_widths = [80, 50, 150, 100, 100, 60, 60, 0]
        for col, width in zip(columns, col_widths):
            self.tree.heading(col, text=col)
            if col == "URL":
                self.tree.column(col, width=0, stretch=False)
            else:
                self.tree.column(col, width=width, anchor="center")

        scrollbar = ttk.Scrollbar(
            result_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.on_double_click)

        lbl_info = ttk.Label(
            result_frame,
            text="* 더블 클릭 시 학과 홈페이지로 이동합니다.",
            foreground="gray",
        )
        lbl_info.pack(side="bottom", anchor="w")

    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                try:
                    self.df = pd.read_excel(file_path)
                except:
                    try:
                        self.df = pd.read_csv(file_path, encoding="utf-8")
                    except:
                        self.df = pd.read_csv(file_path, encoding="cp949")
                self.df.fillna("", inplace=True)
                self.lbl_file_status.config(
                    text=f"로드 완료: {os.path.basename(file_path)}", foreground="green"
                )
                messagebox.showinfo("성공", f"데이터 {len(self.df)}건 로드 완료!")
            except Exception as e:
                messagebox.showerror("에러", f"파일 로드 실패: {e}")

    def calculate_results(self, input_scores):
        if self.df is None:
            return []
        results = []
        math_choice = self.math_type.get()
        tam1_choice = self.tam1_type.get()
        tam2_choice = self.tam2_type.get()

        for _, row in self.df.iterrows():
            req_history = (
                int(row.get("한국사", 0)) if row.get("한국사", "") != "" else 0
            )
            req_math = str(row.get("수학선택", "")).strip()
            req_tam = str(row.get("탐구선택", "")).strip()
            req_eng = str(row.get("영어필수여부", "")).strip()

            if req_history > 0 and input_scores["his"] > req_history:
                continue
            if ("미적" in req_math or "기하" in req_math) and math_choice == "확통":
                continue
            if "확통" in req_math and math_choice == "미적_기하":
                continue

            my_valid_tams = []
            is_tam1_valid = True
            if "과탐" in req_tam and tam1_choice != "과탐":
                is_tam1_valid = False
            if "사탐" in req_tam and tam1_choice != "사탐":
                is_tam1_valid = False
            if is_tam1_valid:
                my_valid_tams.append(input_scores["tam1"])

            is_tam2_valid = True
            if "과탐" in req_tam and tam2_choice != "과탐":
                is_tam2_valid = False
            if "사탐" in req_tam and tam2_choice != "사탐":
                is_tam2_valid = False
            if is_tam2_valid:
                my_valid_tams.append(input_scores["tam2"])

            reflect_tam_count = (
                int(row.get("탐구반영수", 1)) if row.get("탐구반영수", "") != "" else 1
            )
            if len(my_valid_tams) < reflect_tam_count:
                continue

            current_eng = input_scores["eng"]
            if "등급" in req_eng:
                import re

                numbers = re.findall(r"\d+", req_eng)
                if numbers:
                    limit = int(numbers[0])
                    if input_scores["eng"] > limit:
                        continue
                if "연세대" in str(row.get("대학명", "")):
                    current_eng = 99

            limit_sum = int(row.get("등급합", 0)) if row.get("등급합", "") != "" else 0
            reflect_total_count = (
                int(row.get("반영영역수", 0)) if row.get("반영영역수", "") != "" else 0
            )

            if limit_sum > 0:
                my_valid_tams.sort()
                final_tam = (
                    int(sum(my_valid_tams[:2]) / 2)
                    if reflect_tam_count == 2
                    else my_valid_tams[0]
                )
                subjects = [input_scores["kor"], input_scores["math"], final_tam]
                if current_eng != 99:
                    subjects.append(current_eng)
                subjects.sort()
                if sum(subjects[:reflect_total_count]) > limit_sum:
                    continue

            results.append(row)
        return results

    def run_primary_filter(self):
        try:
            scores = {}
            for key, ent in self.entries.items():
                val = ent.get()
                if not val:
                    raise ValueError("성적 입력")
                scores[key] = float(val)
                if not (1 <= scores[key] <= 9):
                    raise ValueError("1~9 등급 입력")
        except:
            messagebox.showwarning("오류", "성적을 올바르게 입력해주세요.")
            return

        if self.df is None:
            messagebox.showwarning("경고", "데이터 로드 필요")
            return

        self.initial_results = pd.DataFrame(self.calculate_results(scores))
        self.update_filter_options()
        self.reset_detail_filter()

    def update_filter_options(self):
        if self.initial_results is None or self.initial_results.empty:
            return
        univs = sorted(self.initial_results["대학명"].unique().tolist())
        self.cb_univ["values"] = ["전체"] + univs
        cates = sorted(self.initial_results["계열"].unique().tolist())
        self.cb_cate["values"] = ["전체"] + cates
        types = sorted(self.initial_results["전형명"].unique().tolist())
        self.cb_type["values"] = ["전체"] + types

    def on_filter_change(self, event=None):
        if self.initial_results is None:
            return
        df = self.initial_results.copy()

        if self.var_limit.get() == "최저있음":
            df = df[df["등급합"].apply(lambda x: x != "" and int(x) > 0)]
        elif self.var_limit.get() == "최저없음":
            df = df[df["등급합"].apply(lambda x: x == "" or int(x) == 0)]

        self.cb_cate["values"] = ["전체"] + sorted(df["계열"].unique().tolist())
        if self.var_cate.get() != "전체":
            df = df[df["계열"] == self.var_cate.get()]

        self.cb_univ["values"] = ["전체"] + sorted(df["대학명"].unique().tolist())
        if self.var_univ.get() != "전체":
            df = df[df["대학명"] == self.var_univ.get()]

        self.cb_type["values"] = ["전체"] + sorted(df["전형명"].unique().tolist())
        if self.var_type.get() != "전체":
            df = df[df["전형명"] == self.var_type.get()]

        self.final_results = df
        self.update_treeview()
        self.lbl_count.config(text=f"🔍 충족된 학과: {len(df)}개")

    def reset_detail_filter(self):
        self.var_limit.set("전체")
        self.var_cate.set("전체")
        self.var_univ.set("전체")
        self.var_type.set("전체")
        self.on_filter_change()

    def update_treeview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.final_results is not None:
            for _, row in self.final_results.iterrows():
                limit_text = (
                    f"{row['반영영역수']}합 {row['등급합']}"
                    if row.get("등급합", "") != "" and int(row.get("등급합", 0)) > 0
                    else "-"
                )
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("대학명", ""),
                        row.get("계열", ""),
                        row.get("모집단위", ""),
                        row.get("전형명", ""),
                        limit_text,
                        row.get("50컷", "-"),
                        row.get("70컷", "-"),
                        row.get("URL", ""),
                    ),
                )

    def on_double_click(self, event):
        item = self.tree.selection()[0]
        url = self.tree.item(item, "values")[-1]
        if url and str(url).startswith("http"):
            webbrowser.open(url)

    # [수정됨] 간결한 메시지
    def save_excel_report(self):
        if self.final_results is None or self.final_results.empty:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                self.final_results.to_excel(writer, index=False, sheet_name="분석결과")
                workbook = writer.book
                worksheet = writer.sheets["분석결과"]
                header_fmt = workbook.add_format(
                    {
                        "bold": True,
                        "text_wrap": True,
                        "valign": "top",
                        "fg_color": "#D7E4BC",
                        "border": 1,
                    }
                )
                for col_num, value in enumerate(self.final_results.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                for i, col in enumerate(self.final_results.columns):
                    max_len = (
                        max(
                            self.final_results[col].astype(str).map(len).max(), len(col)
                        )
                        + 2
                    )
                    worksheet.set_column(i, i, max_len)

            # 메시지 변경
            messagebox.showinfo("성공", "파일이 생성되었습니다.")

        except Exception as e:
            messagebox.showerror("에러", f"저장 실패: {e}")

    # ================= 시뮬레이션 =================

    def open_simulation_dialog(self):
        if self.initial_results is None:
            messagebox.showwarning(
                "알림", "먼저 현재 점수로 분석(1차 필터링)을 실행해주세요."
            )
            return

        diag = tk.Toplevel(self.root)
        diag.title("🎓 종합 성적 시뮬레이터 (가상 성적표)")
        diag.geometry("400x500")

        ttk.Label(diag, text="가정할 수능 등급을 설정하세요.", font=("bold", 12)).pack(
            pady=20
        )

        sim_entries = {}
        grid_frame = ttk.Frame(diag)
        grid_frame.pack(padx=20, pady=10)

        subjects = [
            ("국어", "kor"),
            ("수학", "math"),
            ("영어", "eng"),
            ("한국사", "his"),
            ("탐구1", "tam1"),
            ("탐구2", "tam2"),
        ]
        grade_list = [str(i) for i in range(1, 10)]

        for i, (label_text, key) in enumerate(subjects):
            ttk.Label(grid_frame, text=label_text, font=("", 10)).grid(
                row=i, column=0, padx=10, pady=8, sticky="e"
            )
            cb = ttk.Combobox(
                grid_frame,
                values=grade_list,
                width=5,
                state="readonly",
                justify="center",
            )
            cb.grid(row=i, column=1, padx=10, pady=8, sticky="w")
            try:
                val = self.entries[key].get()
                if val:
                    cb.set(str(int(float(val))))
                else:
                    cb.current(0)
            except:
                cb.current(0)
            sim_entries[key] = cb

        def run_full_sim():
            try:
                new_scores = {}
                for key, cb in sim_entries.items():
                    new_scores[key] = float(cb.get())

                sim_res = pd.DataFrame(self.calculate_results(new_scores))

                self.initial_results["ID"] = (
                    self.initial_results["대학명"]
                    + self.initial_results["모집단위"]
                    + self.initial_results["전형명"]
                )
                orig_ids = set(self.initial_results["ID"])

                if not sim_res.empty:
                    sim_res["ID"] = (
                        sim_res["대학명"] + sim_res["모집단위"] + sim_res["전형명"]
                    )
                    sim_ids = set(sim_res["ID"])
                else:
                    sim_ids = set()

                added_ids = sim_ids - orig_ids
                removed_ids = orig_ids - sim_ids

                if len(added_ids) == 0 and len(removed_ids) == 0:
                    messagebox.showinfo("결과", "변동 사항이 없습니다.")
                    return

                added_df = sim_res[sim_res["ID"].isin(added_ids)]
                removed_df = self.initial_results[
                    self.initial_results["ID"].isin(removed_ids)
                ]

                self.show_complex_sim_result(added_df, removed_df)
                diag.destroy()

            except Exception as e:
                messagebox.showerror("에러", f"오류 발생: {e}")

        ttk.Button(diag, text="시뮬레이션 분석 시작 ▶", command=run_full_sim).pack(
            pady=20
        )

    def show_complex_sim_result(self, added_df, removed_df):
        win = tk.Toplevel(self.root)
        win.title("📊 시뮬레이션 비교 분석 리포트")
        win.geometry("1100x850")

        tab_control = ttk.Notebook(win)
        tab1 = ttk.Frame(tab_control)
        tab2 = ttk.Frame(tab_control)

        tab_control.add(tab1, text=f"🎉 추가 지원 가능 (+{len(added_df)}개)")
        tab_control.add(tab2, text=f"🚨 지원 불가능 전환 (-{len(removed_df)}개)")
        tab_control.pack(expand=1, fill="both")

        def create_tab_content(parent, dataframe):
            if dataframe.empty:
                ttk.Label(parent, text="해당하는 학과가 없습니다.", font=("", 15)).pack(
                    pady=50
                )
                return

            f_frame = ttk.LabelFrame(parent, text="결과 내 필터링", padding=5)
            f_frame.pack(fill="x", padx=10, pady=5)

            v_univ = tk.StringVar(value="전체")
            v_cate = tk.StringVar(value="전체")
            v_type = tk.StringVar(value="전체")

            ttk.Label(f_frame, text="계열:").pack(side="left", padx=5)
            cb_cate = ttk.Combobox(
                f_frame, textvariable=v_cate, state="readonly", width=10
            )
            cb_cate.pack(side="left")
            ttk.Label(f_frame, text="학교:").pack(side="left", padx=5)
            cb_univ = ttk.Combobox(
                f_frame, textvariable=v_univ, state="readonly", width=12
            )
            cb_univ.pack(side="left")
            ttk.Label(f_frame, text="전형:").pack(side="left", padx=5)
            cb_type = ttk.Combobox(
                f_frame, textvariable=v_type, state="readonly", width=12
            )
            cb_type.pack(side="left")

            tree = ttk.Treeview(
                parent,
                columns=["대학", "계열", "학과", "전형", "최저", "50컷", "70컷", "URL"],
                show="headings",
            )
            cols = ["대학", "계열", "학과", "전형", "최저", "50컷", "70컷", "URL"]
            widt = [80, 50, 150, 100, 100, 60, 60, 0]
            for c, w in zip(cols, widt):
                tree.heading(c, text=c)
                if c == "URL":
                    tree.column(c, width=0, stretch=False)
                else:
                    tree.column(c, width=w, anchor="center")

            scr = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
            tree.configure(yscroll=scr.set)
            tree.pack(side="left", fill="both", expand=True, padx=10, pady=5)
            scr.pack(side="right", fill="y", pady=5)

            def on_dbl_click(event):
                item = tree.selection()[0]
                u = tree.item(item, "values")[-1]
                if u.startswith("http"):
                    webbrowser.open(u)

            tree.bind("<Double-1>", on_dbl_click)

            def update_list(event=None):
                temp = dataframe.copy()
                if v_cate.get() != "전체":
                    temp = temp[temp["계열"] == v_cate.get()]
                if v_univ.get() != "전체":
                    temp = temp[temp["대학명"] == v_univ.get()]
                if v_type.get() != "전체":
                    temp = temp[temp["전형명"] == v_type.get()]

                for i in tree.get_children():
                    tree.delete(i)
                for _, r in temp.iterrows():
                    l_txt = (
                        f"{r['반영영역수']}합 {r['등급합']}"
                        if r.get("등급합", "") != "" and int(r.get("등급합", 0)) > 0
                        else "-"
                    )
                    tree.insert(
                        "",
                        "end",
                        values=(
                            r["대학명"],
                            r["계열"],
                            r["모집단위"],
                            r["전형명"],
                            l_txt,
                            r["50컷"],
                            r["70컷"],
                            r["URL"],
                        ),
                    )

                cb_cate["values"] = ["전체"] + sorted(
                    dataframe["계열"].unique().tolist()
                )
                cb_univ["values"] = ["전체"] + sorted(temp["대학명"].unique().tolist())
                cb_type["values"] = ["전체"] + sorted(temp["전형명"].unique().tolist())

            cb_cate.bind("<<ComboboxSelected>>", update_list)
            cb_univ.bind("<<ComboboxSelected>>", update_list)
            cb_type.bind("<<ComboboxSelected>>", update_list)

            update_list()

        create_tab_content(tab1, added_df)
        create_tab_content(tab2, removed_df)


if __name__ == "__main__":
    root = tk.Tk()
    app = UniversityFilterApp(root)
    root.mainloop()
