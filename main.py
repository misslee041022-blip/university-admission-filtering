import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import os
from datetime import datetime


class UniversityFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("대입 최저학력기준 자동 필터링 시스템 (V8 - 엑셀 전용)")
        self.root.geometry("1100x800")

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
        file_frame.pack(fill="x", padx=10, pady=5)

        self.lbl_file_status = ttk.Label(
            file_frame, text="엑셀 파일(.xlsx)을 불러와주세요.", foreground="red"
        )
        self.lbl_file_status.pack(side="left", padx=5)

        btn_load = ttk.Button(file_frame, text="엑셀 파일 열기", command=self.load_file)
        btn_load.pack(side="right")

        # 2. 성적 입력
        input_frame = ttk.LabelFrame(self.root, text="2. 내 성적 입력", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)

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
            text="1차 필터링 (최저 기준 분석) 🚀",
            command=self.run_primary_filter,
        )
        btn_run.grid(row=2, column=0, columnspan=8, pady=15, sticky="ew")

        # 3. 상세 필터링
        filter_frame = ttk.LabelFrame(
            self.root, text="3. 상세 조건 검색 (동적 필터링)", padding=10
        )
        filter_frame.pack(fill="x", padx=10, pady=5)

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

        # 4. 결과 출력
        result_frame = ttk.LabelFrame(self.root, text="4. 최종 결과", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

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

        # 5. 저장 버튼
        save_frame = ttk.Frame(self.root, padding=10)
        save_frame.pack(fill="x")
        self.lbl_count = ttk.Label(
            save_frame, text="총 0개 학과 검색됨", font=("bold", 12)
        )
        self.lbl_count.pack(side="left")
        btn_save = ttk.Button(
            save_frame, text="결과 저장 (CSV)", command=self.save_file
        )
        btn_save.pack(side="right")

    def load_file(self):
        # [수정] 엑셀 파일만 선택 가능하도록 변경
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                # 1. 엑셀로 먼저 시도
                try:
                    self.df = pd.read_excel(file_path)
                except:
                    # 2. 실패 시 CSV로 시도 (확장자만 xlsx인 경우 대비)
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

    def run_primary_filter(self):
        try:
            scores = {}
            for key, ent in self.entries.items():
                val = ent.get()
                if not val:
                    raise ValueError("모든 등급을 입력해주세요.")
                scores[key] = float(val)
                if not (1 <= scores[key] <= 9):
                    raise ValueError("등급은 1~9 사이여야 합니다.")

            math_choice = self.math_type.get()
            tam1_choice = self.tam1_type.get()
            tam2_choice = self.tam2_type.get()

        except ValueError as e:
            messagebox.showwarning("입력 오류", str(e))
            return

        if self.df is None:
            messagebox.showwarning("경고", "데이터 파일을 먼저 불러와주세요.")
            return

        results = []
        for _, row in self.df.iterrows():
            req_history = (
                int(row.get("한국사", 0)) if row.get("한국사", "") != "" else 0
            )
            req_math = str(row.get("수학선택", "")).strip()
            req_tam = str(row.get("탐구선택", "")).strip()
            req_eng = str(row.get("영어필수여부", "")).strip()

            if req_history > 0 and scores["his"] > req_history:
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
                my_valid_tams.append(scores["tam1"])

            is_tam2_valid = True
            if "과탐" in req_tam and tam2_choice != "과탐":
                is_tam2_valid = False
            if "사탐" in req_tam and tam2_choice != "사탐":
                is_tam2_valid = False
            if is_tam2_valid:
                my_valid_tams.append(scores["tam2"])

            reflect_tam_count = (
                int(row.get("탐구반영수", 1)) if row.get("탐구반영수", "") != "" else 1
            )
            if len(my_valid_tams) < reflect_tam_count:
                continue

            current_eng = scores["eng"]
            if "등급" in req_eng:
                import re

                numbers = re.findall(r"\d+", req_eng)
                if numbers:
                    limit = int(numbers[0])
                    if scores["eng"] > limit:
                        continue
                if "연세대" in str(row.get("대학명", "")):
                    current_eng = 99

            limit_sum = int(row.get("등급합", 0)) if row.get("등급합", "") != "" else 0
            reflect_total_count = (
                int(row.get("반영영역수", 0)) if row.get("반영영역수", "") != "" else 0
            )

            if limit_sum > 0:
                my_valid_tams.sort()
                if reflect_tam_count == 2:
                    final_tam_score = int(sum(my_valid_tams[:2]) / 2)
                else:
                    final_tam_score = my_valid_tams[0]
                subjects = [scores["kor"], scores["math"], final_tam_score]
                if current_eng != 99:
                    subjects.append(current_eng)
                subjects.sort()
                my_sum = sum(subjects[:reflect_total_count])
                if my_sum > limit_sum:
                    continue

            results.append(row)

        self.initial_results = pd.DataFrame(results)
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
        df_pool = self.initial_results.copy()

        limit_val = self.var_limit.get()
        if limit_val == "최저있음":
            df_pool = df_pool[df_pool["등급합"].apply(lambda x: x != "" and int(x) > 0)]
        elif limit_val == "최저없음":
            df_pool = df_pool[df_pool["등급합"].apply(lambda x: x == "" or int(x) == 0)]

        valid_cates = sorted(df_pool["계열"].unique().tolist())
        self.cb_cate["values"] = ["전체"] + valid_cates
        if self.var_cate.get() not in ["전체"] + valid_cates:
            self.var_cate.set("전체")

        cate_val = self.var_cate.get()
        if cate_val != "전체":
            df_pool = df_pool[df_pool["계열"] == cate_val]

        valid_univs = sorted(df_pool["대학명"].unique().tolist())
        self.cb_univ["values"] = ["전체"] + valid_univs
        if self.var_univ.get() not in ["전체"] + valid_univs:
            self.var_univ.set("전체")

        univ_val = self.var_univ.get()
        if univ_val != "전체":
            df_pool = df_pool[df_pool["대학명"] == univ_val]

        valid_types = sorted(df_pool["전형명"].unique().tolist())
        self.cb_type["values"] = ["전체"] + valid_types
        if self.var_type.get() not in ["전체"] + valid_types:
            self.var_type.set("전체")

        type_val = self.var_type.get()
        if type_val != "전체":
            df_pool = df_pool[df_pool["전형명"] == type_val]

        self.final_results = df_pool
        self.update_treeview()
        self.lbl_count.config(text=f"🔍 조건에 맞는 학과: {len(df_pool)}개")

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
                limit_text = "-"
                if row.get("등급합", "") != "" and int(row.get("등급합", 0)) > 0:
                    limit_text = f"{row['반영영역수']}합 {row['등급합']}"
                type_name = row.get("전형명", "기타")
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("대학명", ""),
                        row.get("계열", ""),
                        row.get("모집단위", ""),
                        type_name,
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
        else:
            messagebox.showinfo("알림", "홈페이지 링크가 없습니다.")

    def save_file(self):
        if self.final_results is None or self.final_results.empty:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV file", "*.csv")]
        )
        if file_path:
            try:
                self.final_results.to_csv(file_path, index=False, encoding="utf-8-sig")
                messagebox.showinfo("완료", "파일 저장 완료!")
            except Exception as e:
                messagebox.showerror("에러", f"저장 실패: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = UniversityFilterApp(root)
    root.mainloop()
