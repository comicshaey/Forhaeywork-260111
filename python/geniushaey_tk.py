# geniushaey_tk.py
# 교육공무직 연차계산기 (대충 미완)

import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date
import core


def run_calculation():
    try:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        emp = core.Employee(
            name="직원",
            first_hire_date=date.fromisoformat(entry_hire.get()),
            period_type=period_var.get(),
            schedule_key=entry_denom.get()
        )

        worklog = core.read_worklog(path)
        result = core.calculate_annual_leave(emp, worklog, int(entry_year.get()))

        output.delete("1.0", tk.END)
        for k, v in result.items():
            # 80%미만_설명은 80% 이상일 때 빈 문자열이라 출력 스킵해도 됨
            if k == "80%미만_설명" and not str(v).strip():
                continue
            output.insert(tk.END, f"{k}: {v}\n")

    except Exception as e:
        messagebox.showerror("오류", str(e))


root = tk.Tk()
root.title("교육공무직 연차계산기 (대충 미완)")

tk.Label(root, text="일 좀 편하게 해보자 힘들다…", font=("맑은 고딕", 10)).grid(
    row=0, column=0, columnspan=2, pady=(5, 10)
)

tk.Label(root, text="부여연도").grid(row=1, column=0, sticky="e")
entry_year = tk.Entry(root)
entry_year.insert(0, "2026")
entry_year.grid(row=1, column=1)

tk.Label(root, text="최초임용일 (YYYY-MM-DD)").grid(row=2, column=0, sticky="e")
entry_hire = tk.Entry(root)
entry_hire.grid(row=2, column=1)

tk.Label(root, text="분모키").grid(row=3, column=0, sticky="e")
entry_denom = tk.Entry(root)
entry_denom.insert(0, "FULLTIME_260")
entry_denom.grid(row=3, column=1)

period_var = tk.StringVar(value="SCHOOL_YEAR")
tk.Radiobutton(root, text="학교형", variable=period_var, value="SCHOOL_YEAR").grid(row=4, column=0)
tk.Radiobutton(root, text="기관형", variable=period_var, value="CALENDAR_YEAR").grid(row=4, column=1)

tk.Button(root, text="근무상황목록 선택 & 계산", command=run_calculation).grid(
    row=5, column=0, columnspan=2, pady=10
)

output = tk.Text(root, height=14, width=60)
output.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

root.mainloop()
