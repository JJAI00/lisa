import os
import sys
import shutil
import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext
import openpyxl
import win32com.client as wc

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

def resource_root():
    import sys, os
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath('.')

def create_certificate(company, product, license_type, qty, email):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = resource_path('license_certificate/Marmoset/Marmoset_origin.xlsx')
    output_dir = os.path.join(resource_root(), 'license_certificate', 'Marmoset')
    os.makedirs(output_dir, exist_ok=True)
    today = datetime.date.today()
    today_str = today.strftime("%Y-%m-%d")
    # 파일명 생성
    safe_company = company.replace(' ', '')
    safe_product = product.replace(' ', '')
    filename_base = f"Marmoset_{safe_product}_{safe_company}_{today.strftime('%Y%m%d')}"
    new_xlsx_path = os.path.join(output_dir, filename_base + '.xlsx')
    # 원본 복사
    shutil.copy(template_path, new_xlsx_path)
    # 엑셀 치환
    wb = openpyxl.load_workbook(new_xlsx_path)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value.replace('{Today}', today_str)
                cell.value = cell.value.replace('{Company Name}', company)
                cell.value = cell.value.replace('{Product}', product)
                cell.value = cell.value.replace('{License Type}', license_type)
                cell.value = cell.value.replace('{Qty}', qty)
                cell.value = cell.value.replace('{Email}', email)
    wb.save(new_xlsx_path)
    # PDF 변환 (엑셀→PDF via 독립 Excel 인스턴스)
    excel = wc.DispatchEx('Excel.Application')
    excel.Visible = False
    wb_com = excel.Workbooks.Open(os.path.abspath(new_xlsx_path))
    pdf_path = os.path.join(output_dir, filename_base + '.pdf')
    wb_com.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # 0=PDF
    wb_com.Close(False)
    excel.Quit()
    return new_xlsx_path, pdf_path

def open_marmoset_dialog(parent_root):
    dialog = tk.Toplevel(parent_root)
    dialog.title("Marmoset 인증서 만들기")
    dialog.geometry("500x600")
    dialog.configure(bg="#F7F9FB")
    dialog.resizable(False, False)
    # 카드 프레임
    card = tk.Frame(dialog, bg="white", relief="solid", bd=1)
    card.place(relx=0.5, rely=0.5, anchor="center", width=460, height=520)
    card.grid_propagate(False)
    # 회사명
    tk.Label(card, text="회사명", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=0, column=0, padx=10, pady=10, sticky='e')
    entry_company = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=28)
    entry_company.grid(row=0, column=1, padx=10, pady=10, sticky='w')
    # 제품명
    tk.Label(card, text="제품명", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=1, column=0, padx=10, pady=10, sticky='e')
    entry_product = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=28)
    entry_product.grid(row=1, column=1, padx=10, pady=10, sticky='w')
    # 라이선스 타입
    tk.Label(card, text="라이선스 타입", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=2, column=0, padx=10, pady=10, sticky='e')
    license_type_var = tk.StringVar(value="Annual Subscription")
    radio_frame = tk.Frame(card, bg="white")
    radio_frame.grid(row=2, column=1, padx=10, pady=10, sticky='w')
    tk.Radiobutton(radio_frame, text="Annual Subscription", variable=license_type_var, value="Annual Subscription", font=("맑은 고딕", 10), bg="white").pack(side='left', padx=5)
    tk.Radiobutton(radio_frame, text="Perpetual", variable=license_type_var, value="Perpetual", font=("맑은 고딕", 10), bg="white").pack(side='left', padx=5)
    # 수량
    tk.Label(card, text="수량", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=3, column=0, padx=10, pady=10, sticky='e')
    entry_qty = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=28)
    entry_qty.grid(row=3, column=1, padx=10, pady=10, sticky='w')
    # 이메일
    tk.Label(card, text="이메일", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=4, column=0, padx=10, pady=10, sticky='e')
    entry_email = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=28)
    entry_email.grid(row=4, column=1, padx=10, pady=10, sticky='w')
    # 버튼 프레임
    btn_frame = tk.Frame(card, bg="white")
    btn_frame.grid(row=5, column=0, columnspan=2, pady=18)
    def on_create():
        company = entry_company.get().strip()
        product = entry_product.get().strip()
        license_type = license_type_var.get()
        qty = entry_qty.get().strip()
        email = entry_email.get().strip()
        if not company or not product or not qty or not email:
            messagebox.showwarning("입력 오류", "모든 항목을 입력해주세요.")
            return
        try:
            xlsx_path, pdf_path = create_certificate(company, product, license_type, qty, email)
            messagebox.showinfo("완료", f"XLSX 파일: {xlsx_path}\nPDF 파일: {pdf_path}\n\n[확인]을 누르면 폴더가 열립니다.")
            try:
                folder = os.path.dirname(pdf_path)
                os.startfile(folder)
            except Exception as e:
                messagebox.showwarning("폴더 열기 실패", f"폴더를 열 수 없습니다: {e}")
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("오류 발생", str(e))
    create_btn = tk.Button(btn_frame, text="만들기", font=("맑은 고딕", 11, "bold"), bg="#3B82F6", fg="white", relief="flat", activebackground="#2563EB", activeforeground="white", command=on_create, width=10, height=1)
    create_btn.pack(side='left', padx=10)
    create_btn.bind("<Enter>", lambda e: create_btn.config(bg="#2563EB"))
    create_btn.bind("<Leave>", lambda e: create_btn.config(bg="#3B82F6"))
    cancel_btn = tk.Button(btn_frame, text="취소", font=("맑은 고딕", 11, "bold"), bg="#94A3B8", fg="white", relief="flat", activebackground="#64748B", activeforeground="white", command=dialog.destroy, width=10, height=1)
    cancel_btn.pack(side='left', padx=10)
    cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg="#64748B"))
    cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg="#94A3B8"))
    # 사용 방법 안내 텍스트 (카드 하단)
    info_label = tk.Label(card, text="""📋 사용 방법:
1. 회사명과 제품명 등을 입력합니다.
2. '만들기' 버튼을 클릭하여 원본 파일과 PDF 파일을 생성합니다.

📁 생성된 파일은 'license_certificate/'안에 각 제품별 폴더에 저장됩니다.""", 
                         justify='left', anchor='nw', font=("맑은 고딕", 9), bg="white", wraplength=440)
    info_label.grid(row=6, column=0, columnspan=2, sticky='nw', padx=10, pady=(10, 10))
    dialog.transient(parent_root)
    dialog.grab_set()
    parent_root.wait_window(dialog) 