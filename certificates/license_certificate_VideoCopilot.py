import os
import sys
import shutil
import datetime
import tkinter as tk
from tkinter import messagebox
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

def create_certificate(company, product, qty, email):
    template_path = resource_path('license_certificate/VideoCopilot/VIDEO COPILOT_origin.xlsx')
    output_dir = os.path.join(resource_root(), 'license_certificate', 'VideoCopilot')
    os.makedirs(output_dir, exist_ok=True)
    today = datetime.date.today()
    today_str = today.strftime("%Y-%m-%d")
    # 파일명 생성
    safe_company = company.replace(' ', '')
    safe_product = product.replace(' ', '')
    filename_base = f"{safe_company}_{safe_product}_{today.strftime('%Y%m%d')}"
    new_xlsx_path = os.path.join(output_dir, filename_base + '.xlsx')
    # 원본 복사
    shutil.copy(template_path, new_xlsx_path)
    # 엑셀 치환
    try:
        wb = openpyxl.load_workbook(new_xlsx_path)
        ws = wb.active
    except Exception as e:
        raise RuntimeError(f"엑셀 파일을 불러올 수 없습니다: {e}")
    if ws is not None:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = cell.value.replace('{Today}', today_str)
                    cell.value = cell.value.replace('{Company Name}', company)
                    cell.value = cell.value.replace('{Product}', product)
                    cell.value = cell.value.replace('{Qty}', qty)
                    cell.value = cell.value.replace('{Email}', email)
        wb.save(new_xlsx_path)
    else:
        raise RuntimeError("엑셀 워크시트를 찾을 수 없습니다.")
    # PDF 변환 (엑셀→PDF via 독립 Excel 인스턴스)
    excel = wc.DispatchEx('Excel.Application')
    excel.Visible = False
    wb_com = excel.Workbooks.Open(os.path.abspath(new_xlsx_path))
    pdf_path = os.path.join(output_dir, filename_base + '.pdf')
    wb_com.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # 0=PDF
    wb_com.Close(False)
    excel.Quit()
    return new_xlsx_path, pdf_path

def open_videocopilot_dialog(parent_root):
    dialog = tk.Toplevel(parent_root)
    dialog.title("Video Copilot 주문확인서 만들기")
    dialog.geometry("420x500")
    dialog.configure(bg="#F7F9FB")
    dialog.resizable(False, False)

    # 카드 프레임
    card = tk.Frame(dialog, bg="white", relief="solid", bd=1)
    card.place(relx=0.5, rely=0.5, anchor="center", width=370, height=380)
    card.grid_propagate(False)

    # 입력 항목
    labels = ["회사명", "제품명", "수량", "이메일"]
    entries = {}
    for i, label in enumerate(labels):
        tk.Label(card, text=label, font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=i, column=0, padx=10, pady=(4, 8), sticky='e')
        entry = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=22)
        entry.grid(row=i, column=1, padx=10, pady=(4, 8), sticky='w')
        entries[label] = entry

    # 버튼 프레임
    btn_frame = tk.Frame(card, bg="white")
    btn_frame.grid(row=len(labels), column=0, columnspan=2, pady=18)
    def on_create():
        company = entries["회사명"].get().strip()
        product = entries["제품명"].get().strip()
        qty = entries["수량"].get().strip()
        email = entries["이메일"].get().strip()
        if not company or not product or not qty or not email:
            messagebox.showwarning("입력 오류", "모든 항목을 입력해주세요.")
            return
        try:
            xlsx_path, pdf_path = create_certificate(company, product, qty, email)
            messagebox.showinfo("완료", f"XLSX 파일: {xlsx_path}\nPDF 파일: {pdf_path}\n\n[확인]을 누르면 폴더가 열립니다.")
            try:
                folder = os.path.dirname(pdf_path)
                os.startfile(folder)
            except Exception as e:
                messagebox.showwarning("폴더 열기 실패", f"폴더를 열 수 없습니다: {e}")
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("오류 발생", str(e))
    def style_btn(btn, bg, hover_bg):
        btn.configure(font=("맑은 고딕", 11, "bold"), fg="white", bg=bg, relief="flat", padx=20, pady=6)
        btn.bind('<Enter>', lambda e: btn.configure(bg=hover_bg))
        btn.bind('<Leave>', lambda e: btn.configure(bg=bg))
    make_btn = tk.Button(btn_frame, text="만들기", command=on_create)
    make_btn.pack(side='left', padx=10)
    style_btn(make_btn, "#3B82F6", "#2563EB")
    cancel_btn = tk.Button(btn_frame, text="취소", command=dialog.destroy)
    cancel_btn.pack(side='left', padx=10)
    style_btn(cancel_btn, "#6B7280", "#374151")

    # 사용 방법 안내 (프레임 안으로 이동)
    info_label = tk.Label(card, text="""📋 사용 방법:
1. 회사명과 제품명 등을 입력합니다.
2. '만들기' 버튼을 클릭하여 원본 파일과 PDF 파일을 생성합니다.

📁 생성된 파일은 'license_certificate/'안에 각 제품별 폴더에 저장됩니다.""", 
                         justify='left', anchor='nw', font=("맑은 고딕", 9), bg="white", wraplength=340)
    info_label.grid(row=5, column=0, columnspan=2, pady=(10, 15), padx=10, sticky='w')

    dialog.transient(parent_root)
    dialog.grab_set()
    parent_root.wait_window(dialog) 