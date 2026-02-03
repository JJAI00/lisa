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

# Red Giant용
try:
    from pptx import Presentation
except ImportError:
    Presentation = None  # python-pptx 미설치시 예외 처리 (pip install python-pptx)


def create_excel_certificate(template_path, output_dir, company, product1, product2, qty, email1, email2, exp1, exp2, clear_extra_cells=False):
    today = datetime.date.today()
    today_str = today.strftime("%Y-%m-%d")
    safe_company = company.replace(' ', '')
    safe_product = product1.replace(' ', '')
    filename_base = f"{safe_company}_{safe_product}_{today.strftime('%Y%m%d')}"
    new_xlsx_path = os.path.join(output_dir, filename_base + '.xlsx')
    shutil.copy(template_path, new_xlsx_path)
    try:
        wb = openpyxl.load_workbook(new_xlsx_path)
        ws = wb.active
    except Exception as e:
        raise RuntimeError(f"엑셀 파일을 불러올 수 없습니다: {e}")
    # 치환 및 비어있는 경우 삭제
    if ws is not None:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = cell.value.replace('{Today}', today_str)
                    cell.value = cell.value.replace('{Company Name}', company)
                    cell.value = cell.value.replace('{Product_01}', product1)
                    cell.value = cell.value.replace('{Product_02}', product2)
                    cell.value = cell.value.replace('{Qty}', qty)
                    cell.value = cell.value.replace('{Email_01}', email1)
                    cell.value = cell.value.replace('{Email_02}', email2)
                    cell.value = cell.value.replace('{Exp date_01}', exp1)
                    cell.value = cell.value.replace('{Exp date_02}', exp2)
                    # 비어있는 Product_02, Email_02, Exp date_02 셀 삭제
                    if '{Product_02}' in cell.value and not product2:
                        cell.value = ''
                    if '{Email_02}' in cell.value and not email2:
                        cell.value = ''
                    if '{Exp date_02}' in cell.value and not exp2:
                        cell.value = ''
        # Maxon/ZBrush에서 두번째 이메일, 만료일 모두 비어있으면 A18, B19 삭제
        if clear_extra_cells and not email2 and not exp2:
            ws['A18'] = ''
            ws['B19'] = ''
        wb.save(new_xlsx_path)
    else:
        raise RuntimeError("엑셀 워크시트를 찾을 수 없습니다.")
    # PDF 변환
    excel = wc.DispatchEx('Excel.Application')
    excel.Visible = False
    wb_com = excel.Workbooks.Open(os.path.abspath(new_xlsx_path))
    pdf_path = os.path.join(output_dir, filename_base + '.pdf')
    wb_com.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
    wb_com.Close(False)
    excel.Quit()
    return new_xlsx_path, pdf_path

def create_pptx_certificate(template_path, output_dir, company, product1, product2, qty, email1, email2, exp1, exp2):
    if Presentation is None:
        raise ImportError('python-pptx 라이브러리가 필요합니다. (pip install python-pptx)')
    today = datetime.date.today()
    today_str = today.strftime("%Y-%m-%d")
    safe_company = company.replace(' ', '')
    safe_product = product1.replace(' ', '')
    filename_base = f"{safe_company}_{safe_product}_{today.strftime('%Y%m%d')}"
    new_pptx_path = os.path.join(output_dir, filename_base + '.pptx')
    shutil.copy(template_path, new_pptx_path)
    prs = Presentation(new_pptx_path)
    # 텍스트 치환 (텍스트프레임 + 표 셀 모두)
    def replace_all_text(prs, old, new):
        for slide in prs.slides:
            # 일반 텍스트 프레임 치환
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, new)
            # 표 셀 치환
            for shape in slide.shapes:
                if hasattr(shape, 'has_table') and shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            if old in cell.text:
                                cell.text = cell.text.replace(old, new)
    # 1행 치환
    replace_all_text(prs, '{Today}', today_str)
    replace_all_text(prs, '{Company Name}', company)
    replace_all_text(prs, '{Product_01}', product1)
    replace_all_text(prs, '{Qty_01}', qty)
    replace_all_text(prs, '{Email_01}', email1)
    replace_all_text(prs, '{Exp date_01}', exp1)
    # 2행 치환 (비어있으면 공백)
    replace_all_text(prs, '{Product_02}', product2 if product2 else '')
    replace_all_text(prs, '{Qty_02}', qty if product2 else '')
    replace_all_text(prs, '{Email_02}', email2 if email2 else '')
    replace_all_text(prs, '{Exp date_02}', exp2 if exp2 else '')
    prs.save(new_pptx_path)
    # PDF 변환 (PowerPoint COM)
    ppt = wc.DispatchEx('PowerPoint.Application')
    ppt.Visible = True  # 창을 숨기지 않음
    pres = ppt.Presentations.Open(os.path.abspath(new_pptx_path))  # WithWindow 인자 제거
    pdf_path = os.path.join(output_dir, filename_base + '.pdf')
    pres.SaveAs(os.path.abspath(pdf_path), 32)  # 32=PDF
    pres.Close()
    ppt.Quit()
    return new_pptx_path, pdf_path

def open_maxon_dialog(parent_root):
    dialog = tk.Toplevel(parent_root)
    dialog.title("Maxon/Red Giant/ZBrush 주문확인서 만들기")
    dialog.geometry("450x700")
    dialog.configure(bg="#F7F9FB")
    dialog.resizable(False, False)

    # 카드 프레임
    card = tk.Frame(dialog, bg="white", relief="solid", bd=1)
    card.place(relx=0.5, rely=0.5, anchor="center", width=400, height=580)
    card.grid_propagate(False)

    # 제품 선택
    tk.Label(card, text="제품 선택", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=0, column=0, padx=10, pady=8, sticky='e')
    product_var = tk.StringVar(value="Maxon")
    radio_frame = tk.Frame(card, bg="white")
    radio_frame.grid(row=0, column=1, padx=10, pady=8, sticky='w')
    for txt in ["Maxon", "ZBrush", "Red Giant"]:
        tk.Radiobutton(radio_frame, text=txt, variable=product_var, value=txt, font=("맑은 고딕", 10), bg="white").pack(side='left', padx=5)

    # 회사명
    tk.Label(card, text="회사명", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=1, column=0, padx=10, pady=8, sticky='e')
    entry_company = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=22)
    entry_company.grid(row=1, column=1, padx=10, pady=8, sticky='w')

    # 사용자 1
    user1_frame = tk.LabelFrame(card, text="사용자 1", font=("맑은 고딕", 10), bg="white", fg="#374151", relief="flat")
    user1_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 2), sticky='ew')
    tk.Label(user1_frame, text="제품명", font=("맑은 고딕", 10), bg="white").grid(row=0, column=0, padx=5, pady=2, sticky='e')
    entry_product1 = tk.Entry(user1_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_product1.grid(row=0, column=1, padx=5, pady=2, sticky='w')
    tk.Label(user1_frame, text="이메일", font=("맑은 고딕", 10), bg="white").grid(row=1, column=0, padx=5, pady=2, sticky='e')
    entry_email1 = tk.Entry(user1_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_email1.grid(row=1, column=1, padx=5, pady=2, sticky='w')
    tk.Label(user1_frame, text="만료일", font=("맑은 고딕", 10), bg="white").grid(row=2, column=0, padx=5, pady=2, sticky='e')
    entry_exp1 = tk.Entry(user1_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_exp1.grid(row=2, column=1, padx=5, pady=2, sticky='w')

    # 사용자 2 (선택)
    user2_frame = tk.LabelFrame(card, text="사용자 2 (선택)", font=("맑은 고딕", 10), bg="white", fg="#374151", relief="flat")
    user2_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=(2, 10), sticky='ew')
    tk.Label(user2_frame, text="제품명", font=("맑은 고딕", 10), bg="white").grid(row=0, column=0, padx=5, pady=2, sticky='e')
    entry_product2 = tk.Entry(user2_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_product2.grid(row=0, column=1, padx=5, pady=2, sticky='w')
    tk.Label(user2_frame, text="이메일", font=("맑은 고딕", 10), bg="white").grid(row=1, column=0, padx=5, pady=2, sticky='e')
    entry_email2 = tk.Entry(user2_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_email2.grid(row=1, column=1, padx=5, pady=2, sticky='w')
    tk.Label(user2_frame, text="만료일", font=("맑은 고딕", 10), bg="white").grid(row=2, column=0, padx=5, pady=2, sticky='e')
    entry_exp2 = tk.Entry(user2_frame, font=("맑은 고딕", 10), bg="#F1F5F9", relief="flat", width=22)
    entry_exp2.grid(row=2, column=1, padx=5, pady=2, sticky='w')

    # 수량
    tk.Label(card, text="수량", font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=10).grid(row=4, column=0, padx=10, pady=8, sticky='e')
    entry_qty = tk.Entry(card, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=22)
    entry_qty.grid(row=4, column=1, padx=10, pady=8, sticky='w')

    # 버튼 프레임
    btn_frame = tk.Frame(card, bg="white")
    btn_frame.grid(row=5, column=0, columnspan=2, pady=18)
    def on_create():
        company = entry_company.get().strip()
        product1 = entry_product1.get().strip()
        product2 = entry_product2.get().strip()
        qty = entry_qty.get().strip()
        email1 = entry_email1.get().strip()
        email2 = entry_email2.get().strip()
        exp1 = entry_exp1.get().strip()
        exp2 = entry_exp2.get().strip()
        if not company or not product1 or not qty or not email1 or not exp1:
            messagebox.showwarning("입력 오류", "필수 항목(회사명, 사용자1의 제품명, 수량, 이메일, 만료일)을 입력해주세요.")
            return
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(resource_root(), 'license_certificate', 'Maxon')
            os.makedirs(output_dir, exist_ok=True)
            selected = product_var.get()
            if selected == 'Maxon':
                template = resource_path('license_certificate/Maxon/Maxon_origin.xlsx')
                xlsx_path, pdf_path = create_excel_certificate(template, output_dir, company, product1, product2, qty, email1, email2, exp1, exp2, clear_extra_cells=True)
            elif selected == 'ZBrush':
                template = resource_path('license_certificate/Maxon/ZBrush_origin.xlsx')
                xlsx_path, pdf_path = create_excel_certificate(template, output_dir, company, product1, product2, qty, email1, email2, exp1, exp2, clear_extra_cells=True)
            elif selected == 'Red Giant':
                if Presentation is None:
                    messagebox.showerror("필수 라이브러리 없음", "Red Giant는 python-pptx가 필요합니다.\n명령프롬프트에서 pip install python-pptx 실행 후 이용하세요.")
                    return
                template = resource_path('license_certificate/Maxon/RedGiant_origin.pptx')
                pptx_path, pdf_path = create_pptx_certificate(template, output_dir, company, product1, product2, qty, email1, email2, exp1, exp2)
            else:
                messagebox.showerror("오류", "알 수 없는 제품 선택")
                return
            messagebox.showinfo("완료", f"PDF 파일: {pdf_path}\n\n[확인]을 누르면 폴더가 열립니다.")
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
1. 회사명과 사용자별 제품명, 이메일, 만료일을 입력합니다.
2. '만들기' 버튼을 클릭하여 원본 파일과 PDF 파일을 생성합니다.

📁 생성된 파일은 'license_certificate/'안에 각 제품별 폴더에 저장됩니다.""", 
                         justify='left', anchor='nw', font=("맑은 고딕", 9), bg="white", wraplength=400)
    info_label.grid(row=6, column=0, columnspan=2, pady=(10, 15), padx=10, sticky='w')

    dialog.transient(parent_root)
    dialog.grab_set()
    parent_root.wait_window(dialog) 