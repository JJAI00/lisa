import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

class QuoteProductAddWindow:
    """신규 제품(가격표) 등록 다이얼로그"""
    def __init__(self, parent, refresh_callback=None):
        self.parent = parent
        self.refresh_callback = refresh_callback
        self.win = tk.Toplevel(self.parent)
        self.win.title("신규 제품(가격표) 등록")
        self.win.geometry("420x650")
        self.win.configure(bg="#F7F9FB")
        self.win.resizable(False, False)

        # 카드 프레임
        card = tk.Frame(self.win, bg="white", relief="solid", bd=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=370, height=590)
        card.grid_propagate(False)

        # price_list에서 대분류 목록 로드
        self.categories = self.load_categories()

        # 입력 항목 정의
        fields = [
            ("대분류", "combo"),
            ("소분류", "entry"),
            ("상세 제품명", "entry"),
            ("구분", "combo"),
            ("USD", "entry"),
            ("EUR", "entry"),
            ("DC", "entry"),
            ("원가", "entry"),
            ("매입처", "entry"),
        ]
        self.vars = {}
        for i, (label, ftype) in enumerate(fields):
            tk.Label(card, text=label, font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=12).grid(row=i, column=0, padx=10, pady=7, sticky='e')
            if ftype == "entry":
                var = tk.StringVar()
                tk.Entry(card, textvariable=var, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=22).grid(row=i, column=1, padx=10, pady=7, sticky='w')
                self.vars[label] = var
            elif ftype == "combo":
                var = tk.StringVar()
                combo = tk.ttk.Combobox(card, textvariable=var, font=("맑은 고딕", 11), width=20, state='readonly')
                if label == "대분류":
                    combo['values'] = self.categories
                elif label == "구분":
                    combo['values'] = ["Renewal", "New", "Perpetual"]
                combo.grid(row=i, column=1, padx=10, pady=7, sticky='w')
                self.vars[label] = var

        # 저장 버튼
        save_btn = tk.Button(card, text="저장", font=("맑은 고딕", 11, "bold"),
                             bg="#3B82F6", fg="white", relief="flat", activebackground="#2563EB", activeforeground="white",
                             command=self.save)
        save_btn.grid(row=len(fields), column=0, columnspan=2, pady=18, ipadx=30, ipady=6)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#2563EB"))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg="#3B82F6"))

        # 모달 동작: 다른 창 선택 불가
        self.win.transient(self.parent)
        self.win.grab_set()
        self.win.wait_window(self.win)

    def load_categories(self):
        """price_list에서 대분류 목록 로드"""
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            from core.sendlog import KEY_FILE, SCOPES
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet1')
            
            # 모든 데이터 가져오기
            all_data = ws.get_all_records()
            df = pd.DataFrame(all_data)
            
            # 대분류 컬럼에서 고유값 추출
            if '대분류' in df.columns:
                categories = sorted(df['대분류'].dropna().unique().tolist())
                print(f"로드된 대분류 목록: {categories}")
                return categories
            else:
                print("대분류 컬럼을 찾을 수 없습니다.")
                return []
        except Exception as e:
            print(f"대분류 로드 오류: {e}")
            return []

    def save(self):
        data = {col: var.get().strip() for col, var in self.vars.items()}
        if not data.get('대분류', ''):
            messagebox.showwarning("입력 오류", "대분류를 반드시 입력해주세요.")
            return
        if not data.get('소분류', ''):
            messagebox.showwarning("입력 오류", "소분류를 반드시 입력해주세요.")
            return
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            from core.sendlog import KEY_FILE, SCOPES
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet1')
            # A열의 마지막 데이터 행 찾기
            a_col = ws.col_values(1)
            last_row = len([v for v in a_col if v.strip()]) + 1  # 다음 빈 행
            headers = ws.row_values(1)
            row = []
            for i, h in enumerate(headers):
                if i == 1:
                    # B열: 소분류 저장
                    row.append(data.get('소분류', ''))
                elif i == 2:
                    # C열: 제품명 수식 =D{행번호}&" - "&E{행번호}
                    row.append(f'=D{last_row}&" - "&E{last_row}')
                else:
                    row.append(data.get(h, ''))
            ws.append_row(row, value_input_option='USER_ENTERED')
            messagebox.showinfo("성공", "신규 제품이 price_list에 등록되었습니다.")
            self.win.destroy()
            if self.refresh_callback:
                self.refresh_callback()
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("저장 오류", f"price_list 저장 중 오류 발생:\n{e}") 