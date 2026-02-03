import tkinter as tk
from tkinter import messagebox
from modules.renewal import renewal_search as rs

class CustomerAddWindow:
    """신규 고객사 등록 다이얼로그"""
    def __init__(self, parent, refresh_callback):
        self.parent = parent
        self.refresh_callback = refresh_callback
        self.win = tk.Toplevel(self.parent)
        self.win.title("신규 고객사 등록")
        self.win.geometry("420x750")
        self.win.configure(bg="#F7F9FB")
        self.win.resizable(False, False)

        # 카드 프레임
        card = tk.Frame(self.win, bg="white", relief="solid", bd=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=370, height=690)
        card.grid_propagate(False)

        # 컬럼 헤더 조회
        creds = rs._authorize()
        sh = creds.open_by_key(rs.CUSTOMER_LIST_SHEET_ID)
        ws = sh.worksheet(rs.CUSTOMER_SHEET_NAME)
        headers = ws.row_values(1)

        # 입력 변수 생성 및 레이아웃
        self.vars = {}
        row_idx = 0
        
        for i, col in enumerate(headers):
            # 담당자명 필드는 숨기기
            if col == '담당자명':
                continue
                
            # 첫 번째 사업자번호 필드 제거 (두 번째 사업자번호만 유지)
            if col == '사업자번호' and i == 0:  # 첫 번째 사업자번호
                continue
                
            tk.Label(card, text=col, font=("맑은 고딕", 11), bg="white", fg="#374151", anchor='e', width=12).grid(row=row_idx, column=0, padx=10, pady=7, sticky='e')
            var = tk.StringVar()
            tk.Entry(card, textvariable=var, font=("맑은 고딕", 11), bg="#F1F5F9", relief="flat", width=22).grid(row=row_idx, column=1, padx=10, pady=7, sticky='w')
            self.vars[col] = var
            row_idx += 1

        # 저장 버튼
        save_btn = tk.Button(card, text="저장", font=("맑은 고딕", 11, "bold"),
                             bg="#3B82F6", fg="white", relief="flat", activebackground="#2563EB", activeforeground="white",
                             command=self.save)
        save_btn.grid(row=row_idx, column=0, columnspan=2, pady=18, ipadx=30, ipady=6)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#2563EB"))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg="#3B82F6"))

    def save(self):
        # 입력값 수집
        data = {col: var.get().strip() for col, var in self.vars.items()}
        name = data.get('고객사명', '')
        if not name:
            messagebox.showwarning("입력 오류", "고객사명을 입력해주세요.")
            return

        # 담당자명을 CONCATENATE 함수로 생성
        담당자 = data.get('담당자', '')
        직함 = data.get('직함', '')
        if 담당자 and 직함:
            # 행 번호를 한 번만 계산
            next_row = self._get_next_row()
            # 담당자와 직함 컬럼의 실제 위치 확인
            담당자_col = self._get_column_letter('담당자')
            직함_col = self._get_column_letter('직함')
            담당자명 = f'=CONCATENATE({담당자_col}{next_row}, " ",{직함_col}{next_row})'
            print(f"담당자명 수식 생성: {담당자명}")
            print(f"담당자 컬럼: {담당자_col}, 직함 컬럼: {직함_col}, 행 번호: {next_row}")
        else:
            담당자명 = f"{담당자} {직함}".strip()

        # 시트 접근
        creds = rs._authorize()
        sh = creds.open_by_key(rs.CUSTOMER_LIST_SHEET_ID)
        ws = sh.worksheet(rs.CUSTOMER_SHEET_NAME)
        headers = ws.row_values(1)
        records = ws.get_all_records()

        # 동일한 고객사명 레코드 필터링
        matching = [rec for rec in records if rec.get('고객사명', '') == name]

        if matching:
            # 동일한 고객사명이 존재할 경우
            # 동일한 담당자가 있는지 확인
            same_contact = [rec for rec in matching if rec.get('담당자', '') == 담당자]
            if same_contact:
                # 동일한 고객사명·담당자: 덮어쓰기
                if not messagebox.askyesno(
                    "중복 확인",
                    f"'{name}' 고객사와 담당자 '{담당자}' 정보가 이미 존재합니다.\n기존 정보를 덮어쓰시겠습니까?"
                ):
                    return
                idx = records.index(same_contact[0]) + 2  # 헤더 이후 행 번호
                for col, val in data.items():
                    if col in headers and val:
                        col_idx = headers.index(col) + 1
                        ws.update_cell(idx, col_idx, val)
                
                # 담당자명 업데이트
                if '담당자명' in headers:
                    담당자명_idx = headers.index('담당자명') + 1
                    ws.update_cell(idx, 담당자명_idx, 담당자명)
                    
                messagebox.showinfo("성공", "기존 고객사 정보가 업데이트되었습니다.")
            else:
                # 동일한 이름, 다른 담당자: 신규 등록 처리
                row = []
                for col in headers:
                    if col == '담당자명':
                        row.append(담당자명)
                    else:
                        row.append(data.get(col, ''))
                ws.append_row(row, value_input_option='USER_ENTERED')
                messagebox.showinfo("성공", "동일 고객사명에 다른 담당자로 신규 등록되었습니다.")
        else:
            # 존재하지 않을 경우: 신규 추가
            row = []
            for col in headers:
                if col == '담당자명':
                    row.append(담당자명)
                else:
                    row.append(data.get(col, ''))
            ws.append_row(row, value_input_option='USER_ENTERED')
            messagebox.showinfo("성공", "신규 고객사가 등록되었습니다.")

        # 창 닫기 및 메인 UI 갱신
        self.win.destroy()
        self.refresh_callback()

    def _get_next_row(self):
        """다음 행 번호를 반환하는 헬퍼 함수"""
        try:
            creds = rs._authorize()
            sh = creds.open_by_key(rs.CUSTOMER_LIST_SHEET_ID)
            ws = sh.worksheet(rs.CUSTOMER_SHEET_NAME)
            return len(ws.get_all_values()) + 1
        except:
            return 1000  # 기본값

    def _get_column_letter(self, column_name):
        """컬럼명에 해당하는 엑셀 컬럼 문자를 반환하는 헬퍼 함수"""
        try:
            creds = rs._authorize()
            sh = creds.open_by_key(rs.CUSTOMER_LIST_SHEET_ID)
            ws = sh.worksheet(rs.CUSTOMER_SHEET_NAME)
            headers = ws.row_values(1)
            
            print(f"컬럼 검색: '{column_name}'")
            print(f"전체 헤더: {headers}")
            
            if column_name in headers:
                col_index = headers.index(column_name)
                # 엑셀 컬럼 문자로 변환 (A=0, B=1, C=2, ...)
                col_letter = chr(ord('A') + col_index)
                print(f"컬럼 '{column_name}' 위치: {col_index} -> {col_letter}")
                return col_letter
            else:
                print(f"경고: 컬럼 '{column_name}'을 찾을 수 없습니다.")
                return 'A'  # 기본값
        except Exception as e:
            print(f"컬럼 문자 변환 오류: {e}")
            return 'A'  # 기본값
