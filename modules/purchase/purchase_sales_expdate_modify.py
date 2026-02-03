# purchase_sales_expdate_modify.py
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
from modules.renewal import renewal_search as rs
import pandas as pd
from datetime import datetime

class PurchaseSalesExpdateModifyDialog:
    """매입/결제 만료일 수정 다이얼로그"""
    def __init__(self, parent, row_data, refresh_callback=None):
        self.parent = parent
        self.row_data = row_data  # dict 형태
        self.refresh_callback = refresh_callback
        self.win = tk.Toplevel(self.parent)
        self.win.title("만료일 수정")
        self.win.geometry("600x400")  # 만료일만 수정하므로 크기 축소
        self.win.configure(bg="#F7F9FB")
        self.win.grab_set()
        self.win.resizable(True, True)

        # 수정 가능한 컬럼: 만료일만
        self.target_cols = [
            ("만료일", "N")  # 만료일 컬럼
        ]
        
        self.create_widgets()

    def create_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.win, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="📅 만료일 수정", 
                              font=("맑은 고딕", 16, "bold"), 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)

        # 메인 컨테이너
        main_container = tk.Frame(self.win, bg="#F7F9FB")
        main_container.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 메인 카드
        main_card = tk.Frame(main_container, bg="white", relief="solid", bd=1)
        main_card.pack(fill='both', expand=True, pady=(0, 20))
        
        # 카드 헤더
        card_header = tk.Frame(main_card, bg="#F8FAFC", height=40)
        card_header.pack(fill='x')
        card_header.pack_propagate(False)
        tk.Label(card_header, text="📝 수정할 정보", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드 내용
        card_content = tk.Frame(main_card, bg="white")
        card_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 현재 선택된 행 정보 표시
        info_frame = tk.Frame(card_content, bg="#F8FAFC", relief="solid", bd=1)
        info_frame.pack(fill='x', pady=(0, 20))
        tk.Label(info_frame, text="📋 선택된 행 정보", font=("맑은 고딕", 10, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(anchor='w', padx=10, pady=5)
        
        # 선택된 행의 기본 정보 표시
        basic_info = f"고객사: {self.row_data.get('고객사명', 'N/A')} | 제품: {self.row_data.get('제품', 'N/A')} | 현재 만료일: {self.row_data.get('만료일', 'N/A')}"
        tk.Label(info_frame, text=basic_info, font=("맑은 고딕", 9), 
                fg="#6B7280", bg="#F8FAFC", wraplength=600).pack(anchor='w', padx=10, pady=(0, 10))
        
        # 입력 필드들을 담을 프레임
        fields_frame = tk.Frame(card_content, bg="white")
        fields_frame.pack(fill='both', expand=True)
        
        # 입력 필드들 생성
        self.vars = {}
        
        for i, (col, _) in enumerate(self.target_cols):
            # 라벨 프레임
            label_frame = tk.Frame(fields_frame, bg="white")
            label_frame.pack(fill='x', pady=8)
            
            # 라벨
            tk.Label(label_frame, text=f"{col}:", font=("맑은 고딕", 10, "bold"), 
                    fg="#374151", bg="white", width=12, anchor='w').pack(side='left')
            
            # 현재 값 가져오기
            current_value = self.row_data.get(col, '')
            if current_value is None or pd.isna(current_value):
                current_value = ''
            else:
                current_value = str(current_value)
            
            # 날짜 선택 위젯
            date_frame = tk.Frame(label_frame, bg="white")
            date_frame.pack(side='left', fill='x', expand=True)
            
            # DateEntry 위젯 생성
            try:
                if current_value and current_value.strip():
                    # 기존 날짜가 있으면 파싱
                    try:
                        current_date = datetime.strptime(current_value, '%Y-%m-%d')
                    except:
                        current_date = datetime.today()
                else:
                    current_date = datetime.today()
            except:
                current_date = datetime.today()
            
            date_entry = DateEntry(
                date_frame,
                width=15,
                date_pattern='yyyy-mm-dd',
                font=("맑은 고딕", 10)
            )
            date_entry.set_date(current_date)
            date_entry.pack(side='left', padx=(0, 10))
            
            self.vars[col] = date_entry
        
        # 버튼 프레임
        button_frame = tk.Frame(card_content, bg="white")
        button_frame.pack(fill='x', pady=(20, 0))
        
        # 저장 버튼
        save_btn = tk.Button(button_frame, text="💾 저장", command=self.save,
                            font=("맑은 고딕", 11, "bold"), bg="#10B981", fg="white",
                            relief="flat", padx=20, pady=8)
        save_btn.pack(side='right', padx=(10, 0))
        save_btn.bind('<Enter>', lambda e: save_btn.configure(bg="#059669"))
        save_btn.bind('<Leave>', lambda e: save_btn.configure(bg="#10B981"))
        
        # 취소 버튼
        cancel_btn = tk.Button(button_frame, text="❌ 취소", command=self.win.destroy,
                              font=("맑은 고딕", 11), bg="#6B7280", fg="white",
                              relief="flat", padx=20, pady=8)
        cancel_btn.pack(side='right')
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.configure(bg="#4B5563"))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.configure(bg="#6B7280"))

    def save(self):
        """수정된 데이터를 Google Sheets에 저장"""
        try:
            # 입력된 값들 수집
            updated_data = {}
            for col, _ in self.target_cols:
                date_entry = self.vars[col]
                updated_data[col] = date_entry.get_date().strftime('%Y-%m-%d')
            
            # Google Sheets에서 해당 행 찾기 및 업데이트
            creds = rs._authorize()
            sh = creds.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
            ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
            
            # 모든 데이터 가져오기
            all_data = ws.get_all_values()
            headers = all_data[0]
            
            # 현재 행 찾기 (고객사명, 제품, 만료일로 매칭)
            target_row = None
            for i, row in enumerate(all_data[1:], start=2):  # 2부터 시작 (헤더 제외)
                if (len(row) > headers.index('고객사명') and 
                    len(row) > headers.index('제품') and 
                    len(row) > headers.index('만료일')):
                    
                    # 기존 만료일과 현재 만료일 비교
                    current_expiry = self.row_data.get('만료일', '')
                    if current_expiry:
                        # 시간 부분 제거
                        if ' ' in str(current_expiry):
                            current_expiry = str(current_expiry).split(' ')[0]
                    
                    if (row[headers.index('고객사명')] == self.row_data.get('고객사명', '') and
                        row[headers.index('제품')] == self.row_data.get('제품', '') and
                        row[headers.index('만료일')] == current_expiry):
                        target_row = i
                        break
            
            if target_row is None:
                messagebox.showerror("오류", "수정할 행을 찾을 수 없습니다.")
                return
            
            # 수정된 데이터 업데이트
            for col, _ in self.target_cols:
                if col in headers:
                    col_idx = headers.index(col)
                    new_value = updated_data[col]
                    ws.update_cell(target_row, col_idx + 1, new_value)  # 1-based indexing
            
            messagebox.showinfo("완료", "만료일이 성공적으로 수정되었습니다.")
            
            # 부모 창 새로고침
            if self.refresh_callback:
                self.refresh_callback()
            
            # 창 닫기
            self.win.destroy()
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"데이터 저장 중 오류가 발생했습니다:\n{str(e)}")
            print(f"만료일 수정 저장 오류: {e}") 