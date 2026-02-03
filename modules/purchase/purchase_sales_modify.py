# purchase_sales_modify.py
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
from modules.renewal import renewal_search as rs
import pandas as pd
from datetime import datetime

class PurchaseSalesModifyDialog:
    """매입/결제 단가 수정 다이얼로그"""
    def __init__(self, parent, row_data, refresh_callback=None):
        self.parent = parent
        self.row_data = row_data  # dict 형태
        self.refresh_callback = refresh_callback
        self.win = tk.Toplevel(self.parent)
        self.win.title("계산서/결제/매입 수정")
        self.win.geometry("600x800")  # 이미지에 맞게 크기 조정
        self.win.configure(bg="#F7F9FB")
        self.win.grab_set()
        self.win.resizable(True, True)  # 크기 조정 가능하도록 변경

        # 수정 가능한 컬럼: 계산서 발행일, 입금 예정일, 입금일, 입금액, 매입 계산서, 실 매입 단가
        self.target_cols = [
            ("계산서 발행일", "N"),
            ("입금 예정일", "O"),
            ("입금일", "P"),
            ("입금액", "Q"),
            ("매입 계산서", "T"),
            ("실 매입 단가", "U")
        ]
        

        
        self.create_widgets()

    def create_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.win, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="💰 계산서/결제/매입 수정", 
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
        basic_info = f"고객사: {self.row_data.get('고객사명', 'N/A')} | 제품: {self.row_data.get('제품', 'N/A')} | 만료일: {self.row_data.get('만료일', 'N/A')}"
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
            
            # 입력 필드 - row_data에서 값을 가져오되, 없으면 빈 문자열로 설정
            current_value = self.row_data.get(col, '')
            if current_value is None or pd.isna(current_value):
                current_value = ''
            else:
                current_value = str(current_value)
            
            # 계산서 발행일은 DateEntry 사용
            if col == "계산서 발행일":
                try:
                    # 날짜 형식 변환 시도
                    if current_value and current_value.strip():
                        date_obj = datetime.strptime(current_value, '%Y-%m-%d')
                        date_entry = DateEntry(label_frame, width=20, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                            font=("맑은 고딕", 10))
                        date_entry.set_date(date_obj)
                    else:
                        date_entry = DateEntry(label_frame, width=20, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                            font=("맑은 고딕", 10))
                except:
                    # 날짜 변환 실패 시 현재 날짜로 설정
                    date_entry = DateEntry(label_frame, width=20, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                        font=("맑은 고딕", 10))
                
                date_entry.pack(side='left', fill='x', expand=True, padx=(10, 0), pady=8)
                self.vars[col] = date_entry
            else:
                # 일반 텍스트 필드
                var = tk.StringVar(value=current_value)
                entry = tk.Entry(label_frame, textvariable=var, font=("맑은 고딕", 10),
                               bg="#F1F5F9", relief="flat", bd=1)
                entry.pack(side='left', fill='x', expand=True, padx=(10, 0), pady=8)
                self.vars[col] = var
                
                # 입력 필드 스타일링
                entry.bind('<FocusIn>', lambda e, ent=entry: ent.configure(bg="white", relief="solid", bd=1))
                entry.bind('<FocusOut>', lambda e, ent=entry: ent.configure(bg="#F1F5F9", relief="flat", bd=1))

        # 버튼 프레임 (가운데 정렬)
        button_frame = tk.Frame(self.win, bg="#F7F9FB")
        button_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # 버튼들을 담을 중앙 프레임
        center_frame = tk.Frame(button_frame, bg="#F7F9FB")
        center_frame.pack(expand=True)
        
        # 저장 버튼
        save_btn = tk.Button(center_frame, text="💾 저장", command=self.save,
                            font=("맑은 고딕", 11, "bold"), bg="#10B981", fg="white",
                            relief="flat", padx=20, pady=8, cursor="hand2")
        save_btn.pack(side='left', padx=(0, 10))
        save_btn.bind('<Enter>', lambda e: save_btn.configure(bg="#059669"))
        save_btn.bind('<Leave>', lambda e: save_btn.configure(bg="#10B981"))
        
        # 취소 버튼
        cancel_btn = tk.Button(center_frame, text="❌ 취소", command=self.win.destroy,
                              font=("맑은 고딕", 11), bg="#6B7280", fg="white",
                              relief="flat", padx=20, pady=8, cursor="hand2")
        cancel_btn.pack(side='left')
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.configure(bg="#374151"))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.configure(bg="#6B7280"))

    def save(self):
        try:
            # 구글시트 접근
            creds = rs._authorize()
            sh = creds.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
            ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
            headers = ws.row_values(1)
            records = ws.get_all_records()

            # 고유 식별자(예: 고객사명+제품+만료일)로 행 찾기
            key_cols = ['고객사명', '제품', '만료일']
            key = tuple(str(self.row_data.get(k, '')).strip() for k in key_cols)
            target_idx = None
            
            for i, rec in enumerate(records):
                rec_key = tuple(str(rec.get(k, '')).strip() for k in key_cols)
                if rec_key == key:
                    target_idx = i + 2  # 헤더 제외
                    break
                    
            if not target_idx:
                messagebox.showerror("오류", f"해당 행을 찾을 수 없습니다.\n찾고 있는 키: {key}")
                return

            # 각 입력값이 공란이 아니면 업데이트
            updated_count = 0
            for col, _ in self.target_cols:
                # DateEntry인지 일반 Entry인지 확인
                if isinstance(self.vars[col], DateEntry):
                    val = self.vars[col].get_date().strftime('%Y-%m-%d')
                else:
                    val = self.vars[col].get().strip()
                
                if val:
                    if col in headers:
                        col_idx = headers.index(col) + 1
                        ws.update_cell(target_idx, col_idx, val)
                        updated_count += 1

            if updated_count > 0:
                messagebox.showinfo("성공", f"{updated_count}개 항목이 수정되었습니다.")
            else:
                messagebox.showwarning("알림", "수정할 내용이 없습니다.")
                
            self.win.destroy()
            if self.refresh_callback:
                self.refresh_callback()
                
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}") 