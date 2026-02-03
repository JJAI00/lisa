# customer_info.py를 기반으로 견적서 탭용 quote_info.py로 복사 및 수정
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from modules.renewal import renewal_search as rs
from core.data_loader import get_global_auth, safe_api_call, clear_cache
import numpy as np
from datetime import datetime, timedelta
from integrations.lisa_daou_sign import show_context_menu
from .quote_status_modify import show_status_context_menu, QuoteStatusModifier
from integrations.lisa_logging import setup_logger

# 모듈 로거 설정
logger = setup_logger('quote_info')

class QuoteInfoGUI:
    """견적서 탭 내 임베디드 UI (신규 고객사 등록 기능 없음)"""
    def __init__(self, parent_frame, username=None, data_loader=None):
        self.parent_frame = parent_frame
        self.username = username or '테스트유저'
        self.data_loader = data_loader
        
        # Status 수정기 초기화
        self.status_modifier = QuoteStatusModifier()
        
        # 데이터 로더가 있으면 사용, 없으면 기존 방식 사용
        if self.data_loader:
            self.contacts_df = self.data_loader.get_customer_data()
            # 초기에는 전체 데이터를 로드하지 않고, load_initial_data에서 이번달 필터링된 데이터만 로드
            self.renewals_df = pd.DataFrame()  # 빈 DataFrame으로 초기화
            
            # 데이터가 비어있으면 기존 방식으로 로드
            if self.contacts_df.empty:
                print("견적서: 데이터 로더에서 고객사 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                self._load_data_directly()
            else:
                print(f"견적서: 데이터 로더에서 고객사 데이터를 성공적으로 가져왔습니다. 고객사: {len(self.contacts_df)}행")
        else:
            # 기존 방식 (하위 호환성)
            print("견적서: 기존 방식으로 데이터를 로드합니다.")
            self._load_data_directly()

        # 필터 옵션 리스트 (전체 데이터에서 추출)
        if self.data_loader:
            all_quote_data = self.data_loader.get_quote_data()
        else:
            all_quote_data = self.renewals_df  # 기존 방식에서는 전체 데이터 사용
            
        self.all_customers = sorted(all_quote_data['고객사명'].dropna().unique().tolist()) if '고객사명' in all_quote_data.columns else []
        self.all_clients   = sorted(all_quote_data['거래처명'].dropna().unique().tolist()) if '거래처명' in all_quote_data.columns else []
        self.all_categories = sorted(all_quote_data['대분류'].dropna().unique().tolist()) if '대분류' in all_quote_data.columns else []
        self.all_sales     = sorted(all_quote_data['영업사원'].dropna().unique().tolist()) if '영업사원' in all_quote_data.columns else []

        # 정렬 상태
        self.sort_column = None
        self.sort_reverse = False
        
        # 기준 선택 변수 (기본값: 일자)
        self.criteria_var = tk.StringVar(value="일자")
        
        # 현재 기간
        self.current_period = 'this_month'
        
        # 컬럼 순서 조정 (예상 발주일을 제일 앞에, Status를 일자 왼쪽으로 이동, 일자 오른쪽에 거래처명, 고객사명 순서 변경)
        base_cols = ['ID', 'FCST', 'Status', '일자', '거래처명', '고객사명', '영업사원', '대분류', '소분류', '제품', '수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액', '이익율', '만료일', '구분', 'USD', 'EUR']
        additional_cols = ['DC', '매입처']
        
        # DataFrame에 Status 컬럼 추가 (기본값 80%)
        if 'Status' not in self.renewals_df.columns:
            self.renewals_df['Status'] = '80%'
        
        # 기존 컬럼에서 base_cols에 없는 것들을 추가
        existing_cols = list(self.renewals_df.columns)
        extra_cols = [col for col in existing_cols if col not in base_cols and col not in additional_cols]
        
        # 구글 시트에서 컬럼 순서(헤더) 읽기
        if not self.renewals_df.empty:
            self.purchase_cols = list(self.renewals_df.columns)
        else:
            # 기존 방식 fallback
            self.purchase_cols = base_cols + additional_cols + extra_cols
        
        # 기타 컬럼 정의
        self.cost_cols = ['수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액']
        self.numeric_cols = ['수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액', '이익율']
        
        # UI 설정
        self.setup_widgets()
        self.load_initial_data()  # 먼저 데이터 로드 (이번달 필터링 포함)
        self.load_filters()       # 그 다음 필터 설정
        self.load_data()          # 마지막에 데이터 표시 (이번달 데이터만 표시됨)

    def _load_data_directly(self):
        """구글 시트에서 직접 데이터 로드"""
        try:
            # 전역 인증 사용
            gc = get_global_auth()
            
            # ── 고객사 정보(customer_list) 로드 ──
            def load_customer_data():
                sh = gc.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
                ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
                return ws.get_all_values()
            
            all_values = safe_api_call(load_customer_data)
            if all_values:
                raw = [h.strip() for h in all_values[0]]
                idxs = [i for i, h in enumerate(raw) if h]
                hdrs = pd.Index([raw[i] for i in idxs])
                rows = [[r[i] for i in idxs] for r in all_values[1:]]
                df_all = pd.DataFrame(rows, columns=hdrs)
                df_all.columns = df_all.columns.str.strip()
                if '고객사명' in df_all.columns:
                    info_cols = ['고객사명', '담당자', '연락처', '이메일', '주소']
                    if all(col in df_all.columns for col in info_cols):
                        self.contacts_df = df_all.drop_duplicates(subset='고객사명').loc[:, info_cols]
                    elif info_cols:
                        self.contacts_df = pd.DataFrame(columns=info_cols)
                    else:
                        self.contacts_df = pd.DataFrame()
                else:
                    self.contacts_df = pd.DataFrame()
            else:
                self.contacts_df = pd.DataFrame()
            self.contacts_df.columns = self.contacts_df.columns.str.strip()

            # ── 구매정보(quote_list) 로드 ──
            def load_quote_data():
                sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
                ws = sh.worksheet('Sheet1')
                return ws.get_all_records()
            
            quote_data = safe_api_call(load_quote_data)
            if quote_data:
                self.renewals_df = pd.DataFrame(quote_data)
                self.renewals_df.columns = self.renewals_df.columns.str.strip()
                
                # 일자 컬럼의 원본 데이터 확인
                if '일자' in self.renewals_df.columns:
                    print(f"견적서: 일자 컬럼 원본 데이터 샘플:")
                    print(self.renewals_df['일자'].head(5).tolist())
                    
                    # 한국어 날짜 형식을 파싱하는 함수
                    def parse_korean_date(date_str):
                        if pd.isna(date_str) or date_str == '':
                            return None
                        
                        try:
                            # 먼저 일반적인 형식으로 시도
                            return pd.to_datetime(date_str)
                        except:
                            try:
                                # 한국어 형식 파싱 (예: "2025. 8. 4 오전 9:25:07")
                                if isinstance(date_str, str):
                                    # "2025. 8. 4 오전 9:25:07" 형식 처리
                                    if '오전' in date_str or '오후' in date_str:
                                        # 오전/오후를 AM/PM으로 변환
                                        date_str = date_str.replace('오전', 'AM').replace('오후', 'PM')
                                        return pd.to_datetime(date_str, format='%Y. %m. %d %p %I:%M:%S')
                                    else:
                                        # "2025. 8. 4" 형식 처리
                                        return pd.to_datetime(date_str, format='%Y. %m. %d')
                            except:
                                try:
                                    # "2025-08-01 17:25:23" 형식 처리
                                    if '-' in date_str and ':' in date_str:
                                        return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S')
                                    # "2025-08-01" 형식 처리
                                    elif '-' in date_str:
                                        return pd.to_datetime(date_str, format='%Y-%m-%d')
                                except:
                                    print(f"견적서: 날짜 파싱 실패 - {date_str}")
                                    return None
                    
                    # 일자 컬럼을 파싱
                    self.renewals_df['일자_parsed'] = self.renewals_df['일자'].apply(parse_korean_date)
                    
                    # 파싱 성공한 데이터 확인
                    valid_dates = self.renewals_df[self.renewals_df['일자_parsed'].notna()]
                    failed_dates = self.renewals_df[self.renewals_df['일자_parsed'].isna()]
                    
                    print(f"견적서: 날짜 파싱 결과 - 성공: {len(valid_dates)}행, 실패: {len(failed_dates)}행")
                    
                    if not failed_dates.empty:
                        print(f"견적서: 파싱 실패한 날짜 샘플:")
                        print(failed_dates['일자'].head(3).tolist())
                    
                    # 파싱된 날짜를 문자열로 변환하여 일자 컬럼 업데이트
                    self.renewals_df['일자'] = self.renewals_df['일자_parsed'].dt.strftime('%Y-%m-%d %H:%M:%S')
                    self.renewals_df = self.renewals_df.drop('일자_parsed', axis=1)
                    
                    print(f"견적서: 날짜 변환 완료 - 샘플:")
                    print(self.renewals_df['일자'].head(3).tolist())
                
                if '만료일' in self.renewals_df.columns:
                    self.renewals_df['만료일'] = pd.to_datetime(self.renewals_df['만료일']).dt.date
            else:
                self.renewals_df = pd.DataFrame()
                
        except Exception as e:
            print(f"견적서 초기 데이터 로드 오류: {e}")
            self.contacts_df = pd.DataFrame()
            self.renewals_df = pd.DataFrame()

    def setup_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.parent_frame, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="견적서 관리", 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 사용자 정보 (우측)
        user_label = tk.Label(header_frame, text=f"👤 {self.username}", 
                             fg="#6B7280", bg="#F7F9FB")
        user_label.pack(side='right', padx=20, pady=15)

        # ── 필터 카드 영역 ──
        filter_section = tk.Frame(self.parent_frame, bg="#F7F9FB")
        filter_section.pack(fill='x', padx=20, pady=(0, 20))

        # 첫 번째 카드: 기준 선택 및 날짜
        card1 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card1.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 카드1 헤더
        card1_header = tk.Frame(card1, bg="#F8FAFC", height=40)
        card1_header.pack(fill='x')
        card1_header.pack_propagate(False)
        tk.Label(card1_header, text="📅 조회 기준", 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드1 내용
        card1_content = tk.Frame(card1, bg="white")
        card1_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 기준 선택
        criteria_frame = tk.Frame(card1_content, bg="white")
        criteria_frame.pack(fill='x', pady=(0, 10))
        tk.Label(criteria_frame, text="조회 기준:", 
                fg="#374151", bg="white").pack(side='left')
        tk.Radiobutton(criteria_frame, text="일자", variable=self.criteria_var, 
                      value="일자", command=self.on_criteria_changed,
                      fg="#374151", bg="white").pack(side='left', padx=5)
        tk.Radiobutton(criteria_frame, text="예상 발주일", variable=self.criteria_var, 
                      value="FCST", command=self.on_criteria_changed,
                      fg="#374151", bg="white").pack(side='left', padx=5)
        
        # 날짜 선택
        date_frame = tk.Frame(card1_content, bg="white")
        date_frame.pack(fill='x', pady=(0, 10))
        tk.Label(date_frame, text="조회 기준일:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        from tkcalendar import DateEntry
        from datetime import datetime
        self.date_entry = DateEntry(
            date_frame,
            width=12,
            date_pattern='yyyy-mm-dd'
        )
        self.date_entry.set_date(datetime.today())
        self.date_entry.pack(side='left', padx=5)
        # 날짜 선택 시 현재 기간 유지
        self.date_entry.bind(
            "<<DateEntrySelected>>",
            lambda e: self.load_period(self.current_period)
        )
        
        # 기간 버튼들 - 2줄로 분리
        # 첫 번째 줄: 이번달, 다음달, 미래 1년
        first_row_frame = tk.Frame(card1_content, bg="white")
        first_row_frame.pack(fill='x', pady=(0, 5))
        
        first_row_periods = [
            ("이번달",   lambda: self.load_period('this_month')),
            ("다음달",   lambda: self.load_period('next_month')),
            ("미래 1년", lambda: self.load_period('one_year')),
        ]
        for text, cmd in first_row_periods:
            btn = tk.Button(
                first_row_frame, text=text,
                command=cmd,
                font=("맑은 고딕", 9), bg="#3B82F6", fg="white",
                relief="flat", padx=10, pady=3
            )
            btn.pack(side='left', padx=2)
            # hover 효과
            btn.bind('<Enter>', lambda e, b=btn: b.configure(bg="#2563EB"))
            btn.bind('<Leave>', lambda e, b=btn: b.configure(bg="#3B82F6"))
        
        # 두 번째 줄: 과거 1년, 새로고침
        second_row_frame = tk.Frame(card1_content, bg="white")
        second_row_frame.pack(fill='x')
        
        # 과거 1년 버튼
        past_year_btn = tk.Button(second_row_frame, text="과거 1년", 
                                 command=lambda: self.load_period('past_year'),
                                 font=("맑은 고딕", 9), bg="#3B82F6", fg="white",
                                 relief="flat", padx=10, pady=3)
        past_year_btn.pack(side='left', padx=2)
        past_year_btn.bind('<Enter>', lambda e, b=past_year_btn: b.configure(bg="#2563EB"))
        past_year_btn.bind('<Leave>', lambda e, b=past_year_btn: b.configure(bg="#3B82F6"))
        
        # 새로고침 버튼
        refresh_btn = tk.Button(second_row_frame, text="🔄 새로고침",
                               command=self.refresh_data,
                               font=("맑은 고딕", 9), bg="#10B981", fg="white",
                               relief="flat", padx=10, pady=3)
        refresh_btn.pack(side='left', padx=2)
        refresh_btn.bind('<Enter>', lambda e, b=refresh_btn: b.configure(bg="#059669"))
        refresh_btn.bind('<Leave>', lambda e, b=refresh_btn: b.configure(bg="#10B981"))
        
        # 전체 데이터 로드 버튼 (A)
        all_data_btn = tk.Button(second_row_frame, text="A", command=self.load_all_data,
                                font=("맑은 고딕", 9, "bold"), bg="#8B5CF6", fg="white",
                                relief="flat", padx=10, pady=3)
        all_data_btn.pack(side='left', padx=2)
        # hover 효과
        all_data_btn.bind('<Enter>', lambda e, b=all_data_btn: b.configure(bg="#7C3AED"))
        all_data_btn.bind('<Leave>', lambda e, b=all_data_btn: b.configure(bg="#8B5CF6"))

        # 두 번째 카드: 필터 옵션
        card2 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card2.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 카드2 헤더
        card2_header = tk.Frame(card2, bg="#F8FAFC", height=40)
        card2_header.pack(fill='x')
        card2_header.pack_propagate(False)
        tk.Label(card2_header, text="🔍 필터", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드2 내용
        card2_content = tk.Frame(card2, bg="white")
        card2_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 첫 번째 행: 고객사명, 거래처명
        row1 = tk.Frame(card2_content, bg="white")
        row1.pack(fill='x', pady=(0, 10))
        
        tk.Label(row1, text="고객사명:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.customer_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_customer_list
        )
        self.customer_cb.pack(side='left', padx=5)
        self.customer_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.customer_cb.bind('<Return>', lambda e: self.on_customer_search())

        tk.Label(row1, text="거래처명:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.client_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_client_list
        )
        self.client_cb.pack(side='left', padx=5)
        self.client_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.client_cb.bind('<Return>', lambda e: self.on_client_search())

        # 두 번째 행: 대분류, 영업사원
        row2 = tk.Frame(card2_content, bg="white")
        row2.pack(fill='x', pady=(0, 10))
        
        tk.Label(row2, text="대분류:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.category_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_category_list
        )
        self.category_cb.pack(side='left', padx=5)
        self.category_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.category_cb.bind('<Return>', lambda e: self.on_search())

        tk.Label(row2, text="영업사원:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.sales_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_sales_list
        )
        self.sales_cb.pack(side='left', padx=5)
        self.sales_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.sales_cb.bind('<Return>', lambda e: self.on_search())

        # 세 번째 행: Status, 검색 버튼
        row3 = tk.Frame(card2_content, bg="white")
        row3.pack(fill='x')
        
        tk.Label(row3, text="Status:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.status_cb = ttk.Combobox(
            row3,
            state='normal',
            width=20,
            postcommand=self.update_status_list
        )
        self.status_cb.pack(side='left', padx=5)
        self.status_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.status_cb.bind('<Return>', lambda e: self.on_search())
        
        search_btn = tk.Button(row3, text="🔍 검색", command=self.on_search,
                              font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white",
                              relief="flat", padx=15, pady=4)
        search_btn.pack(side='left', padx=(20,0))
        search_btn.bind('<Enter>', lambda e: search_btn.configure(bg="#059669"))
        search_btn.bind('<Leave>', lambda e: search_btn.configure(bg="#10B981"))

        # 세 번째 카드: 액션 프레임 (필터 오른쪽에 배치)
        card3 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card3.pack(side='left', fill='both', expand=True)
        
        # 카드3 헤더
        card3_header = tk.Frame(card3, bg="#F8FAFC", height=40)
        card3_header.pack(fill='x')
        card3_header.pack_propagate(False)
        tk.Label(card3_header, text="⚡ 액션", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드3 내용
        card3_content = tk.Frame(card3, bg="white")
        card3_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 액션 버튼들 (세로 배치)
        action_buttons = tk.Frame(card3_content, bg="white")
        action_buttons.pack(fill='both', expand=True)
        
        # 신규 고객사 등록 버튼
        def open_customer_add():
            try:
                from modules.customer.customer_add import CustomerAddWindow
                CustomerAddWindow(self.parent_frame, self.reload_customers)
            except Exception as e:
                import traceback
                traceback.print_exc()
                from tkinter import messagebox
                messagebox.showerror('오류', f'신규 고객사 등록 창을 여는 데 실패했습니다:\n{e}')
        add_btn = tk.Button(action_buttons, text="➕ 신규 고객사 등록",
                           command=open_customer_add,
                           font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                           relief="flat", padx=15, pady=4)
        add_btn.pack(fill='x', pady=(0, 5))
        add_btn.bind('<Enter>', lambda e: add_btn.configure(bg="#2563EB"))
        add_btn.bind('<Leave>', lambda e: add_btn.configure(bg="#3B82F6"))

        # 견적 작성(신규) 버튼
        def open_quote_new():
            try:
                from modules.quote.quote_new import QuoteNewDialog
                QuoteNewDialog(self.parent_frame, username=self.username, callback=self.refresh_quote_data)
            except Exception as e:
                import traceback
                traceback.print_exc()
                from tkinter import messagebox
                messagebox.showerror('오류', f'견적 작성(신규) 창을 여는 데 실패했습니다:\n{e}')
        quote_new_btn = tk.Button(action_buttons, text="📝 견적 작성(신규)",
                                 command=open_quote_new,
                                 font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white",
                                 relief="flat", padx=15, pady=4)
        quote_new_btn.pack(fill='x', pady=(0, 5))
        quote_new_btn.bind('<Enter>', lambda e: quote_new_btn.configure(bg="#059669"))
        quote_new_btn.bind('<Leave>', lambda e: quote_new_btn.configure(bg="#10B981"))

        # 견적 수정 버튼
        def open_quote_modify():
            selected = self.purchase_tree.selection()
            if not selected:
                from tkinter import messagebox
                messagebox.showwarning("경고", "수정할 견적서를 선택해주세요.")
                return
            
            # 하나의 행만 선택하도록 변경
            if len(selected) > 1:
                from tkinter import messagebox
                messagebox.showwarning("경고", "하나의 행만 선택해주세요.")
                return
            
            # 선택된 행의 데이터 추출
            selected_item = selected[0]
            values = self.purchase_tree.item(selected_item)['values']
            selected_row_data = {col: values[i] if i < len(values) else '' for i, col in enumerate(self.purchase_cols)}
            
            # ID 값을 기준으로 같은 견적서의 모든 행 찾기
            target_id = selected_row_data.get('ID', '')
            
            if not target_id:
                from tkinter import messagebox
                messagebox.showwarning("경고", "ID 값이 있는 행을 선택해주세요.")
                return
            
            # 전체 데이터에서 같은 ID를 가진 모든 행들 찾기
            matching_rows = []
            for item in self.purchase_tree.get_children():
                values = self.purchase_tree.item(item)['values']
                row_data = {col: values[i] if i < len(values) else '' for i, col in enumerate(self.purchase_cols)}
                
                row_id = row_data.get('ID', '')
                
                # 같은 ID를 가진 행들 찾기
                if row_id == target_id:
                    matching_rows.append(row_data)
            
            if not matching_rows:
                from tkinter import messagebox
                messagebox.showwarning("경고", "일치하는 데이터를 찾을 수 없습니다.")
                return
            
            print(f"견적 수정: {len(matching_rows)}개 행을 찾았습니다.")
            print(f"조건: ID={target_id}")
            
            # 찾은 모든 행을 견적 수정 창에 전달
            rows_data = matching_rows
            
            try:
                from modules.quote.quote_modify import QuoteModifyDialog
                QuoteModifyDialog(self.parent_frame, rows_data=rows_data, username=self.username, callback=self.refresh_quote_data)
            except Exception as e:
                import traceback
                traceback.print_exc()
                from tkinter import messagebox
                messagebox.showerror('오류', f'견적 수정 창을 여는 데 실패했습니다:\n{e}')
        quote_modify_btn = tk.Button(action_buttons, text="✏️ 견적 작성(수정)",
                                    command=open_quote_modify,
                                    font=("맑은 고딕", 10, "bold"), bg="#F59E42", fg="white",
                                    relief="flat", padx=15, pady=4)
        quote_modify_btn.pack(fill='x', pady=(0, 5))
        quote_modify_btn.bind('<Enter>', lambda e: quote_modify_btn.configure(bg="#EA580C"))
        quote_modify_btn.bind('<Leave>', lambda e: quote_modify_btn.configure(bg="#F59E42"))

        # 고객사 정보 프레임(카드) 완전히 제거
        # info_section = ... (삭제)
        # info_card = ... (삭제)
        # info_header = ... (삭제)
        # info_content = ... (삭제)
        # self.info_labels = ... (삭제)

        # ── 구매 정보 테이블 섹션 ──
        purchase_section = tk.Frame(self.parent_frame, bg="#F7F9FB")
        purchase_section.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 테이블 헤더
        purchase_header = tk.Frame(purchase_section, bg="white", relief="solid", bd=1)
        purchase_header.pack(fill='x', pady=(0, 1))
        tk.Label(purchase_header, text="📊 견적서 데이터", font=("맑은 고딕", 14, "bold"), 
                fg="#1F2937", bg="white").pack(side='left', padx=15, pady=10)
        
        # 테이블 컨테이너
        purchase_container = tk.Frame(purchase_section, bg="white", relief="solid", bd=1)
        purchase_container.pack(fill='both', expand=True)
        
        # 컬럼 순서 조정 (예상 발주일을 제일 앞에, Status를 일자 왼쪽으로 이동, 일자 오른쪽에 거래처명, 고객사명 순서 변경)
        base_cols = ['ID', 'FCST', 'Status', '일자', '거래처명', '고객사명', '영업사원', '대분류', '소분류', '제품', '수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액', '이익율', '만료일', '구분', 'USD', 'EUR']
        additional_cols = ['DC', '매입처']
        
        # DataFrame에 Status 컬럼 추가 (기본값 80%)
        if 'Status' not in self.renewals_df.columns:
            self.renewals_df['Status'] = '80%'
        
        # 기존 컬럼에서 base_cols에 없는 것들을 추가
        existing_cols = list(self.renewals_df.columns)
        extra_cols = [col for col in existing_cols if col not in base_cols and col not in additional_cols]
        
        # 구글 시트에서 컬럼 순서(헤더) 읽기
        if not self.renewals_df.empty:
            self.purchase_cols = list(self.renewals_df.columns)
        else:
            # 기존 방식 fallback
            self.purchase_cols = base_cols + additional_cols + extra_cols
        
        # 숫자 컬럼 정의 (정렬용)
        self.numeric_cols = ['수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액', '이익율', 'USD', 'EUR', 'DC']
        
        self.purchase_tree = ttk.Treeview(purchase_container,
                                          columns=self.purchase_cols,
                                          show='headings',
                                          selectmode='browse')
        vsb = ttk.Scrollbar(purchase_container, orient='vertical',
                             command=self.purchase_tree.yview)
        hsb = ttk.Scrollbar(purchase_container, orient='horizontal',
                             command=self.purchase_tree.xview)
        self.purchase_tree.configure(yscrollcommand=vsb.set,
                                     xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.purchase_tree.pack(fill='both', expand=True, padx=1, pady=1)

        # Shift+휠로 가로 스크롤 빠르게 이동
        def fast_horizontal_scroll(event):
            try:
                widget = event.widget
                if widget and widget.winfo_exists():
                    if event.delta < 0:
                        widget.xview_scroll(3, "units")
                    else:
                        widget.xview_scroll(-3, "units")
            except Exception:
                # 위젯이 파괴되었거나 접근할 수 없는 경우 무시
                pass
        self.purchase_tree.bind("<Shift-MouseWheel>", fast_horizontal_scroll)
        
        # 컬럼 헤더 설정 및 정렬 기능 추가
        for col in self.purchase_cols:
            anchor = 'e' if col in self.numeric_cols else 'center'
            self.purchase_tree.heading(col, text=col,
                command=lambda c=col: self.sort_purchase(c))
            # 컬럼별 기본 너비 설정
            if col in ['FCST']:
                width = 120
            elif col in ['일자', '만료일']:
                width = 120
            elif col in ['고객사명', '거래처명', '제품']:
                width = 150
            elif col in ['대분류', '소분류', '영업사원']:
                width = 100
            elif col in ['수량', '판매 단가', '원가', '판매 합계', '원가계', '이익액', '이익율']:
                width = 120
            elif col in ['USD', 'EUR', 'DC']:
                width = 80
            else:
                width = 100
            self.purchase_tree.column(col, width=width,
                                     anchor=anchor, stretch=False)
        
        # 우클릭 이벤트 바인딩 (Status 메뉴 포함)
        self.purchase_tree.bind("<Button-3>", lambda event: show_status_context_menu(event, self.purchase_tree, self.status_modifier))
        
        # 컬럼 크기 조절 기능 추가 (ResizableTreeview 사용)
        from ui.resizable_treeview import ResizableTreeview
        self.resizable_tree = ResizableTreeview(self.purchase_tree, "quote_table")

    def on_criteria_changed(self):
        """조회 기준 변경 시 호출되는 함수"""
        print(f"견적서: 조회 기준 변경 - {self.criteria_var.get()}")
        # 현재 기간으로 다시 로드
        self.load_period(self.current_period)

    def load_filters(self):
        self.customer_cb['values'] = ['전체'] + self.all_customers
        self.customer_cb.set('전체')
        self.client_cb['values']   = ['전체'] + self.all_clients
        self.client_cb.set('전체')
        self.category_cb['values'] = ['전체'] + self.all_categories
        self.category_cb.set('전체')
        self.sales_cb['values']    = ['전체'] + self.all_sales
        self.sales_cb.set('전체')
        self.status_cb['values']   = ['전체', '100%', '80%', '60%', '40%', '20%', 'Drop']
        self.status_cb.set('전체')
        # 무한 재귀 방지를 위해 load_period 호출 제거

    def load_period(self, mode):
        """기간별 데이터 로드"""
        try:
            print(f"견적서: {mode} 데이터 로드 시작")
            # 현재 기간 업데이트
            self.current_period = mode
            
            if mode == 'this_month':
                # 이번달 데이터 로드 로직
                self.load_initial_data()
                # load_initial_data는 load_data를 호출하지 않으므로 여기서 호출
                self.load_data()
            elif mode == 'next_month':
                # 다음달 데이터 로드 로직
                self.load_next_month_data()
                # load_next_month_data는 load_data를 호출하므로 중복 호출 방지
            elif mode == 'one_year':
                # 미래 1년 데이터 로드 로직
                self.load_future_year_data()
                # load_future_year_data는 load_data를 호출하므로 중복 호출 방지
            elif mode == 'past_year':
                # 과거 1년 데이터 로드 로직
                self.load_past_year_data()
                # load_past_year_data는 load_data를 호출하므로 중복 호출 방지
            elif mode == 'all':
                # 전체 데이터 로드 로직
                self.load_all_data()
                # load_all_data는 load_data를 호출하므로 중복 호출 방지
            
            # 데이터 로드 후 트리뷰 우클릭 이벤트 바인딩 재설정
            try:
                # 기존 바인딩 제거
                self.purchase_tree.unbind("<Button-3>")
                # 새로운 바인딩 설정
                self.purchase_tree.bind("<Button-3>", lambda event: show_status_context_menu(event, self.purchase_tree, self.status_modifier))
                print("견적서: 트리뷰 우클릭 이벤트 바인딩 재설정 완료")
            except Exception as e:
                print(f"견적서: 트리뷰 이벤트 바인딩 재설정 실패: {e}")
                
        except Exception as e:
            print(f"견적서: {mode} 데이터 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"{mode} 데이터를 불러오는 중 오류가 발생했습니다:\n{e}")

    def refresh_data(self):
        """명시적으로 캐시를 클리어하고 최신 데이터를 가져오기"""
        try:
            print("견적서: 데이터 새로고침 시작")
            
            # 캐시 클리어
            clear_cache()
            
            # filtered_df 속성 정리 (새로고침 시 초기화)
            if hasattr(self, 'filtered_df'):
                delattr(self, 'filtered_df')
                print("견적서: filtered_df 속성 정리 완료")
            
            # 데이터 로더가 있으면 강제로 새로 로드
            if self.data_loader:
                print("견적서: 데이터 로더를 통해 최신 데이터 로드 중...")
                
                # 데이터 로더의 모든 캐시 클리어
                self.data_loader.clear_all_cache()
                
                # 전체 데이터를 다시 로드
                self.renewals_df = self.data_loader.get_quote_data()
                print(f"견적서: 전체 데이터 로드 완료 - {len(self.renewals_df)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if self.renewals_df.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    print(f"견적서: 직접 로드 완료 - {len(self.renewals_df)}행")
                
                if not self.renewals_df.empty:
                    print(f"견적서: 컬럼 목록: {list(self.renewals_df.columns)}")
                    print(f"견적서: 최신 데이터 샘플 (첫 3행):")
                    print(self.renewals_df.head(3).to_string())
                    
                    # 필터 옵션 리스트 업데이트 (전체 데이터에서 추출)
                    self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                    self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                    self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                    self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                    
                    print(f"견적서: 필터 옵션 업데이트 완료")
                    print(f"- 고객사명: {len(self.all_customers)}개")
                    print(f"- 거래처명: {len(self.all_clients)}개")
                    print(f"- 대분류: {len(self.all_categories)}개")
                    print(f"- 영업사원: {len(self.all_sales)}개")
                    
                    # 필터 콤보박스 업데이트
                    self.load_filters()
                    
                    # 현재 기간에 따라 데이터 새로고침 (중복 로직 제거)
                    print(f"견적서: 현재 기간 '{self.current_period}'에 따라 데이터 새로고침")
                    self.load_period(self.current_period)
                else:
                    print("견적서: 새로고침 후에도 데이터가 없습니다.")
                    # 빈 데이터라도 필터 옵션은 초기화
                    self.all_customers = []
                    self.all_clients = []
                    self.all_categories = []
                    self.all_sales = []
                    self.load_filters()
                    # 빈 테이블 표시
                    self.load_data()
            else:
                # 데이터 로더가 없으면 기존 방식
                self._load_data_directly()
                self.load_filters()
                
                # 현재 기간에 따라 데이터 새로고침
                print(f"견적서: 기존 방식으로 현재 기간 '{self.current_period}'에 따라 데이터 새로고침")
                self.load_period(self.current_period)
            
            # 새로고침 후 트리뷰 우클릭 이벤트 바인딩 재설정
            try:
                # 기존 바인딩 제거
                self.purchase_tree.unbind("<Button-3>")
                # 새로운 바인딩 설정
                self.purchase_tree.bind("<Button-3>", lambda event: show_status_context_menu(event, self.purchase_tree, self.status_modifier))
                print("견적서: 트리뷰 우클릭 이벤트 바인딩 재설정 완료")
            except Exception as e:
                print(f"견적서: 트리뷰 이벤트 바인딩 재설정 실패: {e}")
                
            print(f"견적서: 데이터 새로고침 완료 - 총 {len(self.renewals_df)}행")
            
        except Exception as e:
            print(f"견적서: 데이터 새로고침 오류: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"데이터 새로고침 중 오류가 발생했습니다:\n{e}")

    def load_next_month_data(self):
        """다음달 데이터 로드"""
        try:
            # 선택된 날짜 기준으로 다음달 계산
            selected_date = self.date_entry.get_date()
            print(f"견적서: 선택된 조회 기준일 - {selected_date.strftime('%Y-%m-%d')}")
            
            if selected_date.month == 12:
                next_month = selected_date.replace(year=selected_date.year + 1, month=1, day=1)
            else:
                next_month = selected_date.replace(month=selected_date.month + 1, day=1)
            
            print(f"견적서: 다음달 데이터 로드 시작 - {next_month.strftime('%Y-%m')}")
            
            # 데이터 로더에서 전체 데이터 가져온 후 필터링
            if self.data_loader:
                self.renewals_df = self.data_loader.get_quote_data()
                
                print(f"견적서: 전체 데이터 로드 완료 - {len(self.renewals_df)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if self.renewals_df.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    print(f"견적서: 직접 로드 완료 - {len(self.renewals_df)}행")
                
                # 다음달 필터링 - 기준에 따라 다른 컬럼 사용
                criteria = self.criteria_var.get()
                if criteria == "일자" and not self.renewals_df.empty and '일자' in self.renewals_df.columns:
                    print(f"견적서: 일자 컬럼 확인 - 샘플 데이터: {self.renewals_df['일자'].head(3).tolist()}")
                    
                    # 일자 컬럼을 datetime으로 변환 (이미 파싱된 형식 사용)
                    self.renewals_df['일자_dt'] = pd.to_datetime(self.renewals_df['일자'], errors='coerce')
                    
                    # 변환 실패한 데이터 확인
                    failed_conversions = self.renewals_df[self.renewals_df['일자_dt'].isna()]
                    if not failed_conversions.empty:
                        print(f"견적서: 날짜 변환 실패한 데이터 {len(failed_conversions)}행")
                        print(f"견적서: 변환 실패 샘플: {failed_conversions['일자'].head(3).tolist()}")
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['일자_dt'].notna()].copy()
                    print(f"견적서: 유효한 날짜 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 다음달의 마지막 날 계산
                        if next_month.month == 12:
                            end_of_next_month = next_month.replace(year=next_month.year + 1, month=1, day=1) - timedelta(days=1)
                        else:
                            end_of_next_month = next_month.replace(month=next_month.month + 1, day=1) - timedelta(days=1)
                        
                        print(f"견적서: 다음달 필터링 기간 - {next_month.strftime('%Y-%m-%d')} ~ {end_of_next_month.strftime('%Y-%m-%d')}")
                        
                        # 다음달 데이터 필터링 (시간 무시)
                        # 날짜 비교를 위해 모든 날짜를 datetime.date 객체로 변환
                        valid_data['일자_날짜만'] = valid_data['일자_dt'].dt.date
                        start_date = next_month.date() if hasattr(next_month, 'date') else next_month
                        end_date = end_of_next_month.date() if hasattr(end_of_next_month, 'date') else end_of_next_month
                        
                        mask = (valid_data['일자_날짜만'] >= start_date) & (valid_data['일자_날짜만'] <= end_date)
                        filtered_data = valid_data[mask].copy()
                        
                        print(f"견적서: 필터링 마스크 적용 전 데이터 수 - {len(valid_data)}행")
                        print(f"견적서: 필터링 조건 - {start_date} <= 일자 <= {end_date}")
                        
                        # 필터링 전후 데이터 샘플 출력
                        print(f"견적서: 필터링 전 데이터 샘플 (첫 5행):")
                        print(valid_data['일자_날짜만'].head(5).tolist())
                        
                        print(f"견적서: 필터링 후 데이터 샘플 (첫 5행):")
                        print(filtered_data['일자_날짜만'].head(5).tolist())
                        
                        # 일자 컬럼을 날짜만 표시하도록 수정
                        filtered_data['일자'] = filtered_data['일자_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop(['일자_dt', '일자_날짜만'], axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 다음달 필터링 완료 - {len(self.renewals_df)}행")
                        
                        # 메모리 사용량 모니터링
                        try:
                            import psutil
                            process = psutil.Process()
                            memory_mb = process.memory_info().rss / 1024 / 1024
                            print(f"견적서: 메모리 사용량 - {memory_mb:.1f} MB")
                        except ImportError:
                            pass
                        
                        # 디버깅용: 필터링된 데이터 샘플 출력
                        if len(self.renewals_df) > 0:
                            sample_dates = self.renewals_df['일자'].head(5).tolist()
                            print(f"견적서: 필터링된 일자 샘플: {sample_dates}")
                        else:
                            print("견적서: 필터링 결과 데이터가 없습니다.")
                            # 빈 DataFrame으로 초기화하여 메모리 사용량 최소화
                            self.renewals_df = pd.DataFrame()
                            print(f"견적서: 빈 데이터로 초기화")
                    
                elif criteria == "FCST" and not self.renewals_df.empty and 'FCST' in self.renewals_df.columns:
                    print(f"견적서: 예상 발주일(FCST) 컬럼 확인 - 샘플 데이터: {self.renewals_df['FCST'].head(3).tolist()}")
                    
                    # FCST 컬럼을 datetime으로 변환
                    self.renewals_df['FCST_dt'] = pd.to_datetime(self.renewals_df['FCST'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['FCST_dt'].notna()].copy()
                    print(f"견적서: 유효한 예상 발주일 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 다음달의 마지막 날 계산
                        if next_month.month == 12:
                            end_of_next_month = next_month.replace(year=next_month.year + 1, month=1, day=1) - timedelta(days=1)
                        else:
                            end_of_next_month = next_month.replace(month=next_month.month + 1, day=1) - timedelta(days=1)
                        
                        print(f"견적서: 예상 발주일 다음달 필터링 기간 - {next_month.strftime('%Y-%m-%d')} ~ {end_of_next_month.strftime('%Y-%m-%d')}")
                        
                        # 다음달 데이터 필터링 (시간 무시)
                        valid_data['FCST_날짜만'] = valid_data['FCST_dt'].dt.date
                        start_date = next_month.date() if hasattr(next_month, 'date') else next_month
                        end_date = end_of_next_month.date() if hasattr(end_of_next_month, 'date') else end_of_next_month
                        
                        mask = (valid_data['FCST_날짜만'] >= start_date) & (valid_data['FCST_날짜만'] <= end_date)
                        filtered_data = valid_data[mask].copy()
                        
                        print(f"견적서: 예상 발주일 필터링 후 데이터 수 - {len(filtered_data)}행")
                        
                        # FCST 컬럼을 날짜만 표시하도록 수정
                        filtered_data['FCST'] = filtered_data['FCST_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop(['FCST_dt', 'FCST_날짜만'], axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 예상 발주일 다음달 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 예상 발주일 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                else:
                    print(f"견적서: {criteria} 기준으로 필터링할 수 없습니다.")
                    self.renewals_df = pd.DataFrame()
                
                # 필터 옵션 리스트 업데이트
                self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                
                self.load_filters()
                self.load_data()
            else:
                # 데이터 로더가 없으면 직접 로드
                self._load_data_directly()
                self.load_filters()
                self.load_data()
        except Exception as e:
            print(f"견적서: 다음달 데이터 로드 오류: {e}")
            messagebox.showerror("오류", f"다음달 데이터를 불러오는 중 오류가 발생했습니다:\n{e}")
    
    def load_all_data(self):
        """전체 데이터 로드 (기간 제한 없음)"""
        print("견적서: 전체 데이터 로드 버튼 클릭")
        try:
            # 필터 필드 초기화
            self.customer_cb.set('전체')
            self.client_cb.set('전체')
            self.category_cb.set('전체')
            self.sales_cb.set('전체')
            self.status_cb.set('전체')
            
            # 전체 데이터 가져오기
            if self.data_loader:
                self.renewals_df = self.data_loader.get_quote_data()
                print(f"견적서: 전체 데이터 로드 완료 - {len(self.renewals_df)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if self.renewals_df.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    print(f"견적서: 직접 로드 완료 - {len(self.renewals_df)}행")
                
                if not self.renewals_df.empty:
                    print(f"견적서: 전체 데이터 샘플 (첫 3행):")
                    print(self.renewals_df.head(3).to_string())
                    
                    # 필터 옵션 리스트 업데이트
                    self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                    self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                    self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                    self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                    
                    print(f"견적서: 필터 옵션 업데이트 완료")
                    print(f"- 고객사명: {len(self.all_customers)}개")
                    print(f"- 거래처명: {len(self.all_clients)}개")
                    print(f"- 대분류: {len(self.all_categories)}개")
                    print(f"- 영업사원: {len(self.all_sales)}개")
                    
                    self.load_filters()
                    self.load_data()
                    messagebox.showinfo("완료", f"전체 데이터 {len(self.renewals_df)}행을 로드했습니다.")
                else:
                    print("견적서: 로드할 데이터가 없습니다.")
                    messagebox.showwarning("경고", "로드할 데이터가 없습니다.")
            else:
                # 데이터 로더가 없으면 직접 로드
                self._load_data_directly()
                self.load_filters()
                self.load_data()
        except Exception as e:
            print(f"견적서: 전체 데이터 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"전체 데이터를 불러오는 중 오류가 발생했습니다:\n{e}")

    def load_future_year_data(self):
        """미래 1년 데이터 로드"""
        try:
            print("견적서: 미래 1년 데이터 로드 시작")
            
            # 선택된 날짜 기준으로 미래 1년 계산
            selected_date = self.date_entry.get_date()
            print(f"견적서: 선택된 조회 기준일 - {selected_date.strftime('%Y-%m-%d')}")
            
            start_date = selected_date
            end_date = selected_date + timedelta(days=365)
            
            print(f"견적서: 미래 1년 기간 - {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
            
            # 데이터 로더에서 전체 데이터 가져온 후 필터링
            if self.data_loader:
                self.renewals_df = self.data_loader.get_quote_data()
                print(f"견적서: 전체 데이터 로드 완료 - {len(self.renewals_df)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if self.renewals_df.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    print(f"견적서: 직접 로드 완료 - {len(self.renewals_df)}행")
                
                # 기준 선택에 따른 필터링
                criteria = self.criteria_var.get()
                
                if criteria == "일자" and not self.renewals_df.empty and '일자' in self.renewals_df.columns:
                    print(f"견적서: 일자 기준 미래 1년 필터링")
                    
                    # 일자 컬럼을 datetime으로 변환
                    self.renewals_df['일자_dt'] = pd.to_datetime(self.renewals_df['일자'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['일자_dt'].notna()].copy()
                    print(f"견적서: 유효한 날짜 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 날짜 타입을 datetime64로 통일
                        start_date_dt = pd.to_datetime(start_date)
                        end_date_dt = pd.to_datetime(end_date)
                        
                        # 미래 1년 데이터 필터링 (시간 무시)
                        mask = (valid_data['일자_dt'] >= start_date_dt) & (valid_data['일자_dt'] <= end_date_dt)
                        filtered_data = valid_data[mask].copy()
                        
                        # 일자 컬럼을 날짜만 표시하도록 수정
                        filtered_data['일자'] = filtered_data['일자_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop('일자_dt', axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 일자 기준 미래 1년 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 날짜 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                        
                elif criteria == "FCST" and not self.renewals_df.empty and 'FCST' in self.renewals_df.columns:
                    print(f"견적서: 예상 발주일(FCST) 기준 미래 1년 필터링")
                    
                    # FCST 컬럼을 datetime으로 변환
                    self.renewals_df['FCST_dt'] = pd.to_datetime(self.renewals_df['FCST'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['FCST_dt'].notna()].copy()
                    print(f"견적서: 유효한 예상 발주일 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 날짜 타입을 datetime64로 통일
                        start_date_dt = pd.to_datetime(start_date)
                        end_date_dt = pd.to_datetime(end_date)
                        
                        # 미래 1년 데이터 필터링 (시간 무시)
                        mask = (valid_data['FCST_dt'] >= start_date_dt) & (valid_data['FCST_dt'] <= end_date_dt)
                        filtered_data = valid_data[mask].copy()
                        
                        # FCST 컬럼을 날짜만 표시하도록 수정
                        filtered_data['FCST'] = filtered_data['FCST_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop('FCST_dt', axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 예상 발주일 기준 미래 1년 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 예상 발주일 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                else:
                    print(f"견적서: {criteria} 기준으로 필터링할 수 없습니다.")
                    self.renewals_df = pd.DataFrame()
                
                # 필터 옵션 리스트 업데이트
                self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                
                self.load_filters()
                self.load_data()
            else:
                # 데이터 로더가 없으면 직접 로드
                self._load_data_directly()
                self.load_filters()
                self.load_data()
        except Exception as e:
            print(f"견적서: 미래 1년 데이터 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"미래 1년 데이터를 불러오는 중 오류가 발생했습니다:\n{e}")

    def load_past_year_data(self):
        """과거 1년 데이터 로드"""
        try:
            print("견적서: 과거 1년 데이터 로드 시작")
            
            # 선택된 날짜 기준으로 과거 1년 계산
            selected_date = self.date_entry.get_date()
            print(f"견적서: 선택된 조회 기준일 - {selected_date.strftime('%Y-%m-%d')}")
            
            start_date = selected_date - timedelta(days=365)
            end_date = selected_date
            
            print(f"견적서: 과거 1년 기간 - {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
            
            # 데이터 로더에서 전체 데이터 가져온 후 필터링
            if self.data_loader:
                self.renewals_df = self.data_loader.get_quote_data()
                print(f"견적서: 전체 데이터 로드 완료 - {len(self.renewals_df)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if self.renewals_df.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    print(f"견적서: 직접 로드 완료 - {len(self.renewals_df)}행")
                
                # 기준 선택에 따른 필터링
                criteria = self.criteria_var.get()
                
                if criteria == "일자" and not self.renewals_df.empty and '일자' in self.renewals_df.columns:
                    print(f"견적서: 일자 기준 과거 1년 필터링")
                    
                    # 일자 컬럼을 datetime으로 변환
                    self.renewals_df['일자_dt'] = pd.to_datetime(self.renewals_df['일자'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['일자_dt'].notna()].copy()
                    print(f"견적서: 유효한 날짜 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 날짜 타입을 datetime64로 통일
                        start_date_dt = pd.to_datetime(start_date)
                        end_date_dt = pd.to_datetime(end_date)
                        
                        # 과거 1년 데이터 필터링 (시간 무시)
                        mask = (valid_data['일자_dt'] >= start_date_dt) & (valid_data['일자_dt'] <= end_date_dt)
                        filtered_data = valid_data[mask].copy()
                        
                        # 일자 컬럼을 날짜만 표시하도록 수정
                        filtered_data['일자'] = filtered_data['일자_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop('일자_dt', axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 일자 기준 과거 1년 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 날짜 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                        
                elif criteria == "FCST" and not self.renewals_df.empty and 'FCST' in self.renewals_df.columns:
                    print(f"견적서: 예상 발주일(FCST) 기준 과거 1년 필터링")
                    
                    # FCST 컬럼을 datetime으로 변환
                    self.renewals_df['FCST_dt'] = pd.to_datetime(self.renewals_df['FCST'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = self.renewals_df[self.renewals_df['FCST_dt'].notna()].copy()
                    print(f"견적서: 유효한 예상 발주일 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 날짜 타입을 datetime64로 통일
                        start_date_dt = pd.to_datetime(start_date)
                        end_date_dt = pd.to_datetime(end_date)
                        
                        # 과거 1년 데이터 필터링 (시간 무시)
                        mask = (valid_data['FCST_dt'] >= start_date_dt) & (valid_data['FCST_dt'] <= end_date_dt)
                        filtered_data = valid_data[mask].copy()
                        
                        # FCST 컬럼을 날짜만 표시하도록 수정
                        filtered_data['FCST'] = filtered_data['FCST_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop('FCST_dt', axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 예상 발주일 기준 과거 1년 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 예상 발주일 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                else:
                    print(f"견적서: {criteria} 기준으로 필터링할 수 없습니다.")
                    self.renewals_df = pd.DataFrame()
                
                # 필터 옵션 리스트 업데이트
                self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                
                self.load_filters()
                self.load_data()
            else:
                # 데이터 로더가 없으면 직접 로드
                self._load_data_directly()
                self.load_filters()
                self.load_data()
        except Exception as e:
            print(f"견적서: 과거 1년 데이터 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"과거 1년 데이터를 불러오는 중 오류가 발생했습니다:\n{e}")

    def load_initial_data(self):
        """초기 데이터 로드 - 기준 선택에 따른 필터링"""
        try:
            print("견적서: 초기 데이터 로드 시작")
            
            if self.data_loader:
                # 전체 견적서 데이터 가져오기
                all_quote_data = self.data_loader.get_quote_data()
                print(f"견적서: 전체 데이터 로드 완료 - {len(all_quote_data)}행")
                
                # 데이터가 없으면 직접 로드 시도
                if all_quote_data.empty:
                    print("견적서: 데이터 로더에서 데이터를 가져올 수 없어 직접 로드합니다.")
                    self._load_data_directly()
                    all_quote_data = self.renewals_df
                    print(f"견적서: 직접 로드 완료 - {len(all_quote_data)}행")
                
                # 기준 선택에 따른 필터링
                criteria = self.criteria_var.get()
                print(f"견적서: 조회 기준 - {criteria}")
                
                if criteria == "일자" and not all_quote_data.empty and '일자' in all_quote_data.columns:
                    print(f"견적서: 일자 기준 필터링 - 샘플 데이터: {all_quote_data['일자'].head(3).tolist()}")
                    
                    # 일자 컬럼을 datetime으로 변환 (다양한 형식 지원)
                    all_quote_data['일자_dt'] = pd.to_datetime(all_quote_data['일자'], errors='coerce')
                    
                    # 변환 실패한 데이터 확인
                    failed_conversions = all_quote_data[all_quote_data['일자_dt'].isna()]
                    if not failed_conversions.empty:
                        print(f"견적서: 날짜 변환 실패한 데이터 {len(failed_conversions)}행")
                        print(f"견적서: 변환 실패 샘플: {failed_conversions['일자'].head(3).tolist()}")
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = all_quote_data[all_quote_data['일자_dt'].notna()].copy()
                    print(f"견적서: 유효한 날짜 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 선택된 날짜 기준으로 이번달 계산
                        selected_date = self.date_entry.get_date()
                        start_of_month = selected_date.replace(day=1)
                        if selected_date.month == 12:
                            end_of_month = selected_date.replace(year=selected_date.year + 1, month=1, day=1) - timedelta(days=1)
                        else:
                            end_of_month = selected_date.replace(month=selected_date.month + 1, day=1) - timedelta(days=1)
                        
                        print(f"견적서: 이번달 필터링 기간 - {start_of_month.strftime('%Y-%m-%d')} ~ {end_of_month.strftime('%Y-%m-%d')}")
                        
                        # 이번달 데이터 필터링 (시간 무시)
                        valid_data['일자_날짜만'] = valid_data['일자_dt'].dt.date
                        start_date = start_of_month.date() if hasattr(start_of_month, 'date') else start_of_month
                        end_date = end_of_month.date() if hasattr(end_of_month, 'date') else end_of_month
                        mask = (valid_data['일자_날짜만'] >= start_date) & (valid_data['일자_날짜만'] <= end_date)
                        
                        print(f"견적서: 필터링 마스크 적용 전 데이터 수 - {len(valid_data)}행")
                        print(f"견적서: 필터링 조건 - {start_date} <= 일자 <= {end_date}")
                        print(f"견적서: 필터링 전 데이터 샘플 (첫 5행):")
                        print(valid_data['일자_날짜만'].head(5).tolist())
                        
                        filtered_data = valid_data[mask].copy()
                        
                        print(f"견적서: 필터링 후 데이터 수 - {len(filtered_data)}행")
                        if len(filtered_data) > 0:
                            print(f"견적서: 필터링 후 데이터 샘플 (첫 5행):")
                            print(filtered_data['일자_날짜만'].head(5).tolist())
                        
                        # 일자 컬럼을 날짜만 표시하도록 수정
                        filtered_data['일자'] = filtered_data['일자_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop(['일자_dt', '일자_날짜만'], axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 이번달 필터링 완료 - {len(self.renewals_df)}행")
                        
                        # 디버깅용: 필터링된 데이터 샘플 출력
                        if len(self.renewals_df) > 0:
                            sample_dates = self.renewals_df['일자'].head(5).tolist()
                            print(f"견적서: 필터링된 일자 샘플: {sample_dates}")
                        else:
                            print("견적서: 이번달 데이터가 없습니다.")
                    else:
                        print("견적서: 유효한 날짜 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                        
                elif criteria == "FCST" and not all_quote_data.empty and 'FCST' in all_quote_data.columns:
                    print(f"견적서: 예상 발주일(FCST) 기준 필터링 - 샘플 데이터: {all_quote_data['FCST'].head(3).tolist()}")
                    
                    # FCST 컬럼을 datetime으로 변환
                    all_quote_data['FCST_dt'] = pd.to_datetime(all_quote_data['FCST'], errors='coerce')
                    
                    # 성공적으로 변환된 데이터만 사용
                    valid_data = all_quote_data[all_quote_data['FCST_dt'].notna()].copy()
                    print(f"견적서: 유효한 예상 발주일 데이터 {len(valid_data)}행")
                    
                    if not valid_data.empty:
                        # 선택된 날짜 기준으로 이번달 계산
                        selected_date = self.date_entry.get_date()
                        start_of_month = selected_date.replace(day=1)
                        if selected_date.month == 12:
                            end_of_month = selected_date.replace(year=selected_date.year + 1, month=1, day=1) - timedelta(days=1)
                        else:
                            end_of_month = selected_date.replace(month=selected_date.month + 1, day=1) - timedelta(days=1)
                        
                        print(f"견적서: 예상 발주일 이번달 필터링 기간 - {start_of_month.strftime('%Y-%m-%d')} ~ {end_of_month.strftime('%Y-%m-%d')}")
                        
                        # 이번달 데이터 필터링 (시간 무시)
                        valid_data['FCST_날짜만'] = valid_data['FCST_dt'].dt.date
                        start_date = start_of_month.date() if hasattr(start_of_month, 'date') else start_of_month
                        end_date = end_of_month.date() if hasattr(end_of_month, 'date') else end_of_month
                        mask = (valid_data['FCST_날짜만'] >= start_date) & (valid_data['FCST_날짜만'] <= end_date)
                        
                        filtered_data = valid_data[mask].copy()
                        
                        print(f"견적서: 예상 발주일 필터링 후 데이터 수 - {len(filtered_data)}행")
                        
                        # FCST 컬럼을 날짜만 표시하도록 수정
                        filtered_data['FCST'] = filtered_data['FCST_dt'].dt.strftime('%Y-%m-%d')
                        
                        # 임시 컬럼 제거
                        filtered_data = filtered_data.drop(['FCST_dt', 'FCST_날짜만'], axis=1)
                        
                        self.renewals_df = filtered_data
                        
                        print(f"견적서: 예상 발주일 이번달 필터링 완료 - {len(self.renewals_df)}행")
                    else:
                        print("견적서: 유효한 예상 발주일 데이터가 없습니다.")
                        self.renewals_df = pd.DataFrame()
                else:
                    print(f"견적서: {criteria} 기준으로 필터링할 수 없습니다.")
                    self.renewals_df = pd.DataFrame()
            else:
                print("견적서: 데이터 로더가 없어 기존 방식으로 로드합니다.")
                self._load_data_directly()
                
                # 기존 방식에서도 이번달 필터링 적용
                if not self.renewals_df.empty and '일자' in self.renewals_df.columns:
                    print("견적서: 기존 방식에서 이번달 필터링 적용")
                    self.renewals_df['일자_dt'] = pd.to_datetime(self.renewals_df['일자'], errors='coerce')
                    valid_data = self.renewals_df[self.renewals_df['일자_dt'].notna()].copy()
                    
                    if not valid_data.empty:
                        today = datetime.now()
                        start_of_month = today.replace(day=1)
                        if today.month == 12:
                            end_of_month = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
                        else:
                            end_of_month = today.replace(month=today.month + 1, day=1) - timedelta(days=1)
                        
                        valid_data['일자_날짜만'] = valid_data['일자_dt'].dt.date
                        start_date = start_of_month.date() if hasattr(start_of_month, 'date') else start_of_month
                        end_date = end_of_month.date() if hasattr(end_of_month, 'date') else end_of_month
                        mask = (valid_data['일자_날짜만'] >= start_date) & (valid_data['일자_날짜만'] <= end_date)
                        filtered_data = valid_data[mask].copy()
                        
                        filtered_data = filtered_data.drop(['일자_dt', '일자_날짜만'], axis=1)
                        self.renewals_df = filtered_data
                        print(f"견적서: 기존 방식 이번달 필터링 완료 - {len(self.renewals_df)}행")
            
            # 필터 옵션 리스트 업데이트 (전체 데이터에서 추출)
            if self.data_loader:
                all_data = self.data_loader.get_quote_data()
            else:
                all_data = self.renewals_df  # 기존 방식에서는 필터링된 데이터 사용
                
            self.all_customers = sorted(all_data['고객사명'].dropna().unique().tolist()) if '고객사명' in all_data.columns else []
            self.all_clients   = sorted(all_data['거래처명'].dropna().unique().tolist()) if '거래처명' in all_data.columns else []
            self.all_categories = sorted(all_data['대분류'].dropna().unique().tolist()) if '대분류' in all_data.columns else []
            self.all_sales     = sorted(all_data['영업사원'].dropna().unique().tolist()) if '영업사원' in all_data.columns else []
            
            print(f"견적서: 초기 데이터 로드 완료 - {len(self.renewals_df)}행")
            
        except Exception as e:
            print(f"견적서: 초기 데이터 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            self.renewals_df = pd.DataFrame()
            self.all_customers = []
            self.all_clients = []
            self.all_categories = []
            self.all_sales = []

    def load_data(self):
        """필터링된 데이터를 테이블에 표시"""
        try:
            print(f"견적서: load_data 시작 - renewals_df 크기: {len(self.renewals_df)}행")
            
            # 기존 데이터 삭제
            for item in self.purchase_tree.get_children():
                self.purchase_tree.delete(item)
            
            # 필터링된 데이터 표시 (filtered_df가 있으면 사용, 없으면 renewals_df 사용)
            if hasattr(self, 'filtered_df') and self.filtered_df is not None and not self.filtered_df.empty:
                df = self.filtered_df
                print(f"견적서: filtered_df 사용 - {len(df)}행")
            else:
                df = self.renewals_df
                print(f"견적서: renewals_df 사용 - {len(df)}행")
                # filtered_df 속성 정리
                if hasattr(self, 'filtered_df'):
                    delattr(self, 'filtered_df')
            
            if df.empty:
                print("견적서: 표시할 데이터가 없습니다.")
                return
            
            # 일자 기준으로 최신순 정렬 (기본 정렬)
            if '일자' in df.columns:
                try:
                    # 일자 컬럼을 datetime으로 변환하여 정확한 정렬
                    df['일자_정렬용'] = pd.to_datetime(df['일자'], errors='coerce')
                    df = df.sort_values('일자_정렬용', ascending=False, na_position='last')
                    df = df.drop('일자_정렬용', axis=1)
                    print(f"견적서: 데이터를 일자 기준 최신순으로 정렬했습니다. 총 {len(df)}행")
                except Exception as e:
                    print(f"견적서: 일자 정렬 오류: {e}")
            
            # 데이터를 테이블에 추가
            row_count = 0
            for _, row in df.iterrows():
                values = []
                for col in self.purchase_cols:
                    val = row.get(col, "")
                    if col in self.cost_cols and val:
                        try:
                            val = f"{int(str(val).replace(',', '')):,}"
                        except:
                            pass
                    elif col in ['이익율','실 이익율']:
                        try:
                            val = f"{float(str(val)):.1f}%"
                        except:
                            pass
                    values.append(val)
                
                self.purchase_tree.insert('', 'end', values=values)
                row_count += 1
            
            print(f"견적서: 데이터 표시 완료 - {row_count}행")
            
            # 데이터 로드 후 트리뷰 우클릭 이벤트 바인딩 재설정
            try:
                # 기존 바인딩 제거
                self.purchase_tree.unbind("<Button-3>")
                # 새로운 바인딩 설정
                self.purchase_tree.bind("<Button-3>", lambda event: show_status_context_menu(event, self.purchase_tree, self.status_modifier))
                print("견적서: 트리뷰 우클릭 이벤트 바인딩 재설정 완료")
            except Exception as e:
                print(f"견적서: 트리뷰 이벤트 바인딩 재설정 실패: {e}")
                
        except Exception as e:
            print(f"견적서: 데이터 표시 오류: {e}")
            import traceback
            traceback.print_exc()

    def update_customer_list(self):
        """고객사 필터 자동 완성"""
        val = self.customer_cb.get().strip()
        filtered = (
            self.all_customers
            if not val or val == '전체'
            else [c for c in self.all_customers if val.lower() in c.lower()]
        )
        self.customer_cb['values'] = ['전체'] + filtered

    def update_client_list(self):
        """거래처 필터 자동 완성"""
        val = self.client_cb.get().strip()
        filtered = (
            self.all_clients
            if not val or val == '전체'
            else [c for c in self.all_clients if val.lower() in c.lower()]
        )
        self.client_cb['values'] = ['전체'] + filtered

    def update_category_list(self):
        """대분류 필터 자동 완성"""
        val = self.category_cb.get().strip()
        filtered = (
            self.all_categories
            if not val or val == '전체'
            else [c for c in self.all_categories if val.lower() in c.lower()]
        )
        self.category_cb['values'] = ['전체'] + filtered

    def update_sales_list(self):
        """영업사원 필터 자동 완성"""
        val = self.sales_cb.get().strip()
        filtered = (
            self.all_sales
            if not val or val == '전체'
            else [c for c in self.all_sales if val.lower() in c.lower()]
        )
        self.sales_cb['values'] = ['전체'] + filtered

    def update_status_list(self):
        """Status 필터 자동 완성"""
        val = self.status_cb.get().strip()
        status_options = ['100%', '80%', '60%', '40%', '20%', 'Drop']
        filtered = (
            status_options
            if not val or val == '전체'
            else [s for s in status_options if val.lower() in s.lower()]
        )
        self.status_cb['values'] = ['전체'] + filtered

    def on_customer_search(self):
        """고객사명 콤보박스에서 엔터키로 검색"""
        entered_customer = self.customer_cb.get().strip()
        if entered_customer:
            # 디버깅용 로그 추가
            print(f"고객사 검색: '{entered_customer}', 전체 고객사 수: {len(self.all_customers)}")
            
            if not self.all_customers:
                print("경고: 고객사 목록이 비어있습니다!")
                return
                
            # 고객사명으로 검색
            matching_customers = [c for c in self.all_customers if entered_customer.lower() in c.lower()]
            print(f"매칭되는 고객사: {matching_customers}")
            
            if matching_customers:
                # 검색 결과를 콤보박스에 설정
                self.customer_cb['values'] = ['전체'] + matching_customers
                # 첫 번째 매칭되는 고객사명을 설정
                self.customer_cb.set(matching_customers[0])
                # 드롭다운 리스트 표시
                self.customer_cb.event_generate('<Down>')
                self.customer_cb.event_generate('<Up>')
            else:
                # 매칭되는 고객사가 없으면 입력값 그대로 사용
                self.customer_cb.set(entered_customer)

    def on_client_search(self):
        """거래처명 콤보박스에서 엔터키로 검색"""
        entered_client = self.client_cb.get().strip()
        if entered_client:
            # 디버깅용 로그 추가
            print(f"거래처 검색: '{entered_client}', 전체 거래처 수: {len(self.all_clients)}")
            
            if not self.all_clients:
                print("경고: 거래처 목록이 비어있습니다!")
                return
                
            # 거래처명으로 검색
            matching_clients = [c for c in self.all_clients if entered_client.lower() in c.lower()]
            print(f"매칭되는 거래처: {matching_clients}")
            
            if matching_clients:
                # 검색 결과를 콤보박스에 설정 (모든 매칭 항목 표시)
                self.client_cb['values'] = ['전체'] + matching_clients
                # 입력값 그대로 유지 (첫 번째 항목으로 자동 설정하지 않음)
                self.client_cb.set(entered_client)
                # 드롭다운 리스트 표시
                self.client_cb.event_generate('<Down>')
                self.client_cb.event_generate('<Up>')
            else:
                # 매칭되는 거래처가 없으면 입력값 그대로 사용
                self.client_cb.set(entered_client)

    def on_search(self):
        """고객사명, 거래처명 기준 데이터 필터링 및 UI 갱신"""
        try:
            if not self.parent_frame.winfo_exists():
                return
        except:
            return
        
        sel_cust   = self.customer_cb.get().strip()
        sel_client = self.client_cb.get().strip()
        sel_cat    = self.category_cb.get().strip()
        sel_sales  = self.sales_cb.get().strip()
        sel_status = self.status_cb.get().strip()
        
        # 구매정보 필터링 (조회 기간 필터링된 데이터 사용)
        if hasattr(self, 'filtered_df') and self.filtered_df is not None and not self.filtered_df.empty:
            df_pur = self.filtered_df.copy()
            print(f"견적서: on_search에서 filtered_df 사용 - {len(df_pur)}행")
        else:
            df_pur = self.renewals_df.copy()
            print(f"견적서: on_search에서 renewals_df 사용 - {len(df_pur)}행")
            
        if sel_cust and sel_cust != '전체':
            df_pur = df_pur[df_pur['고객사명'].str.contains(sel_cust, na=False, case=False, regex=False)]
        if sel_client and sel_client != '전체':
            df_pur = df_pur[df_pur['거래처명'].str.contains(sel_client, na=False, case=False, regex=False)]
        if sel_cat and sel_cat != '전체' and '대분류' in df_pur.columns:
            df_pur = df_pur[df_pur['대분류'].str.contains(sel_cat, na=False, case=False, regex=False)]
        if sel_sales and sel_sales != '전체' and '영업사원' in df_pur.columns:
            df_pur = df_pur[df_pur['영업사원'].str.contains(sel_sales, na=False, case=False, regex=False)]
        if sel_status and sel_status != '전체' and 'Status' in df_pur.columns:
            df_pur = df_pur[df_pur['Status'].str.contains(sel_status, na=False, case=False, regex=False)]

        # 실 매입 단가 우선 교체
        if '실 매입 단가' in df_pur.columns:
            mask = (df_pur['실 매입 단가'].notna()
                    & (df_pur['실 매입 단가'] != df_pur['원가']))
            cols_src = ['실 매입 단가','실 매입 합계','실 이익액','실 이익율']
            cols_dst = ['원가','원가계','이익액','이익율']
            values = df_pur.loc[mask, cols_src].copy()
            for col in values.columns:
                values[col] = pd.to_numeric(values[col], errors='coerce')
            for col in cols_dst:
                df_pur[col] = pd.to_numeric(df_pur[col], errors='coerce')
            df_pur.loc[mask, cols_dst] = values.values

        # 일자 기준으로 최신순 정렬 (기본 정렬)
        if not self.sort_column and '일자' in df_pur.columns:
            try:
                # 일자 컬럼을 datetime으로 변환하여 정확한 정렬
                df_pur['일자_정렬용'] = pd.to_datetime(df_pur['일자'], errors='coerce')
                df_pur = df_pur.sort_values('일자_정렬용', ascending=False, na_position='last')
                df_pur = df_pur.drop('일자_정렬용', axis=1)
                print(f"검색 결과를 일자 기준 최신순으로 정렬했습니다. 총 {len(df_pur)}행")
            except Exception as e:
                print(f"일자 정렬 오류: {e}")
        elif self.sort_column:
            df_pur = df_pur.sort_values(
                by=self.sort_column,
                ascending=not self.sort_reverse,
                key=lambda x: x.fillna('') if x.dtype==object else x
            )

        try:
            if self.purchase_tree.winfo_exists():
                self.purchase_tree.delete(*self.purchase_tree.get_children())
                for _, row in df_pur.iterrows():
                    vals = []
                    for col in self.purchase_cols:
                        v = row.get(col, "")
                        if col in self.cost_cols:  # 천 단위 콤마
                            try:    v = f"{int(v):,}"
                            except: pass
                        elif col in ['이익율','실 이익율']:
                            if v in [None, '', 'NaN', 'nan'] or (isinstance(v, float) and pd.isna(v)):
                                v = ''
                            else:
                                try:
                                    v = f"{float(v):.2f}"
                                except:
                                    pass
                        vals.append(v)
                    self.purchase_tree.insert('', 'end', values=vals)
        except:
            pass

        # 정렬 상태는 sort_purchase 함수에서만 업데이트
        # self.sort_column = col
        # self.sort_reverse = rev

    def sort_purchase(self, col):
        try:
            if not self.parent_frame.winfo_exists():
                return
        except:
            return
        
        sel_cust   = self.customer_cb.get().strip()
        sel_client = self.client_cb.get().strip()
        df = self.renewals_df.copy()
        if sel_cust and sel_cust != '전체':
            df = df[df['고객사명'].str.contains(sel_cust, na=False, case=False, regex=False)]
        if sel_client and sel_client != '전체':
            df = df[df['거래처명'].str.contains(sel_client, na=False, case=False, regex=False)]

        rev = (self.sort_column==col) and not self.sort_reverse
        
        # 일자 컬럼인 경우 datetime으로 변환하여 정확한 정렬
        if col == '일자':
            try:
                df_sorted = df.copy()
                df_sorted['일자_정렬용'] = pd.to_datetime(df_sorted['일자'], errors='coerce')
                df_sorted = df_sorted.sort_values('일자_정렬용', ascending=not rev, na_position='last')
                df_sorted = df_sorted.drop('일자_정렬용', axis=1)
            except Exception as e:
                print(f"일자 컬럼 정렬 오류: {e}")
                df_sorted = df.sort_values(
                    by=col,
                    ascending=not rev,
                    key=lambda x: x.fillna('') if x.dtype==object else x
                )
        else:
            df_sorted = df.sort_values(
                by=col,
                ascending=not rev,
                key=lambda x: x.fillna('') if x.dtype==object else x
            )

        # 일자 기준으로 최신순 정렬 (사용자가 다른 컬럼으로 정렬한 경우에도 일자 기준 보조 정렬)
        if col != '일자' and '일자' in df_sorted.columns:
            try:
                # 일자 컬럼을 datetime으로 변환하여 정확한 정렬
                df_sorted['일자_정렬용'] = pd.to_datetime(df_sorted['일자'], errors='coerce')
                df_sorted = df_sorted.sort_values(['일자_정렬용', col], ascending=[False, not rev], na_position='last')
                df_sorted = df_sorted.drop('일자_정렬용', axis=1)
            except Exception as e:
                print(f"일자 보조 정렬 오류: {e}")

        if '실 매입 단가' in df_sorted.columns:
            mask = (df_sorted['실 매입 단가'].notna()
                    & (df_sorted['실 매입 단가'] != df_sorted['원가']))
            cols_src = ['실 매입 단가','실 매입 합계','실 이익액','실 이익율']
            cols_dst = ['원가','원가계','이익액','이익율']
            values = df_sorted.loc[mask, cols_src].copy()
            for col2 in values.columns:
                values[col2] = pd.to_numeric(values[col2], errors='coerce')
            for col2 in cols_dst:
                df_sorted[col2] = pd.to_numeric(df_sorted[col2], errors='coerce')
            df_sorted.loc[mask, cols_dst] = values.values

        try:
            if self.purchase_tree.winfo_exists():
                self.purchase_tree.delete(*self.purchase_tree.get_children())
                for _, row in df_sorted.iterrows():
                    vals = []
                    for c in self.purchase_cols:
                        v = row.get(c, "")
                        if c in self.cost_cols:
                            try:    v = f"{int(v):,}"
                            except: pass
                        elif c in ['이익율','실 이익율']:
                            if v in [None, '', 'NaN', 'nan'] or (isinstance(v, float) and pd.isna(v)):
                                v = ''
                            else:
                                try:
                                    v = f"{float(v):.2f}"
                                except:
                                    pass
                        vals.append(v)
                    self.purchase_tree.insert('', 'end', values=vals)
        except:
            pass

        self.sort_column  = col
        self.sort_reverse = rev

    # 기존 컬럼 크기 조절 이벤트 핸들러 제거됨 (ResizableTreeview로 대체)

    def reload_customers(self):
        """구글시트에서 고객사 목록 재로딩 후 필터 갱신 (최적화된 방식)"""
        try:
            # 전역 인증 사용
            gc = get_global_auth()
            
            def load_customer_data():
                sh = gc.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
                ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
                return ws.get_all_values()
            
            all_values = safe_api_call(load_customer_data)
            if all_values:
                raw = [h.strip() for h in all_values[0]]
                idxs = [i for i, h in enumerate(raw) if h]
                hdrs = pd.Index([raw[i] for i in idxs])
                rows = [[r[i] for i in idxs] for r in all_values[1:]]
                df_all = pd.DataFrame(rows, columns=hdrs)
                df_all.columns = df_all.columns.str.strip()
                if '고객사명' in df_all.columns:
                    self.contacts_df = df_all.drop_duplicates(subset='고객사명')
                else:
                    self.contacts_df = pd.DataFrame()
            else:
                self.contacts_df = pd.DataFrame()
            self.contacts_df.columns = self.contacts_df.columns.str.strip()
            self.all_customers = sorted(self.contacts_df['고객사명'].dropna().unique().tolist())
            self.load_filters()
        except Exception as e:
            print(f"고객사 목록 재로딩 오류: {e}")
            self.contacts_df = pd.DataFrame()
            self.all_customers = []
            self.load_filters()

    def refresh_quote_data(self):
        """견적서 데이터 새로고침 - 최적화된 방식"""
        try:
            print("견적서: 견적서 데이터 새로고침 시작")
            
            # 데이터 로더가 있으면 강제로 새로 로드
            if self.data_loader:
                print("견적서: 데이터 로더를 통해 최신 견적서 데이터 로드 중...")
                
                # 데이터 로더의 견적서 데이터 강제 재로드
                self.renewals_df = self.data_loader.force_reload_quote_data()
                
                print(f"견적서: 견적서 데이터 로드 완료 - renewals_df 크기: {len(self.renewals_df)}행")
                if not self.renewals_df.empty:
                    print(f"견적서: 컬럼 목록: {list(self.renewals_df.columns)}")
                    print(f"견적서: 최신 견적서 데이터 샘플 (첫 3행):")
                    print(self.renewals_df.head(3).to_string())
                    
                    # 필터 옵션 리스트 업데이트
                    self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                    self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                    self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                    self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                    
                    print(f"견적서: 필터 옵션 업데이트 완료")
                    print(f"- 고객사명: {len(self.all_customers)}개")
                    print(f"- 거래처명: {len(self.all_clients)}개")
                    print(f"- 대분류: {len(self.all_categories)}개")
                    print(f"- 영업사원: {len(self.all_sales)}개")
                    
                    # 필터 콤보박스 업데이트
                    self.load_filters()
                    
                    # 데이터 표시
                    self.load_data()
                    
                    print(f"견적서: 견적서 데이터 새로고침 완료 - 총 {len(self.renewals_df)}행")
                else:
                    print("견적서: 견적서 데이터 로드 실패")
                    messagebox.showerror("오류", "견적서 데이터를 불러올 수 없습니다.")
            else:
                # 데이터 로더가 없으면 기존 방식
                self._load_data_directly()
                self.load_filters()
                self.load_data()
                
        except Exception as e:
            print(f"견적서: 견적서 데이터 새로고침 오류: {e}")
            messagebox.showerror("오류", f"견적서 데이터 새로고침 중 오류가 발생했습니다:\n{e}")

    def on_full_data_loaded(self):
        """전체 데이터 로딩 완료 시 호출되는 메서드"""
        try:
            print("견적서: 전체 데이터 로딩 완료, 데이터 새로고침 중...")
            # 데이터 로더에서 최신 데이터 가져오기
            if self.data_loader:
                quote_data = self.data_loader.get_quote_data()
                if not quote_data.empty:
                    self.renewals_df = quote_data
                    # 필터 옵션 리스트 업데이트
                    self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
                    self.all_clients   = sorted(self.renewals_df['거래처명'].dropna().unique().tolist()) if '거래처명' in self.renewals_df.columns else []
                    self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist()) if '대분류' in self.renewals_df.columns else []
                    self.all_sales     = sorted(self.renewals_df['영업사원'].dropna().unique().tolist()) if '영업사원' in self.renewals_df.columns else []
                    
                    # 필터 콤보박스 업데이트
                    self.load_filters()
                    
                    # 데이터 표시 (일자 기준 최신순 정렬 포함)
                    self.load_data()
                    
                    print(f"견적서: 전체 데이터 로딩 완료 후 데이터 새로고침 완료. 총 {len(self.renewals_df)}행")
                else:
                    print("견적서: 전체 데이터 로딩 완료했지만 quote_data가 비어있습니다.")
            else:
                print("견적서: 데이터 로더가 없어 전체 데이터 로딩 후 처리를 건너뜁니다.")
        except Exception as e:
            print(f"견적서 전체 데이터 로딩 완료 처리 오류: {e}") 