import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import re
import hashlib
from core.sendlog import KEY_FILE, SCOPES
from core.data_loader import get_global_auth, safe_api_call

def resource_path(relative_path):
    import sys
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

# 제품명에서 줄바꿈 문자 제거 함수 추가
def clean_product_name(name):
    return str(name).replace('\n', ' ').replace('\r', ' ').replace('\u2028', ' ').replace('\u2029', ' ')

class QuoteModifyDialog:
    """
    견적 작성(수정) 다이얼로그: 견적 하기(수정) 표 UI 생성
    """
    def __init__(self, parent, rows_data=None, username=None, callback=None):
        self.parent = parent
        # 단일 행 데이터를 리스트로 변환 (하위 호환성)
        if rows_data is None:
            self.rows_data = []
        elif isinstance(rows_data, dict):
            self.rows_data = [rows_data]
        else:
            self.rows_data = rows_data
        self.username = username or '테스트유저'
        self.callback = callback
        
        # 데이터 초기화
        self.price_df = pd.DataFrame()
        self.categories = []
        self.product_costs = {}
        self.customers_df = pd.DataFrame()
        self.all_customers = []
        self.all_clients = []
        
        # 견적 행 데이터
        self.quote_rows = []
        
        # 행 선택 관련 초기화
        self.selected_row_index = None
        
        # Status 변수 초기화
        self.status_var = tk.StringVar(value="60%")
        
        # 데이터 로드
        self.load_price_data()
        self.load_customer_data()
        
        # UI 생성
        self.win = tk.Toplevel(parent)
        self.win.title("견적 작성 (수정)")
        self.win.geometry("1356x800")  # 창 크기 조정
        self.win.configure(bg="#F7F9FB")
        
        # 모달 설정
        self.win.transient(parent)
        self.win.grab_set()
        
        self.setup_ui()
        self.center_window()

    def center_window(self):
        """창을 화면 중앙에 배치"""
        try:
            if not self.win:
                return
                
            # 창이 유효한지 확인
            if not self.win.winfo_exists():
                return
                
            self.win.update_idletasks()
            
            # 창이 파괴되었는지 다시 한번 확인
            if not self.win.winfo_exists():
                return
                
            width = self.win.winfo_width()
            height = self.win.winfo_height()
            
            # 유효한 크기인지 확인
            if width <= 0 or height <= 0:
                return
                
            x = (self.win.winfo_screenwidth() // 2) - (width // 2)
            y = (self.win.winfo_screenheight() // 2) - (height // 2)
            self.win.geometry(f"{width}x{height}+{x}+{y}")
            
        except tk.TclError as e:
            print(f"견적 작성(수정): 창 중앙 배치 중 오류 발생 - {e}")
        except Exception as e:
            print(f"견적 작성(수정): 창 중앙 배치 중 예상치 못한 오류 - {e}")

    def load_price_data(self):
        """price_list에서 대분류, 제품, 원가 정보 로드 (최적화된 방식)"""
        try:
            print("견적 수정: 가격 데이터 로딩 시작...")
            
            # 전역 인증 사용
            gc = get_global_auth()
            
            def load_price():
                print("견적 수정: Google Sheets 연결 중...")
                sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
                ws = sh.worksheet('Sheet1')
                print("견적 수정: 데이터 읽기 중...")
                # 헤더 중복 문제 해결을 위해 expected_headers 사용
                expected_headers = ['대분류', '소분류', '제품', '상세 제품명', '구분', 'USD', 'EUR', 'DC', '원가', '매입처']
                data = ws.get_all_records(expected_headers=expected_headers)
                print(f"견적 수정: {len(data)}행 데이터 읽기 완료")
                return data
            
            # 직접 API 호출 (타임아웃 로직 제거)
            print("견적 수정: safe_api_call 호출 중...")
            price_data = safe_api_call(load_price)
            
            if price_data is None:
                print("견적 수정: safe_api_call이 None을 반환했습니다.")
                # 직접 호출 시도
                try:
                    print("견적 수정: 직접 API 호출 시도...")
                    price_data = load_price()
                except Exception as direct_error:
                    print(f"견적 수정: 직접 API 호출도 실패: {direct_error}")
                    price_data = None
            
            if price_data:
                print("견적 수정: 데이터프레임 변환 중...")
                self.price_df = pd.DataFrame(price_data)
                self.price_df.columns = self.price_df.columns.str.strip()
                
                # 대분류 목록 (중복 제거)
                self.categories = sorted(self.price_df['대분류'].dropna().unique().tolist())
                # 소분류 목록 (중복 제거)
                self.subcategories = sorted(self.price_df['소분류'].dropna().unique().tolist())
                
                # 제품별 원가 딕셔너리
                self.product_costs = {}
                for _, row in self.price_df.iterrows():
                    product = row.get('제품', '')
                    cost = row.get('원가', 0)
                    if product and cost:
                        self.product_costs[product] = cost
            else:
                self.price_df = pd.DataFrame()
                self.categories = []
                self.subcategories = []
                self.product_costs = {}
                    
        except Exception as e:
            print(f"price_list 로드 오류: {e}")
            self.price_df = pd.DataFrame()
            self.categories = []
            self.subcategories = []
            self.product_costs = {}

    def get_subcategories_by_category(self, category):
        """대분류에 따른 소분류 목록 반환"""
        if self.price_df.empty:
            return []
        subcategories = self.price_df[self.price_df['대분류'] == category]['소분류'].dropna().unique().tolist()
        return sorted(subcategories)

    def get_products_by_subcategory(self, subcategory):
        """소분류에 따른 제품 목록 반환"""
        if self.price_df.empty:
            return []
        products = self.price_df[self.price_df['소분류'] == subcategory]['제품'].dropna().unique().tolist()
        return sorted(products)

    def get_products_by_category(self, category):
        """대분류에 따른 제품 목록 반환 (하위 호환성)"""
        if self.price_df.empty:
            return []
        products = self.price_df[self.price_df['대분류'] == category]['제품'].dropna().unique().tolist()
        return sorted(products)

    def load_customer_data(self):
        """renewal_list에서 고객사명과 거래처명 데이터 로드 (최적화된 방식)"""
        try:
            # 전역 인증 사용
            gc = get_global_auth()
            
            def load_customer():
                sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
                ws = sh.worksheet('Sheet1')
                return ws.get_all_records()
            
            customer_data = safe_api_call(load_customer)
            if customer_data:
                df = pd.DataFrame(customer_data)
                df.columns = df.columns.str.strip()
                if '고객사명' in df.columns and '거래처명' in df.columns:
                    self.customers_df = df[['고객사명', '거래처명']].drop_duplicates()
                else:
                    self.customers_df = pd.DataFrame(columns=['고객사명', '거래처명'])
                self.all_customers = sorted(self.customers_df['고객사명'].dropna().unique().tolist())
                self.all_clients = sorted(self.customers_df['거래처명'].dropna().unique().tolist())
                
                # 디버깅용 로그 추가
                print(f"고객사 데이터 로드 완료: 고객사 {len(self.all_customers)}개, 거래처 {len(self.all_clients)}개")
                if len(self.all_clients) > 0:
                    print(f"거래처 목록 샘플: {self.all_clients[:5]}")
            else:
                self.customers_df = pd.DataFrame(columns=['고객사명', '거래처명'])
                self.all_customers = []
                self.all_clients = []
            
        except Exception as e:
            print(f"고객사 데이터 로드 오류: {e}")
            self.customers_df = pd.DataFrame(columns=['고객사명', '거래처명'])
            self.all_customers = []
            self.all_clients = []

    def setup_ui(self):
        self.quote_rows = []
        self.create_exchange_tables_side_by_side()
        self.create_customer_table()
        
        tk.Label(self.win, text="견적 하기(수정)", font=("맑은 고딕", 9, "bold")).pack(anchor='w', padx=10, pady=(20,0))
        self.table_frame = tk.Frame(self.win)
        self.table_frame.pack(fill='x', padx=10, pady=(0,10))
        self.create_quote_headers()
        # 여러 제품 데이터 표시
        if self.rows_data:
            self.add_row(self.rows_data)
        else:
            self.add_row([])  # 빈 행 추가
        btn_frame = tk.Frame(self.win)
        btn_frame.pack(side='bottom', pady=10)
        def style_btn(btn, bg, hover_bg):
            btn.configure(font=("맑은 고딕", 10, "bold"), fg="white", bg=bg, relief="flat", padx=12, pady=3)
            btn.bind('<Enter>', lambda e: btn.configure(bg=hover_bg))
            btn.bind('<Leave>', lambda e: btn.configure(bg=bg))
        # 행 제어 버튼들 (행 추가 버튼 왼쪽에 배치)
        up_btn = tk.Button(btn_frame, text="⬆️ 선택 행 위로", command=self.move_selected_row_up)
        up_btn.pack(side='left', padx=5)
        style_btn(up_btn, "#3B82F6", "#2563EB")
        
        down_btn = tk.Button(btn_frame, text="⬇️ 선택 행 아래로", command=self.move_selected_row_down)
        down_btn.pack(side='left', padx=5)
        style_btn(down_btn, "#3B82F6", "#2563EB")
        
        delete_btn = tk.Button(btn_frame, text="🗑️ 선택 행 삭제", command=self.delete_selected_row)
        delete_btn.pack(side='left', padx=5)
        style_btn(delete_btn, "#EF4444", "#DC2626")
        
        # 행 추가 버튼
        add_btn = tk.Button(btn_frame, text="행 추가", command=self.add_new_row)
        add_btn.pack(side='left', padx=5)
        style_btn(add_btn, "#3B82F6", "#2563EB")
        
        # 행 삭제 버튼
        del_btn = tk.Button(btn_frame, text="행 삭제", command=self.remove_row)
        del_btn.pack(side='left', padx=5)
        style_btn(del_btn, "#F59E42", "#EA580C")
        def open_product_add():
            try:
                from quote_product_add import QuoteProductAddWindow
                QuoteProductAddWindow(self.win)
            except Exception as e:
                import traceback
                traceback.print_exc()
                from tkinter import messagebox
                messagebox.showerror('오류', f'제품 추가 창을 여는 데 실패했습니다:\n{e}')
        prod_btn = tk.Button(btn_frame, text="제품 추가", command=open_product_add)
        prod_btn.pack(side='left', padx=5)
        style_btn(prod_btn, "#10B981", "#059669")
        save_btn = tk.Button(btn_frame, text="저장 및 견적서 작성", width=20, command=self.on_save)
        save_btn.pack(side='left', padx=10)
        style_btn(save_btn, "#8B5CF6", "#7C3AED")
        cancel_btn = tk.Button(btn_frame, text="취소", width=10, command=self.win.destroy)
        cancel_btn.pack(side='left', padx=10)
        style_btn(cancel_btn, "#6B7280", "#374151")
        self.win.transient(self.parent)
        self.win.grab_set()
        self.win.wait_window(self.win)

    def create_status_selection_frame(self):
        """Status 선택 프레임 생성"""
        status_frame = tk.LabelFrame(self.win, text="Status 선택", font=("맑은 고딕", 9, "bold"), bg="white", relief="solid", bd=1, labelanchor='nw')
        status_frame.pack(pady=(10, 10), padx=10, fill='x')
        
        # Status 선택 드롭다운
        tk.Label(status_frame, text="Status:", font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left', padx=(15, 5), pady=15)
        
        self.status_var = tk.StringVar(value="60%")
        status_combo = ttk.Combobox(status_frame, textvariable=self.status_var, width=20, font=("맑은 고딕", 10), state="readonly")
        status_combo['values'] = ['100%', '80%', '60%', '40%', '20%', 'Drop']
        status_combo.pack(side='left', padx=(0, 15), pady=15)

    def create_quote_headers(self):
        self.col_widths = [10, 10, 50, 3, 10, 13, 13, 10, 13, 10, 5, 10]  # 예상 발주일 너비 제거
        headers = ['대분류', '소분류', '제품', '수량', '판매 단가', '판매 합계', '만료일', '원가', '원가계', '이익액', '이익율', '매입처']
        self.header_labels = []
        for i, header in enumerate(headers):
            lbl = tk.Label(self.table_frame, text=header, bg='#F1F5F9', fg="#374151", relief='solid', font=("맑은 고딕", 10, "bold"), width=max(self.col_widths[i]//8, 8), height=2, anchor='center')
            lbl.grid(row=0, column=i, sticky='ew')
            self.table_frame.grid_columnconfigure(i, minsize=self.col_widths[i])
            lbl.bind('<Button-1>', lambda e, idx=i: self.start_resize(e, idx))
            lbl.bind('<B1-Motion>', lambda e, idx=i: self.do_resize(e, idx))
            lbl.bind('<ButtonRelease-1>', lambda e, idx=i: self.end_resize(e, idx))
            self.header_labels.append(lbl)
        self._resize_col = None
        self._resize_start_x = None

    def start_resize(self, event, col_idx):
        self._resize_col = col_idx
        self._resize_start_x = event.x_root

    def do_resize(self, event, col_idx):
        if self._resize_col is None:
            return
        delta = event.x_root - self._resize_start_x
        if abs(delta) >= 15:
            step = delta // 15
            new_width = max(6, min(40, self.col_widths[col_idx] + step))
            self.col_widths[col_idx] = new_width
            self.header_labels[col_idx].config(width=new_width)
            for row in self.quote_rows:
                widget = self.get_widget_by_col(row, col_idx)
                if widget:
                    widget.config(width=new_width)

    def end_resize(self, event, col_idx):
        self._resize_col = None
        self._resize_start_x = None

    def get_widget_by_col(self, row_widgets, col_idx):
        col_keys = ['대분류', '소분류', '제품', '수량', '판매 단가', '판매 합계', '만료일', '원가', '원가계', '이익액', '이익율', '매입처']
        key = col_keys[col_idx]
        if key == '대분류':
            return row_widgets.get('대분류_combo')
        if key == '소분류':
            return row_widgets.get('소분류_combo')
        if key == '제품':
            return row_widgets.get('제품_combo')
        return row_widgets.get(f'{key}_entry')

    def add_row(self, past_rows=None):
        # past_rows가 리스트가 아니면 단일 행으로 처리 (하위 호환성)
        if past_rows is None:
            past_rows = []
        elif not isinstance(past_rows, list):
            past_rows = [past_rows]
        
        # 기존 행들을 모두 제거
        for row_widgets in self.quote_rows:
            for widget in row_widgets.values():
                if hasattr(widget, 'destroy'):
                    widget.destroy()
        self.quote_rows.clear()
        
        # 각 제품에 대해 행 추가
        for past_row in past_rows:
            self._add_single_row(past_row)
        
        # 제품이 없으면 빈 행 하나 추가
        if not past_rows:
            self._add_single_row(None)
    
    def add_new_row(self):
        """새로운 행을 기존 행들에 추가 (행 추가 버튼용)"""
        self._add_single_row(None)
    
    def _add_single_row(self, past_row=None):
        row_num = len(self.quote_rows) + 1
        row_widgets = {}
        # 대분류 (드롭다운 + 수동 입력)
        category_var = tk.StringVar()
        category_combo = ttk.Combobox(self.table_frame, textvariable=category_var, values=self.categories, width=self.col_widths[0], state='normal', font=("맑은 고딕", 10))
        category_combo.grid(row=row_num, column=0, sticky='ew', padx=1, pady=1)
        row_widgets['대분류'] = category_var
        row_widgets['대분류_combo'] = category_combo
        category_combo.bind('<<ComboboxSelected>>', lambda e, row=row_widgets: self.on_category_change(row))
        category_combo.bind('<Return>', lambda e, row=row_widgets: self.on_category_change(row))
        # 소분류 (드롭다운 + 수동 입력)
        subcategory_var = tk.StringVar()
        subcategory_combo = ttk.Combobox(self.table_frame, textvariable=subcategory_var, width=self.col_widths[1], state='normal', font=("맑은 고딕", 10))
        subcategory_combo.grid(row=row_num, column=1, sticky='ew', padx=1, pady=1)
        row_widgets['소분류'] = subcategory_var
        row_widgets['소분류_combo'] = subcategory_combo
        subcategory_combo.bind('<<ComboboxSelected>>', lambda e, row=row_widgets: self.on_subcategory_change(row))
        subcategory_combo.bind('<Return>', lambda e, row=row_widgets: self.on_subcategory_change(row))
        # 제품 (드롭다운 + 수동 입력)
        product_var = tk.StringVar()
        product_combo = ttk.Combobox(self.table_frame, textvariable=product_var, width=self.col_widths[2], state='normal', font=("맑은 고딕", 10))
        product_combo.grid(row=row_num, column=2, sticky='ew', padx=1, pady=1)
        row_widgets['제품'] = product_var
        row_widgets['제품_combo'] = product_combo
        product_combo.bind('<<ComboboxSelected>>', lambda e, row=row_widgets: self.on_product_change(row))
        product_combo.bind('<Return>', lambda e, row=row_widgets: self.on_product_search_and_apply(row))
        product_combo.bind('<KeyRelease>', lambda e, row=row_widgets: self.on_product_key_release(e, row))
        # 수량 (입력, 가운데 정렬)
        qty_var = tk.StringVar()
        qty_entry = tk.Entry(self.table_frame, textvariable=qty_var, width=self.col_widths[3], justify='center', font=("맑은 고딕", 10))
        qty_entry.grid(row=row_num, column=3, sticky='ew', padx=1, pady=1)
        row_widgets['수량'] = qty_var
        row_widgets['수량_entry'] = qty_entry
        # 판매 단가 (입력, 오른쪽 정렬)
        price_var = tk.StringVar()
        price_entry = tk.Entry(self.table_frame, textvariable=price_var, width=self.col_widths[4], justify='right', font=("맑은 고딕", 10))
        price_entry.grid(row=row_num, column=4, sticky='ew', padx=1, pady=1)
        row_widgets['판매 단가'] = price_var
        row_widgets['판매 단가_entry'] = price_entry
        # 판매 합계 (자동 계산, 오른쪽 정렬)
        total_var = tk.StringVar()
        total_entry = tk.Entry(self.table_frame, textvariable=total_var, state='readonly', width=self.col_widths[5], justify='right', font=("맑은 고딕", 10))
        total_entry.grid(row=row_num, column=5, sticky='ew', padx=1, pady=1)
        row_widgets['판매 합계'] = total_var
        row_widgets['판매 합계_entry'] = total_entry
        # 만료일 (입력)
        period_var = tk.StringVar()
        period_value = ''
        if past_row and '만료일' in past_row and past_row['만료일']:
            import datetime
            try:
                expiry = str(past_row['만료일'])
                if isinstance(past_row['만료일'], (datetime.date, datetime.datetime)):
                    expiry_date = past_row['만료일']
                else:
                    expiry_date = datetime.datetime.strptime(expiry, '%Y-%m-%d').date()
                start_date = expiry_date + datetime.timedelta(days=1)
                end_date = start_date.replace(year=start_date.year + 1) - datetime.timedelta(days=1)
                period_value = f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}"
            except Exception:
                period_value = ''
        period_var.set(period_value)
        period_entry = tk.Entry(self.table_frame, textvariable=period_var, width=self.col_widths[6], font=("맑은 고딕", 10))
        period_entry.grid(row=row_num, column=6, sticky='ew', padx=1, pady=1)
        row_widgets['만료일'] = period_var
        row_widgets['만료일_entry'] = period_entry
        # 원가 (수동 입력 가능, 오른쪽 정렬)
        cost_var = tk.StringVar()
        cost_entry = tk.Entry(self.table_frame, textvariable=cost_var, width=self.col_widths[7], justify='right', font=("맑은 고딕", 10))
        cost_entry.grid(row=row_num, column=7, sticky='ew', padx=1, pady=1)
        row_widgets['원가'] = cost_var
        row_widgets['원가_entry'] = cost_entry
        # 원가계 (자동 계산, 오른쪽 정렬)
        cost_total_var = tk.StringVar()
        cost_total_entry = tk.Entry(self.table_frame, textvariable=cost_total_var, state='readonly', width=self.col_widths[8], justify='right', font=("맑은 고딕", 10))
        cost_total_entry.grid(row=row_num, column=8, sticky='ew', padx=1, pady=1)
        row_widgets['원가계'] = cost_total_var
        row_widgets['원가계_entry'] = cost_total_entry
        # 이익액 (자동 계산, 오른쪽 정렬)
        profit_var = tk.StringVar()
        profit_entry = tk.Entry(self.table_frame, textvariable=profit_var, state='readonly', width=self.col_widths[9], justify='right', font=("맑은 고딕", 10))
        profit_entry.grid(row=row_num, column=9, sticky='ew', padx=1, pady=1)
        row_widgets['이익액'] = profit_var
        row_widgets['이익액_entry'] = profit_entry
        # 이익율 (자동 계산, 오른쪽 정렬)
        profit_rate_var = tk.StringVar()
        profit_rate_entry = tk.Entry(self.table_frame, textvariable=profit_rate_var, state='readonly', width=self.col_widths[10], justify='right', font=("맑은 고딕", 10))
        profit_rate_entry.grid(row=row_num, column=10, sticky='ew', padx=1, pady=1)
        row_widgets['이익율'] = profit_rate_var
        row_widgets['이익율_entry'] = profit_rate_entry
        # 매입처 (편집 가능, 오른쪽 정렬)
        supplier_var = tk.StringVar()
        supplier_entry = tk.Entry(self.table_frame, textvariable=supplier_var, width=self.col_widths[11], justify='right', font=("맑은 고딕", 10))
        supplier_entry.grid(row=row_num, column=11, sticky='ew', padx=1, pady=1)
        row_widgets['매입처'] = supplier_var
        row_widgets['매입처_entry'] = supplier_entry
        # 이벤트 바인딩
        qty_var.trace('w', lambda *args, row=row_widgets: self.calculate_row(row))
        price_var.trace('w', lambda *args, row=row_widgets: self.calculate_row(row))
        cost_var.trace('w', lambda *args, row=row_widgets: self.calculate_row(row))
        
        # 과거 견적 내용에서 값 자동 채우기
        if past_row:
            if '대분류' in past_row:
                category_var.set(past_row['대분류'])
            if '소분류' in past_row:
                subcategory_var.set(past_row['소분류'])
            if '제품' in past_row:
                product_var.set(past_row['제품'])
            if '수량' in past_row:
                qty_var.set(past_row['수량'])
            if '판매 단가' in past_row:
                price_var.set(past_row['판매 단가'])
            if '원가' in past_row and past_row['원가']:
                print(f"add_row - past_row에서 원가 설정: '{past_row['원가']}'")
                cost_var.set(past_row['원가'])
                print(f"add_row - cost_var 설정 후: '{cost_var.get()}'")
            if '매입처' in past_row:
                supplier_var.set(past_row['매입처'])
                print(f"add_row - past_row에서 매입처 설정: '{past_row['매입처']}'")

            # Status 정보 로드
            if 'Status' in past_row and past_row['Status']:
                print(f"add_row - past_row에서 Status 설정: '{past_row['Status']}'")
                self.status_var.set(past_row['Status'])
                print(f"add_row - status_var 설정 후: '{self.status_var.get()}'")
            
            # 제품이 설정된 경우 원가 자동 로드
            if '제품' in past_row and past_row['제품']:
                print(f"add_row - 제품이 설정되어 원가 자동 로드 예약: '{past_row['제품']}'")
                # 제품 변경 이벤트를 트리거하여 원가 자동 로드
                self.win.after(100, lambda: self.on_product_change(row_widgets))
                
                # 원가가 이미 있는 경우 즉시 설정
                if '원가' in past_row and past_row['원가']:
                    print(f"add_row - 원가가 이미 있으므로 즉시 설정: '{past_row['원가']}'")
                    cost_var.set(past_row['원가'])

        self.quote_rows.append(row_widgets)
        
        # 행 선택 바인딩 (행이 추가된 후에 인덱스 설정)
        current_row_index = len(self.quote_rows) - 1
        for key, widget in row_widgets.items():
            if hasattr(widget, 'bind'):
                widget.bind('<Button-1>', lambda e, idx=current_row_index: self.select_row(idx))

    def create_customer_table(self):
        """고객사 추가 테이블 생성 (모던 스타일 적용)"""
        # 고객사 테이블 라벨
        tk.Label(self.win, text="영업현황", font=("맑은 고딕", 11, "bold"), bg='#F5F5F5', fg='#374151').pack(anchor='w', padx=10, pady=(15,5))
        
        # 고객사 테이블 프레임 (카드 스타일)
        customer_frame = tk.Frame(self.win, bg='white', relief='solid', bd=1)
        customer_frame.pack(fill='x', padx=10, pady=(0,10))
        
        # 내부 패딩 프레임
        inner_frame = tk.Frame(customer_frame, bg='white')
        inner_frame.pack(fill='x', padx=15, pady=10)
        
        # 영업사원 행
        salesperson_row = tk.Frame(inner_frame, bg='white')
        salesperson_row.pack(fill='x', pady=3)
        tk.Label(salesperson_row, text="영업사원:", font=("맑은 고딕", 10), width=10, anchor='w', bg='white', fg='#374151').pack(side='left')
        
        # 영업사원 표시 (읽기 전용)
        salesperson_label = tk.Label(salesperson_row, text=self.username, font=("맑은 고딕", 10), fg="#374151", bg="#F9FAFB", relief="solid", bd=1, width=30, anchor='w', padx=5)
        salesperson_label.pack(side='left', padx=(5,0))
        
        # Status 행
        status_row = tk.Frame(inner_frame, bg='white')
        status_row.pack(fill='x', pady=3)
        tk.Label(status_row, text="Status:", font=("맑은 고딕", 10), width=10, anchor='w', bg='white', fg='#374151').pack(side='left')
        
        # Status 선택 드롭다운
        self.status_var = tk.StringVar(value="60%")
        status_combo = ttk.Combobox(status_row, textvariable=self.status_var, width=30, font=("맑은 고딕", 10), state="readonly")
        status_combo['values'] = ['100%', '80%', '60%', '40%', '20%', 'Drop']
        status_combo.pack(side='left', padx=(5,0))
        
        # 고객사명 행
        customer_row = tk.Frame(inner_frame, bg='white')
        customer_row.pack(fill='x', pady=3)
        tk.Label(customer_row, text="고객사명:", font=("맑은 고딕", 10), width=10, anchor='w', bg='white', fg='#374151').pack(side='left')
        
        # 고객사명 콤보박스
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(customer_row, textvariable=self.customer_var, 
                                          values=['전체'] + self.all_customers, width=30, state='normal')
        self.customer_combo.pack(side='left', padx=(5,0))
        self.customer_combo.bind('<<ComboboxSelected>>', self.on_customer_change)
        self.customer_combo.bind('<Return>', lambda e: self.on_customer_combo_search())  # 엔터키로 검색
        self.customer_entry = tk.Entry(customer_row, width=30, font=("맑은 고딕", 10), relief='solid', bd=1)
        self.customer_entry.pack(side='left', padx=(5,0))
        self.customer_entry.bind('<Return>', lambda e: self.on_customer_search_and_apply())  # 엔터키로 검색
        self.customer_entry.bind('<KeyRelease>', lambda e: self.on_customer_entry_key_release())  # 실시간 검색
        
        # 거래처명 행
        client_row = tk.Frame(inner_frame, bg='white')
        client_row.pack(fill='x', pady=3)
        tk.Label(client_row, text="거래처명:", font=("맑은 고딕", 10), width=10, anchor='w', bg='white', fg='#374151').pack(side='left')
        
        # 거래처명 콤보박스
        self.client_var = tk.StringVar()
        self.client_combo = ttk.Combobox(client_row, textvariable=self.client_var, 
                                        values=['전체'] + self.all_clients, width=30, state='normal')
        self.client_combo.pack(side='left', padx=(5,0))
        self.client_combo.bind('<<ComboboxSelected>>', self.on_client_change)
        self.client_combo.bind('<Return>', lambda e: self.on_client_combo_search())  # 엔터키로 검색
        self.client_entry = tk.Entry(client_row, width=30, font=("맑은 고딕", 10), relief='solid', bd=1)
        self.client_entry.pack(side='left', padx=(5,0))
        self.client_entry.bind('<Return>', lambda e: self.on_client_search_and_apply())  # 엔터키로 검색
        self.client_entry.bind('<KeyRelease>', lambda e: self.on_client_entry_key_release())  # 실시간 검색
        
        # 예상 발주일 행
        forecast_row = tk.Frame(inner_frame, bg='white')
        forecast_row.pack(fill='x', pady=3)
        tk.Label(forecast_row, text="예상 발주일:", font=("맑은 고딕", 10), width=10, anchor='w', bg='white', fg='#374151').pack(side='left')
        
        # 예상 발주일 입력 필드와 캘린더 버튼
        self.forecast_var = tk.StringVar()
        self.forecast_entry = tk.Entry(forecast_row, textvariable=self.forecast_var, width=30, font=("맑은 고딕", 10), relief='solid', bd=1)
        self.forecast_entry.pack(side='left', padx=(5,0))
        
        # 달력 버튼 (클릭 시 DateEntry 팝업) - 모던 스타일
        calendar_btn = tk.Button(forecast_row, text="📅", command=self.show_calendar, 
                                font=("맑은 고딕", 10), bg="#10B981", fg="white", relief="flat", padx=8, pady=2, cursor="hand2")
        calendar_btn.pack(side='left', padx=(5,0))
        
        # 예상 발주일을 클릭하면 바로 달력이 나오도록 설정
        self.forecast_entry.bind('<Button-1>', self.show_calendar)
        
        # 기존 데이터 설정
        if self.rows_data:
            customer_name = self.rows_data[0].get('고객사명', '')
            client_name = self.rows_data[0].get('거래처명', '')
            status_value = self.rows_data[0].get('Status', '')
            forecast_value = self.rows_data[0].get('예상 발주일', '')  # 예상 발주일 로드 (첫 행 기준)
            
            if customer_name:
                self.customer_var.set(customer_name)
                self.customer_entry.delete(0, tk.END)
                self.customer_entry.insert(0, customer_name)
                self.update_client_list(customer_name)
            
            if client_name:
                self.client_var.set(client_name)
                self.client_entry.delete(0, tk.END)
                self.client_entry.insert(0, client_name)
            
            if status_value:
                self.status_var.set(status_value)

            if forecast_value:
                self.forecast_var.set(forecast_value)

    def show_calendar(self, event=None):
        """달력 팝업 표시"""
        if hasattr(self, 'calendar_window') and self.calendar_window is not None:
            self.calendar_window.destroy()
            
        current_date = self.forecast_var.get().strip()
        from datetime import datetime
        parsed_date = datetime.now().date()
        if current_date:
            try:
                # 다양한 날짜 형식 파싱 시도
                if '-' in current_date:
                    parsed_date = datetime.strptime(current_date, '%Y-%m-%d').date()
                elif '.' in current_date:
                    parsed_date = datetime.strptime(current_date, '%Y.%m.%d').date()
            except:
                pass
        
        # DateEntry 팝업 생성
        self.calendar_window = tk.Toplevel(self.win)
        self.calendar_window.title("예상 발주일 선택")
        self.calendar_window.geometry("400x400")
        self.calendar_window.configure(bg='#F5F5F5')
        self.calendar_window.transient(self.win)
        self.calendar_window.grab_set()
        
        def on_closing():
            if self.calendar_window:
                self.calendar_window.destroy()
            self.calendar_window = None
            
        self.calendar_window.protocol("WM_DELETE_WINDOW", on_closing)
        
        # 캘린더 헤더 스타일
        cal_font = ("맑은 고딕", 12)
        header_bg = '#10B981'  # 청록색
        
        # 달력 위젯
        from tkcalendar import Calendar
        cal = Calendar(self.calendar_window, selectmode='day', 
                      year=parsed_date.year, month=parsed_date.month, day=parsed_date.day,
                      font=cal_font, background=header_bg, foreground='white', 
                      headersbackground=header_bg, headersforeground='white',
                      selectbackground='#059669', selectforeground='white',
                      normalbackground='#F5F5F5', normalforeground='black',
                      weekendbackground='#FED7AA', weekendforeground='black',
                      othermonthwebackground='#E5E7EB', othermonthweforeground='#9CA3AF')
        cal.pack(pady=20, padx=20, fill="both", expand=True)
        
        # 선택 버튼
        def on_select():
            selected = cal.get_date()
            # yyyy-mm-dd 형식으로 변환
            try:
                from datetime import datetime
                # tkcalendar가 반환하는 형식에 따라 처리 (설정에 따라 다를 수 있음)
                # 보통 'm/d/yy' 형식이 기본값이지만, date_pattern 설정이 없으면 기본값 따름
                # 안전하게 파싱 시도
                date_obj = datetime.strptime(selected, '%m/%d/%y')
                formatted = date_obj.strftime('%Y-%m-%d')
            except:
                # 파싱 실패 시 그대로 사용하거나 다른 포맷 시도
                formatted = selected
                
            self.forecast_var.set(formatted)
            on_closing()
            
        btn_frame = tk.Frame(self.calendar_window, bg='#F5F5F5')
        btn_frame.pack(pady=10, fill='x')
        
        select_btn = tk.Button(btn_frame, text="선택", command=on_select,
                              font=("맑은 고딕", 11, "bold"), bg="#10B981", fg="white", 
                              relief="flat", padx=20, pady=5, cursor="hand2")
        select_btn.pack(side='bottom', pady=10)
        
        # 위치 조정
        try:
            x = self.win.winfo_rootx() + (self.win.winfo_width() // 2) - 200
            y = self.win.winfo_rooty() + (self.win.winfo_height() // 2) - 200
            self.calendar_window.geometry(f"+{x}+{y}")
        except:
            pass

    def on_customer_search_and_apply(self):
        """엔터 키로 고객사명 검색 및 적용"""
        entered_customer = self.customer_entry.get().strip()
        if entered_customer:
            # 디버깅용 로그 추가
            print(f"고객사 엔터 검색: '{entered_customer}', 전체 고객사 수: {len(self.all_customers)}")
            
            if not self.all_customers:
                print("경고: 고객사 목록이 비어있습니다!")
                return
                
            # 고객사명으로 검색
            matching_customers = [c for c in self.all_customers if entered_customer.lower() in c.lower()]
            print(f"매칭되는 고객사: {matching_customers}")
            
            if matching_customers:
                # 검색 결과를 콤보박스에 설정
                self.customer_combo['values'] = matching_customers
                # 첫 번째 매칭되는 고객사명을 설정
                self.customer_combo.set(matching_customers[0])
                self.customer_var.set(matching_customers[0])
                # 거래처명 목록 업데이트
                self.update_client_list(matching_customers[0])
                # 드롭다운 리스트 표시
                self.customer_combo.event_generate('<Down>')
                self.customer_combo.event_generate('<Up>')
            else:
                # 매칭되는 고객사가 없으면 입력값 그대로 사용
                self.customer_var.set(entered_customer)
                self.update_client_list(entered_customer)

    def on_client_search_and_apply(self):
        """엔터 키로 거래처명 검색 및 적용"""
        entered_client = self.client_entry.get().strip()
        if entered_client:
            # 디버깅용 로그 추가
            print(f"거래처 엔터 검색: '{entered_client}', 전체 거래처 수: {len(self.all_clients)}")
            
            if not self.all_clients:
                print("경고: 거래처 목록이 비어있습니다!")
                return
                
            # 거래처명으로 검색
            matching_clients = [c for c in self.all_clients if entered_client.lower() in c.lower()]
            print(f"매칭되는 거래처: {matching_clients}")
            
            if matching_clients:
                # 검색 결과를 콤보박스에 설정
                self.client_combo['values'] = matching_clients
                # 첫 번째 매칭되는 거래처명을 설정
                self.client_combo.set(matching_clients[0])
                self.client_var.set(matching_clients[0])
                # 드롭다운 리스트 표시
                self.client_combo.event_generate('<Down>')
                self.client_combo.event_generate('<Up>')
            else:
                # 매칭되는 거래처가 없으면 입력값 그대로 사용
                self.client_var.set(entered_client)

    def on_customer_change(self, event=None):
        selected_customer = self.customer_var.get()
        if selected_customer:
            self.customer_entry.delete(0, tk.END)
            self.customer_entry.insert(0, selected_customer)
            self.update_client_list(selected_customer)

    def on_customer_entry_key_release(self):
        """고객사명 수동 입력 시 실시간 검색"""
        entered_customer = self.customer_entry.get().strip()
        if entered_customer:
            matching_customers = [c for c in self.all_customers if entered_customer.lower() in c.lower()]
            if matching_customers:
                self.customer_combo['values'] = matching_customers
                self.customer_combo.set(matching_customers[0])
                self.customer_var.set(matching_customers[0])
                self.update_client_list(matching_customers[0])
                self.customer_combo.event_generate('<Down>')
            else:
                self.customer_var.set(entered_customer)
                self.update_client_list(entered_customer)

    def on_client_change(self, event=None):
        selected_client = self.client_var.get()
        if selected_client:
            self.client_entry.delete(0, tk.END)
            self.client_entry.insert(0, selected_client)

    def on_client_entry_key_release(self):
        """거래처명 수동 입력 시 실시간 검색"""
        entered_client = self.client_entry.get().strip()
        if entered_client:
            # 디버깅용 로그 추가
            print(f"거래처 검색: '{entered_client}', 전체 거래처 수: {len(self.all_clients)}")
            
            if not self.all_clients:
                print("경고: 거래처 목록이 비어있습니다!")
                return
                
            matching_clients = [c for c in self.all_clients if entered_client.lower() in c.lower()]
            print(f"매칭되는 거래처: {matching_clients}")
            
            if matching_clients:
                self.client_combo['values'] = matching_clients
                self.client_combo.set(matching_clients[0])
                self.client_var.set(matching_clients[0])
                self.client_combo.event_generate('<Down>')
                self.client_combo.event_generate('<Up>')
            else:
                self.client_var.set(entered_client)

    def on_customer_combo_search(self):
        """고객사명 콤보박스에서 엔터키로 검색"""
        entered_customer = self.customer_var.get().strip()
        
        # "전체" 선택 시 전체 리스트 표시
        if entered_customer == "전체" or not entered_customer:
            print("고객사 전체 리스트 표시")
            self.customer_combo['values'] = ['전체'] + self.all_customers
            self.customer_combo.set("전체")
            self.customer_var.set("전체")
            # 거래처명도 전체 리스트로 초기화
            self.client_combo['values'] = ['전체'] + self.all_clients
            self.client_var.set("전체")
            # 드롭다운 리스트 표시
            self.customer_combo.event_generate('<Down>')
            self.customer_combo.event_generate('<Up>')
            return
            
        if entered_customer:
            # 디버깅용 로그 추가
            print(f"고객사 콤보박스 검색: '{entered_customer}', 전체 고객사 수: {len(self.all_customers)}")
            
            if not self.all_customers:
                print("경고: 고객사 목록이 비어있습니다!")
                return
                
            # 고객사명으로 검색
            matching_customers = [c for c in self.all_customers if entered_customer.lower() in c.lower()]
            print(f"매칭되는 고객사: {matching_customers}")
            
            if matching_customers:
                # 검색 결과를 콤보박스에 설정
                self.customer_combo['values'] = ['전체'] + matching_customers
                # 첫 번째 매칭되는 고객사명을 설정
                self.customer_combo.set(matching_customers[0])
                self.customer_var.set(matching_customers[0])
                # 거래처명 목록 업데이트
                self.update_client_list(matching_customers[0])
                # 드롭다운 리스트 표시
                self.customer_combo.event_generate('<Down>')
                self.customer_combo.event_generate('<Up>')
            else:
                # 매칭되는 고객사가 없으면 입력값 그대로 사용
                self.customer_var.set(entered_customer)
                self.update_client_list(entered_customer)

    def on_client_combo_search(self):
        """거래처명 콤보박스에서 엔터키로 검색"""
        entered_client = self.client_var.get().strip()
        
        # "전체" 선택 시 전체 리스트 표시
        if entered_client == "전체" or not entered_client:
            print("거래처 전체 리스트 표시")
            self.client_combo['values'] = ['전체'] + self.all_clients
            self.client_combo.set("전체")
            self.client_var.set("전체")
            # 드롭다운 리스트 표시
            self.client_combo.event_generate('<Down>')
            self.client_combo.event_generate('<Up>')
            return
            
        if entered_client:
            # 디버깅용 로그 추가
            print(f"거래처 콤보박스 검색: '{entered_client}', 전체 거래처 수: {len(self.all_clients)}")
            
            if not self.all_clients:
                print("경고: 거래처 목록이 비어있습니다!")
                return
                
            # 거래처명으로 검색
            matching_clients = [c for c in self.all_clients if entered_client.lower() in c.lower()]
            print(f"매칭되는 거래처: {matching_clients}")
            
            if matching_clients:
                # 검색 결과를 콤보박스에 설정
                self.client_combo['values'] = ['전체'] + matching_clients
                # 첫 번째 매칭되는 거래처명을 설정
                self.client_combo.set(matching_clients[0])
                self.client_var.set(matching_clients[0])
                # 드롭다운 리스트 표시
                self.client_combo.event_generate('<Down>')
                self.client_combo.event_generate('<Up>')
            else:
                # 매칭되는 거래처가 없으면 입력값 그대로 사용
                self.client_var.set(entered_client)

    def update_client_list(self, customer_name):
        if not self.customers_df.empty and customer_name and customer_name != "전체":
            client_list = self.customers_df[self.customers_df['고객사명'] == customer_name]['거래처명'].dropna().unique().tolist()
            self.client_combo['values'] = ['전체'] + client_list
        else:
            # 고객사가 "전체"이거나 비어있으면 전체 거래처 목록 표시
            self.client_combo['values'] = ['전체'] + self.all_clients

    def on_product_search_and_apply(self, row_widgets):
        """엔터 키로 제품 검색 및 드롭다운 표시"""
        if not self.price_df.empty:
            product_input = row_widgets['제품'].get().strip()
            if len(product_input) >= 2:  # 2글자 이상 입력 시 검색 시작
                # 제품명으로 검색 (대분류 무관)
                matching_rows = self.price_df[self.price_df['제품'].str.contains(product_input, case=False, na=False, regex=False)]
                if not matching_rows.empty:
                    # 검색 결과를 콤보박스에 표시
                    product_combo = row_widgets.get('제품_combo')
                    if product_combo:
                        # 검색된 제품명들을 콤보박스 값으로 설정
                        matching_products = matching_rows['제품'].unique().tolist()
                        product_combo['values'] = matching_products
                        
                        # 드롭다운 리스트 표시 (엔터키를 눌렀을 때만)
                        product_combo.event_generate('<Down>')
                        product_combo.event_generate('<Up>')

    def on_product_key_release(self, event, row_widgets):
        """제품 입력 시 검색 기능 (자동완성 제거, 엔터키로만 드롭다운 표시)"""
        # 자동완성 기능 제거 - 엔터키를 눌렀을 때만 드롭다운이 표시되도록 함
        pass

    def on_category_change(self, row_widgets):
        """대분류 변경 시 해당 행의 소분류와 제품 콤보박스 업데이트"""
        category = row_widgets['대분류'].get()
        
        # 소분류 목록 업데이트
        subcategories = self.get_subcategories_by_category(category)
        subcategory_combo = row_widgets.get('소분류_combo')
        if subcategory_combo:
            subcategory_combo['values'] = subcategories
            subcategory_combo.set('')  # 소분류 초기화
        
        # 제품 목록 업데이트 (대분류에 따른 전체 제품)
        products = self.get_products_by_category(category)
        product_combo = row_widgets.get('제품_combo')
        if product_combo:
            product_combo['values'] = products
            # 현재 제품이 새로운 대분류에 있으면 유지, 없으면 초기화
            current_product = row_widgets['제품'].get()
            if current_product in products:
                product_combo.set(current_product)
                # 현재 제품의 원가 다시 설정
                self.on_product_change(row_widgets)
            else:
                product_combo.set('')  # 선택 초기화
                row_widgets['원가'].set('')
                print(f"대분류 '{category}'에서 제품 '{current_product}'을 찾을 수 없어 원가 초기화")
        else:
            # 원가 초기화
            row_widgets['원가'].set('')
        self.calculate_row(row_widgets)

    def on_subcategory_change(self, row_widgets):
        """소분류 변경 시 해당 행의 제품 콤보박스 업데이트"""
        subcategory = row_widgets['소분류'].get()
        category = row_widgets['대분류'].get()
        
        if subcategory:
            # 소분류에 따른 제품 목록 업데이트
            products = self.get_products_by_subcategory(subcategory)
        else:
            # 소분류가 비어있으면 대분류에 따른 전체 제품 목록
            products = self.get_products_by_category(category)
        
        product_combo = row_widgets.get('제품_combo')
        if product_combo:
            product_combo['values'] = products
            # 현재 제품이 새로운 소분류에 있으면 유지, 없으면 초기화
            current_product = row_widgets['제품'].get()
            if current_product in products:
                product_combo.set(current_product)
                # 현재 제품의 원가 다시 설정
                self.on_product_change(row_widgets)
            else:
                product_combo.set('')  # 선택 초기화
                row_widgets['원가'].set('')
        else:
            # 원가 초기화
            row_widgets['원가'].set('')
        self.calculate_row(row_widgets)

    def get_valid_price(self, val):
        print(f"get_valid_price: Input val = '{val}' (type: {type(val)})")
        try:
            if val is None:
                print("get_valid_price: val is None, returning ''")
                return ''
            if isinstance(val, float) and pd.isna(val):
                print("get_valid_price: val is NaN, returning ''")
                return ''
            val_str = str(val).strip()
            print(f"get_valid_price: val_str after strip = '{val_str}'")
            if val_str == '' or val_str.lower() == 'nan':
                print("get_valid_price: val_str is empty or 'nan', returning ''")
                return ''
            # 통화 기호, 공백 등 제거 (숫자, 소수점만 남김)
            val_str = re.sub(r'[^\d.]', '', val_str)
            print(f"get_valid_price: val_str after regex = '{val_str}'")
            if val_str == '':
                print("get_valid_price: val_str after regex is empty, returning ''")
                return ''
            print(f"get_valid_price: Returning '{val_str}'")
            return val_str
        except Exception as e:
            print(f"get_valid_price: Error during processing '{val}' - {e}, returning ''")
            return ''

    def on_product_change(self, row_widgets):
        """제품 변경 시 원가를 기존 선택행 또는 price_list의 '원가' 열에서 가져옴"""
        product = row_widgets['제품'].get()
        category = row_widgets['대분류'].get()
        cost = ''
        
        print(f"견적 수정: 제품 변경 - '{product}', 대분류: '{category}'")
        
        # 1. 먼저 기존 선택행의 원가 확인 (우선순위 1)
        if hasattr(self, 'rows_data') and self.rows_data:
            original_product = self.rows_data[0].get('제품', '')
            original_category = self.rows_data[0].get('대분류', '')
            original_cost = self.rows_data[0].get('원가', '')
            
            # 동일한 제품과 대분류인 경우 기존 원가 사용
            if (product == original_product and category == original_category and 
                original_cost and original_cost.strip()):
                cost = original_cost
                print(f"견적 수정: 기존 선택행에서 원가 가져옴: {cost}")
            else:
                print(f"견적 수정: 기존 선택행과 다른 제품/대분류이므로 가격표에서 검색")
        
        # 2. 기존 원가가 없으면 가격표에서 검색 (우선순위 2)
        if not cost and not self.price_df.empty and product:
            # 2-1. 먼저 현재 대분류에서 제품 검색
            row = self.price_df[(self.price_df['대분류'] == category) & (self.price_df['제품'] == product)]
            
            if not row.empty:
                # 현재 대분류에서 제품을 찾은 경우
                matched_row = row.iloc[0]
                raw_cost = matched_row.get('원가', '')
                cost = self.get_valid_price(raw_cost)
                
                # 소분류 자동 설정
                subcategory = matched_row.get('소분류', '')
                subcategory_combo = row_widgets.get('소분류_combo')
                if subcategory_combo and subcategory:
                    # 소분류 콤보박스의 values 업데이트
                    subcategories = self.get_subcategories_by_category(category)
                    subcategory_combo['values'] = subcategories
                    subcategory_combo.set(subcategory)
                    row_widgets['소분류'].set(subcategory)
                
                # 매입처 자동 설정
                supplier = matched_row.get('매입처', '')
                supplier_entry = row_widgets.get('매입처')
                if supplier_entry and supplier:
                    supplier_entry.set(supplier)
                
                print(f"견적 수정: 현재 대분류에서 원가 찾음 - 원본값: '{raw_cost}', 변환값: '{cost}', 소분류: {subcategory}, 매입처: {supplier}")
            else:
                # 2-2. 전체 제품에서 정확히 일치하는 제품 검색
                exact_match = self.price_df[self.price_df['제품'] == product]
                if not exact_match.empty:
                    matched_row = exact_match.iloc[0]
                    matched_category = matched_row['대분류']
                    raw_cost = matched_row.get('원가', '')
                    cost = self.get_valid_price(raw_cost)
                    
                    # 대분류 자동 설정
                    category_combo = row_widgets.get('대분류_combo')
                    if category_combo:
                        category_combo.set(matched_category)
                        row_widgets['대분류'].set(matched_category)
                    
                    # 소분류 자동 설정
                    subcategory = matched_row.get('소분류', '')
                    subcategory_combo = row_widgets.get('소분류_combo')
                    if subcategory_combo and subcategory:
                        # 소분류 콤보박스의 values 업데이트
                        subcategories = self.get_subcategories_by_category(matched_category)
                        subcategory_combo['values'] = subcategories
                        subcategory_combo.set(subcategory)
                        row_widgets['소분류'].set(subcategory)
                    
                    # 매입처 자동 설정
                    supplier = matched_row.get('매입처', '')
                    supplier_entry = row_widgets.get('매입처')
                    if supplier_entry and supplier:
                        supplier_entry.set(supplier)
                    
                    print(f"견적 수정: 전체 제품에서 원가 찾음 - 원본값: '{raw_cost}', 변환값: '{cost}', 대분류: {matched_category}, 소분류: {subcategory}, 매입처: {supplier}")
                else:
                    # 2-3. 제품명에 포함되는 제품 검색
                    matching_rows = self.price_df[self.price_df['제품'].str.contains(product, case=False, na=False, regex=False)]
                    if not matching_rows.empty:
                        matched_row = matching_rows.iloc[0]
                        matched_category = matched_row['대분류']
                        matched_product = matched_row['제품']
                        raw_cost = matched_row.get('원가', '')
                        cost = self.get_valid_price(raw_cost)
                        
                        # 대분류 자동 설정
                        category_combo = row_widgets.get('대분류_combo')
                        if category_combo:
                            category_combo.set(matched_category)
                            row_widgets['대분류'].set(matched_category)
                        
                        # 소분류 자동 설정
                        subcategory = matched_row.get('소분류', '')
                        subcategory_combo = row_widgets.get('소분류_combo')
                        if subcategory_combo and subcategory:
                            # 소분류 콤보박스의 values 업데이트
                            subcategories = self.get_subcategories_by_category(matched_category)
                            subcategory_combo['values'] = subcategories
                            subcategory_combo.set(subcategory)
                            row_widgets['소분류'].set(subcategory)
                        
                        # 매입처 자동 설정
                        supplier = matched_row.get('매입처', '')
                        supplier_entry = row_widgets.get('매입처')
                        if supplier_entry and supplier:
                            supplier_entry.set(supplier)
                        
                        # 제품명 정확히 설정
                        product_combo = row_widgets.get('제품_combo')
                        if product_combo:
                            product_combo.set(matched_product)
                            row_widgets['제품'].set(matched_product)
                        
                        print(f"견적 수정: 부분 일치로 원가 찾음 - 원본값: '{raw_cost}', 변환값: '{cost}', 제품: {matched_product}, 소분류: {subcategory}, 매입처: {supplier}")
                    else:
                        print(f"견적 수정: 제품 '{product}'에 대한 원가 정보를 찾을 수 없습니다.")
        else:
            if not cost and self.price_df.empty:
                print("견적 수정: 가격 데이터가 비어있습니다.")
            if not cost and not product:
                print("견적 수정: 제품명이 비어있습니다.")
        
        # 원가를 항상 천 단위 콤마로 표시
        try:
            if cost and cost != '':
                # 이미 콤마가 포함된 경우 제거 후 변환
                cost_clean = cost.replace(',', '')
                cost_val = float(cost_clean)
                cost = f"{int(cost_val):,}"
                print(f"견적 수정: 원가 천 단위 콤마 변환: {cost}")
        except Exception as e:
            print(f"견적 수정: 원가 천 단위 콤마 변환 오류: {e}")
            cost = ''
        
        print(f"견적 수정: 최종 원가 설정: '{cost}'")
        row_widgets['원가'].set(str(cost))
        # 원가 설정 후에는 원가계와 이익액만 계산 (전체 calculate_row 대신)
        self.calculate_profit_only(row_widgets)

    def calculate_row(self, row_widgets):
        try:
            qty = float(row_widgets['수량'].get().replace(',', '') or 0)
            price = float(row_widgets['판매 단가'].get().replace(',', '') or 0)
            
            # 원가 값 처리 - 콤마 제거 후 변환
            cost_raw = row_widgets['원가'].get()
            print(f"calculate_row - 원가 원본 값: '{cost_raw}'")
            
            # 원가가 비어있거나 0인 경우에만 기본값 사용
            if not cost_raw or cost_raw.strip() == '' or cost_raw == '0':
                cost = 0
                print(f"calculate_row - 원가가 비어있어 0으로 설정")
            else:
                # 콤마 제거 후 float 변환
                cost = float(cost_raw.replace(',', '') or 0)
                print(f"calculate_row - 원가 변환 후: {cost}")
            
            if price > 0:
                row_widgets['판매 단가'].set(f"{int(price):,}")
            total = qty * price
            row_widgets['판매 합계'].set(f"{int(total):,}" if total > 0 else "")
            cost_total = qty * cost
            row_widgets['원가계'].set(f"{int(cost_total):,}" if cost_total > 0 else "")
            profit = total - cost_total
            row_widgets['이익액'].set(f"{int(profit):,}" if profit != 0 else "")
            if total > 0:
                profit_rate = (profit / total) * 100
                row_widgets['이익율'].set(f"{profit_rate:.2f}%" if profit_rate != 0 else "")
            else:
                row_widgets['이익율'].set("")
        except ValueError as e:
            print(f"calculate_row 오류: {e}")
            pass

    def calculate_profit_only(self, row_widgets):
        """원가가 설정된 후 원가계와 이익액만 계산"""
        try:
            qty = float(row_widgets['수량'].get().replace(',', '') or 0)
            price = float(row_widgets['판매 단가'].get().replace(',', '') or 0)
            
            # 원가 값 처리 - 콤마 제거 후 변환
            cost_raw = row_widgets['원가'].get()
            print(f"calculate_profit_only - 원가 원본 값: '{cost_raw}'")
            
            if not cost_raw or cost_raw.strip() == '' or cost_raw == '0':
                cost = 0
                print(f"calculate_profit_only - 원가가 비어있어 0으로 설정")
            else:
                # 콤마 제거 후 float 변환
                cost = float(cost_raw.replace(',', '') or 0)
                print(f"calculate_profit_only - 원가 변환 후: {cost}")
            
            print(f"calculate_profit_only - 수량: {qty}, 판매단가: {price}, 원가: {cost}")
            
            total = qty * price
            cost_total = qty * cost
            profit = total - cost_total
            
            row_widgets['원가계'].set(f"{int(cost_total):,}" if cost_total > 0 else "")
            row_widgets['이익액'].set(f"{int(profit):,}" if profit != 0 else "")
            
            if total > 0:
                profit_rate = (profit / total) * 100
                row_widgets['이익율'].set(f"{profit_rate:.2f}%" if profit_rate != 0 else "")
            else:
                row_widgets['이익율'].set("")
                
            print(f"calculate_profit_only - 원가계: {cost_total}, 이익액: {profit}")
        except ValueError as e:
            print(f"calculate_profit_only 오류: {e}")
            pass

    def select_row(self, row_index):
        """행 선택"""
        # 이전 선택 해제
        if self.selected_row_index is not None and self.selected_row_index < len(self.quote_rows):
            self.clear_row_selection(self.selected_row_index)
        
        # 새로운 행 선택
        self.selected_row_index = row_index
        if row_index < len(self.quote_rows):
            self.highlight_row_selection(row_index)

    def clear_row_selection(self, row_index):
        """행 선택 해제 (시각적 피드백 제거)"""
        if row_index >= len(self.quote_rows):
            return
            
        row_widgets = self.quote_rows[row_index]
        for key, widget in row_widgets.items():
            if hasattr(widget, 'configure'):
                try:
                    # ttk.Combobox는 기본 스타일로 복원
                    if hasattr(widget, 'configure') and 'style' in widget.configure():
                        widget.configure(style='TCombobox')
                    else:
                        widget.configure(bg='white')
                except:
                    # 배경색 변경이 실패하면 무시
                    pass

    def highlight_row_selection(self, row_index):
        """행 선택 시 시각적 피드백 (배경색 변경)"""
        if row_index >= len(self.quote_rows):
            return
            
        row_widgets = self.quote_rows[row_index]
        
        for key, widget in row_widgets.items():
            if hasattr(widget, 'configure'):
                try:
                    # ttk.Combobox는 style을 통해 배경색 변경
                    if hasattr(widget, 'configure') and 'style' in widget.configure():
                        # 커스텀 스타일 생성 (파란색 배경)
                        style = ttk.Style()
                        style.configure('Selected.TCombobox', fieldbackground='#E0F2FE')
                        widget.configure(style='Selected.TCombobox')
                    else:
                        widget.configure(bg='#E0F2FE')  # 연한 파란색 배경
                except:
                    # 배경색 변경이 실패하면 무시
                    pass

    def move_selected_row_up(self):
        """선택된 행을 위로 이동"""
        if (self.selected_row_index is None or 
            self.selected_row_index <= 0 or 
            self.selected_row_index >= len(self.quote_rows)):
            return
        
        target_row = self.selected_row_index - 1
        self.move_row_to_position(self.selected_row_index, target_row)

    def move_selected_row_down(self):
        """선택된 행을 아래로 이동"""
        if (self.selected_row_index is None or 
            self.selected_row_index < 0 or 
            self.selected_row_index >= len(self.quote_rows) - 1):
            return
        
        target_row = self.selected_row_index + 1
        self.move_row_to_position(self.selected_row_index, target_row)

    def move_row_to_position(self, source_row, target_row):
        """행을 특정 위치로 이동"""
        if (source_row == target_row or 
            source_row < 0 or 
            source_row >= len(self.quote_rows) or
            target_row < 0 or 
            target_row >= len(self.quote_rows)):
            return
        
        # 행 데이터 이동
        row_data = self.quote_rows.pop(source_row)
        self.quote_rows.insert(target_row, row_data)
        
        # 위젯 위치 재배치
        self.rearrange_widgets()
        
        # 선택된 행 인덱스 업데이트
        if self.selected_row_index is not None:
            if self.selected_row_index == source_row:
                self.selected_row_index = target_row
            elif source_row < self.selected_row_index <= target_row:
                self.selected_row_index -= 1
            elif target_row <= self.selected_row_index < source_row:
                self.selected_row_index += 1
            
            # 선택된 행 하이라이트 유지
            if 0 <= self.selected_row_index < len(self.quote_rows):
                self.highlight_row_selection(self.selected_row_index)

    def delete_selected_row(self):
        """선택된 행 삭제"""
        try:
            if self.selected_row_index is None:
                from tkinter import messagebox
                messagebox.showwarning("선택 필요", "삭제할 행을 먼저 선택하세요.")
                return
            
            if len(self.quote_rows) <= 1:
                from tkinter import messagebox
                messagebox.showwarning("삭제 불가", "최소 1행은 유지해야 합니다.")
                return
            
            # 선택된 행 번호
            row_num = self.selected_row_index + 1
            
            # 선택된 행의 위젯들 제거
            widgets_to_destroy = []
            for widget in self.table_frame.winfo_children():
                try:
                    grid_info = widget.grid_info()
                    if grid_info and grid_info['row'] == row_num:
                        widgets_to_destroy.append(widget)
                except:
                    continue
            
            for widget in widgets_to_destroy:
                try:
                    widget.destroy()
                except:
                    continue
            
            # quote_rows에서 제거
            self.quote_rows.pop(self.selected_row_index)
            
            # 나머지 행들의 인덱스 업데이트
            for i, row in enumerate(self.quote_rows):
                row['row_index'] = i
            
            # 위젯 재배치
            self.rearrange_widgets()
            
            # 선택 상태 초기화
            self.selected_row_index = None
            print(f"견적 작성(수정): 행 {row_num} 삭제 완료")
            
        except Exception as e:
            print(f"견적 작성(수정): 행 삭제 오류 - {e}")
            from tkinter import messagebox
            messagebox.showerror("오류", f"행 삭제 중 오류가 발생했습니다:\n{e}")

    def rearrange_widgets(self):
        """위젯들을 새로운 행 순서에 맞게 재배치"""
        try:
            # 실제 컬럼 순서 (get_widget_by_col과 일치)
            col_keys = ['대분류', '소분류', '제품', '수량', '판매 단가', '판매 합계', '만료일', '원가', '원가계', '이익액', '이익율', '매입처', '예상 발주일']
            
            for row_idx, row_widgets in enumerate(self.quote_rows):
                row_num = row_idx + 1
                
                # 모든 컬럼들 재배치
                for col_idx, col_key in enumerate(col_keys):
                    try:
                        widget = self.get_widget_by_col(row_widgets, col_idx)
                        if widget and hasattr(widget, 'grid'):  # 위젯이 존재하고 grid 메서드가 있는 경우만
                            widget.grid(row=row_num, column=col_idx, sticky='ew', padx=1, pady=1)
                    except Exception as e:
                        print(f"견적 작성(수정): 컬럼 {col_key} 재배치 오류 - {e}")
                        continue
                
                # 행 선택 이벤트 바인딩 재설정
                for key, widget in row_widgets.items():
                    try:
                        if hasattr(widget, 'bind') and key.endswith(('_combo', '_entry')):
                            # 기존 바인딩 제거
                            widget.unbind('<Button-1>')
                            # 새로운 바인딩 추가
                            widget.bind('<Button-1>', lambda e, idx=row_idx: self.select_row(idx))
                    except Exception as e:
                        print(f"견적 작성(수정): 이벤트 바인딩 오류 - {e}")
                        continue
            
            print("견적 작성(수정): 위젯 재배치 완료")
            
        except Exception as e:
            print(f"견적 작성(수정): 위젯 재배치 오류 - {e}")

    def remove_row(self):
        """마지막 행 삭제 (기존 기능 유지)"""
        if len(self.quote_rows) > 1:
            last_row = len(self.quote_rows)
            for widget in self.table_frame.winfo_children():
                if widget.grid_info()['row'] == last_row:
                    widget.destroy()
            self.quote_rows.pop()
            
            # 선택된 행이 삭제된 행이었다면 선택 상태 초기화
            if self.selected_row_index == len(self.quote_rows):
                self.selected_row_index = None

    def on_save(self):
        import gspread
        from google.oauth2.service_account import Credentials
        from datetime import datetime
        import openpyxl
        import os
        import shutil
        import subprocess
        from tkinter import messagebox

        # 견적서 ID 생성
        def generate_quote_id():
            """견적서 고유 ID 생성 (날짜_순번 형태)"""
            today = datetime.today()
            date_str = today.strftime('%Y-%m-%d')
            
            # 견적서 번호 생성 (오늘 날짜의 전체 견적서 순번)
            quote_number = self._get_next_quote_number(date_str)
            
            quote_id = f"{date_str}_{quote_number:03d}"
            return quote_id, quote_number
        
        def generate_quote_filename(customer_name, client_name, main_category, today):
            """견적서 파일명 생성 (기존 형식 유지)"""
            if customer_name == client_name or not client_name:
                filename = f"{customer_name}_{main_category}_{today}.xlsx"
            else:
                filename = f"{customer_name}_{client_name}_{main_category}_{today}.xlsx"
            return filename

        # 입력값 검증
        # Status 검증
        selected_status = self.status_var.get()
        if selected_status == "Drop":
            messagebox.showwarning("입력 오류", "영업 활동 Status를 기입해 주세요.")
            return

        # 견적 행 데이터 검증
        if not self.quote_rows:
            messagebox.showwarning("입력 오류", "최소 하나의 제품을 추가해주세요.")
            return

        # 각 견적 행 검증
        for i, row_widgets in enumerate(self.quote_rows):
            unit_price = row_widgets['판매 단가'].get().strip()
            cost = row_widgets['원가'].get().strip()
            
            if not unit_price or float(unit_price.replace(',', '') or 0) <= 0:
                messagebox.showwarning("입력 오류", f"{i+1}번째 행의 판매 단가를 입력해주세요.")
                return
            
            if not cost or float(cost.replace(',', '') or 0) <= 0:
                messagebox.showwarning("입력 오류", f"{i+1}번째 행의 원가를 입력해주세요.")
                return

        # 견적서 ID 생성
        customer_name = self.customer_entry.get().strip()
        client_name = self.client_entry.get().strip()
        if not client_name:
            client_name = customer_name
            
        quote_id, quote_number = generate_quote_id()
        print(f"견적서 ID 생성: {quote_id}")

        def write_to_cell_safely(ws, cell_coord, value):
            """병합된 셀을 안전하게 처리하여 값을 쓰는 함수 (병합 해제 없이)"""
            try:
                # 셀 좌표를 행과 열로 변환
                from openpyxl.utils.cell import coordinate_from_string
                from openpyxl.utils import column_index_from_string
                col_str, row_str = coordinate_from_string(cell_coord)
                row = int(row_str)
                col = column_index_from_string(col_str)
                
                # 병합 범위 확인
                merged = False
                for merged_range in ws.merged_cells.ranges:
                    # openpyxl의 MergedCellRange.__contains__는 'A1' 형태의 좌표 문자열을 기대함
                    if cell_coord in merged_range:
                        # 병합 범위의 최상단-왼쪽 셀에 값 쓰기
                        tl_row, tl_col = merged_range.min_row, merged_range.min_col
                        ws.cell(row=tl_row, column=tl_col, value=value)
                        merged = True
                        break
                
                if not merged:
                    # 병합되지 않은 셀에 직접 값 쓰기
                    ws.cell(row=row, column=col, value=value)
                    
            except Exception as e:
                print(f"셀 {cell_coord}에 값 쓰기 실패: {e}")
                # 병합 해제 없이 다른 방법 시도
                try:
                    # 병합된 셀의 경우 최상단-왼쪽 셀에 직접 접근
                    for merged_range in ws.merged_cells.ranges:
                        if cell_coord in merged_range:
                            tl_row, tl_col = merged_range.min_row, merged_range.min_col
                            ws.cell(row=tl_row, column=tl_col, value=value)
                            return
                    
                    # 병합되지 않은 경우 기본 방법
                    ws[cell_coord] = value
                except Exception as e2:
                    print(f"병합 셀 처리 실패: {e2}")
                    # 마지막 시도: 셀 참조 방식
                    try:
                        cell = ws.cell(row=row, column=col)
                        cell.value = value
                    except Exception as e3:
                        print(f"모든 방법 실패: {e3}")

        try:
            today = datetime.today().strftime('%Y%m%d')
            main_category = self.quote_rows[0]['대분류'].get() if self.quote_rows else ''
            
            # 견적서 파일명 생성 (기존 형식 유지)
            filename = generate_quote_filename(customer_name, client_name, main_category, today)
            out_dir = os.path.join(os.getcwd(), 'quote')
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, filename)
            template_path = resource_path('quote/quote_origin_01.xlsx')
            if not os.path.exists(template_path):
                messagebox.showerror("오류", f"견적서 템플릿 파일을 찾을 수 없습니다.\n경로: {template_path}")
                return
            shutil.copy(template_path, out_path)
            wb = openpyxl.load_workbook(out_path)
            ws = wb.active
            
            # C4 셀에 고객사 정보 입력 (병합된 셀 안전하게 처리)
            target_value = customer_name if customer_name == client_name or not client_name else f"{customer_name} - {client_name}"
            write_to_cell_safely(ws, 'C4', target_value)
            
            # 견적 데이터 입력 (병합된 셀 안전하게 처리)
            start_row = 14
            for i, row in enumerate(self.quote_rows):
                d_row = start_row + i
                write_to_cell_safely(ws, f'D{d_row}', clean_product_name(row['제품'].get()))
                
                def parse_number(val):
                    try:
                        # 튜플인 경우 첫 번째 요소 사용
                        if isinstance(val, tuple):
                            val = val[0] if val else ''
                        # 문자열로 변환
                        val_str = str(val).replace(',', '')
                        # 빈 문자열이거나 None인 경우 0 반환
                        if not val_str or val_str.strip() == '':
                            return 0
                        # 숫자 변환 시도
                        if '.' in val_str:
                            return float(val_str)
                        else:
                            return int(val_str)
                    except (ValueError, TypeError):
                        return 0
                
                write_to_cell_safely(ws, f'L{d_row}', parse_number(row['수량'].get()))
                write_to_cell_safely(ws, f'M{d_row}', parse_number(row['판매 단가'].get()))
                write_to_cell_safely(ws, f'T{d_row}', parse_number(row['원가'].get()))
                period = row['만료일'].get()
                if period:
                    write_to_cell_safely(ws, f'Q{d_row}', period)
            
            # 워크시트 이름에서 작은따옴표 제거
            for sheet in wb.sheetnames:
                if sheet.startswith("'"):
                    # 작은따옴표로 시작하는 워크시트 이름 수정
                    new_name = sheet.lstrip("'")
                    wb[sheet].title = new_name
            
            wb.save(out_path)
            try:
                os.startfile(out_path)
            except AttributeError:
                subprocess.Popen(['open', out_path])
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror('엑셀 저장 오류', str(e))

        # Google Sheets 업데이트 (견적서 ID 포함)
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
            ws = sh.worksheet('Sheet1')
            
            # 전체 데이터 가져오기
            all_data = ws.get_all_values()
            sheet_header = all_data[0] if all_data else []
            
            # 컬럼 인덱스 찾기 (ID는 첫 번째 열)
            col_idx = {k: None for k in ['ID','견적서ID','일자','영업사원','Status','고객사명','거래처명','대분류','소분류','제품','수량','판매 단가','판매 합계','원가','원가계','이익액','이익율','만료일','구분','USD','EUR','DC','매입처','예상 발주일']}
            for idx, h in enumerate(sheet_header):
                if h in col_idx:
                    col_idx[h] = idx
            
            # ID 컬럼이 첫 번째 열에 있는지 확인
            if col_idx['ID'] is None and col_idx['견적서ID'] is None:
                print("ID 또는 견적서ID 컬럼이 첫 번째 열에 없습니다. 확인해주세요.")
                messagebox.showerror("오류", "ID 또는 견적서ID 컬럼이 첫 번째 열에 없습니다. Google Sheet를 확인해주세요.")
                return
            
            # 기존 견적서 ID가 있는지 확인
            existing_quote_id = None
            if self.rows_data and len(self.rows_data) > 0:
                existing_quote_id = self.rows_data[0].get('ID', '') or self.rows_data[0].get('견적서ID', '')
                print(f"기존 견적서 ID 확인: {existing_quote_id}")
            
            if existing_quote_id:
                # 기존 견적서 수정
                print(f"기존 견적서 수정: {existing_quote_id}")
                self._update_existing_quote(ws, col_idx, existing_quote_id, customer_name, client_name)
                messagebox.showinfo("성공", "견적서 데이터가 성공적으로 수정되었습니다.")
            else:
                # 새로운 견적서 생성
                print(f"새로운 견적서 생성: {quote_id}")
                self._create_new_quote(ws, col_idx, quote_id, customer_name, client_name)
                messagebox.showinfo("성공", "새로운 견적서가 성공적으로 생성되었습니다.")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            from tkinter import messagebox
            messagebox.showerror('구글시트 저장 오류', str(e))
        
        # 창 닫기
        self.win.destroy()

    def _get_next_quote_number(self, date_str):
        """오늘 날짜의 전체 견적서 번호 생성"""
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
            ws = sh.worksheet('Sheet1')
            
            # 전체 데이터 가져오기
            all_data = ws.get_all_values()
            if len(all_data) <= 1:  # 헤더만 있거나 비어있음
                return 1
            
            # 컬럼 인덱스 찾기 (ID 또는 견적서ID 컬럼 찾기)
            header = all_data[0]
            id_idx = None
            date_idx = None
            
            for idx, h in enumerate(header):
                if h == 'ID':
                    id_idx = idx
                elif h == '견적서ID':
                    id_idx = idx  # 견적서ID도 ID로 사용
                elif h == '일자':
                    date_idx = idx
            
            if id_idx is None or date_idx is None:
                return 1
            
            # 오늘 날짜의 모든 견적서 번호 찾기
            max_number = 0
            found_ids = set()  # 중복 제거를 위한 set
            
            for row in all_data[1:]:  # 헤더 제외
                if len(row) > max(id_idx, date_idx):
                    row_date = row[date_idx]
                    row_id = row[id_idx]
                    
                    if (row_date.startswith(date_str) and 
                        row_id and 
                        '_' in row_id and
                        row_id not in found_ids):  # 중복 제거
                        try:
                            # ID에서 번호 추출 (예: 2025-08-20_001 -> 1)
                            number_part = row_id.split('_')[-1]
                            if number_part.isdigit():
                                max_number = max(max_number, int(number_part))
                                found_ids.add(row_id)  # 중복 제거를 위해 추가
                                print(f"견적서ID 발견: {row_id} -> 번호: {number_part}")
                        except:
                            continue
            
            next_number = max_number + 1
            print(f"오늘 날짜({date_str}) 최대 번호: {max_number}, 다음 번호: {next_number}")
            return next_number
            
        except Exception as e:
            print(f"견적서 번호 생성 오류: {e}")
            return 1

    def _create_new_quote(self, ws, col_idx, quote_id, customer_name, client_name):
        """새로운 견적서 생성 (모든 행에 같은 견적서 ID 사용)"""
        from datetime import datetime
        today = datetime.today().strftime('%Y-%m-%d')
        user = self.username
        
        for i, row in enumerate(self.quote_rows):
            if not (row['대분류'].get() and row['제품'].get() and row['수량'].get() and row['판매 단가'].get() and row['원가'].get()):
                continue
            
            # 새로운 행 데이터 생성
            data = [''] * len(ws.row_values(1))
            
            # 견적서 ID (첫 번째 열, 모든 행에 같은 ID)
            if col_idx.get('견적서ID') is not None:
                data[col_idx['견적서ID']] = quote_id
                print(f"견적서 ID 저장: {quote_id} -> 행 {i+1}")
            elif col_idx.get('ID') is not None:
                data[col_idx['ID']] = quote_id
                print(f"ID 저장: {quote_id} -> 행 {i+1}")
            
            # 기본 정보
            if col_idx['일자'] is not None:
                data[col_idx['일자']] = today
            if col_idx['영업사원'] is not None:
                data[col_idx['영업사원']] = user
            if col_idx['대분류'] is not None:
                data[col_idx['대분류']] = row['대분류'].get()
            if col_idx.get('소분류') is not None:
                data[col_idx['소분류']] = row['소분류'].get()
                print(f"소분류 저장: {row['소분류'].get()} -> 인덱스 {col_idx['소분류']}")
            if col_idx['제품'] is not None:
                data[col_idx['제품']] = clean_product_name(row['제품'].get())
            if col_idx['수량'] is not None:
                data[col_idx['수량']] = row['수량'].get()
            if col_idx['판매 단가'] is not None:
                data[col_idx['판매 단가']] = row['판매 단가'].get()
            if col_idx['원가'] is not None:
                data[col_idx['원가']] = row['원가'].get()
            if col_idx['고객사명'] is not None:
                data[col_idx['고객사명']] = customer_name
            if col_idx['거래처명'] is not None:
                data[col_idx['거래처명']] = client_name
            if col_idx.get('만료일') is not None:
                data[col_idx['만료일']] = row['만료일'].get()
            if col_idx.get('예상 발주일') is not None:
                data[col_idx['예상 발주일']] = self.forecast_var.get()
            
            # price_list에서 제품 정보 가져와서 저장
            product_name = clean_product_name(row['제품'].get())
            if product_name:
                product_info = self.get_product_info_from_price_list(product_name)
                if product_info:
                    print(f"제품 정보 저장 시작: {product_name}")
                    print(f"찾은 제품 정보: {product_info}")
                    print(f"컬럼 인덱스: {col_idx}")
                    
                    # 안전하게 컬럼 존재 여부 확인 후 저장
                    if '구분' in col_idx and col_idx['구분'] is not None:
                        data[col_idx['구분']] = product_info.get('구분', '')
                        print(f"구분 저장: {product_info.get('구분', '')} -> 인덱스 {col_idx['구분']}")
                    if 'USD' in col_idx and col_idx['USD'] is not None:
                        data[col_idx['USD']] = product_info.get('USD', '')
                        print(f"USD 저장: {product_info.get('USD', '')} -> 인덱스 {col_idx['USD']}")
                    if 'EUR' in col_idx and col_idx['EUR'] is not None:
                        data[col_idx['EUR']] = product_info.get('EUR', '')
                        print(f"EUR 저장: {product_info.get('EUR', '')} -> 인덱스 {col_idx['EUR']}")
                    if 'DC' in col_idx and col_idx['DC'] is not None:
                        data[col_idx['DC']] = product_info.get('DC', '')
                        print(f"DC 저장: {product_info.get('DC', '')} -> 인덱스 {col_idx['DC']}")
                    if '매입처' in col_idx and col_idx['매입처'] is not None:
                        # UI에서 입력된 매입처가 있으면 우선 사용, 없으면 price_list에서 가져온 값 사용
                        ui_supplier = row['매입처'].get()
                        if ui_supplier and ui_supplier.strip():
                            data[col_idx['매입처']] = ui_supplier
                            print(f"매입처 저장 (UI): {ui_supplier} -> 인덱스 {col_idx['매입처']}")
                        else:
                            data[col_idx['매입처']] = product_info.get('매입처', '')
                            print(f"매입처 저장 (price_list): {product_info.get('매입처', '')} -> 인덱스 {col_idx['매입처']}")
                    
                    print(f"최종 저장될 데이터: {data}")
                else:
                    print(f"제품 정보를 찾지 못함: {product_name}")
            else:
                print(f"제품명이 비어있음")
            
            # Status 컬럼 추가
            if 'Status' in col_idx and col_idx['Status'] is not None:
                selected_status = self.status_var.get()
                print(f"Status 저장: {selected_status} -> 인덱스 {col_idx['Status']}")
                # Status가 비어있으면 기본값 60% 사용, Drop이면 Drop 그대로 사용
                if not selected_status:
                    selected_status = "60%"
                    print(f"Status가 비어있으므로 기본값 60% 사용")
                data[col_idx['Status']] = selected_status
                print(f"Status 최종 저장: {selected_status}")
            else:
                print(f"Status 컬럼을 찾을 수 없음")
            
            # 수식 계산 (FCST 열 추가로 인해 한 칸씩 밀림)
            last_row = len(ws.get_all_values()) + 1
            if col_idx['판매 합계'] is not None:
                data[col_idx['판매 합계']] = f'=K{last_row}*L{last_row}'  # 수량(K) × 판매 단가(L)
            if col_idx['원가계'] is not None:
                data[col_idx['원가계']] = f'=N{last_row}*K{last_row}'  # 원가(N) × 수량(K)
            if col_idx['이익액'] is not None:
                data[col_idx['이익액']] = f'=M{last_row}-O{last_row}'  # 판매 합계(M) - 원가계(O)
            if col_idx['이익율'] is not None:
                data[col_idx['이익율']] = f'=IF(M{last_row}=0,0,P{last_row}/M{last_row})'  # 이익액(P) / 판매 합계(M)
            
            ws.append_row(data, value_input_option='USER_ENTERED')
            
            # 숫자 컬럼들 오른쪽 정렬 설정 (FCST 열 추가로 인해 한 칸씩 밀림)
            last_row = len(ws.get_all_values())
            if col_idx['수량'] is not None:
                ws.format(f'K{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['판매 단가'] is not None:
                ws.format(f'L{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['판매 합계'] is not None:
                ws.format(f'M{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['원가'] is not None:
                ws.format(f'N{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['원가계'] is not None:
                ws.format(f'O{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['이익액'] is not None:
                ws.format(f'P{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['이익율'] is not None:
                ws.format(f'Q{last_row}', {'horizontalAlignment': 'RIGHT'})

    def _update_existing_quote(self, ws, col_idx, existing_quote_id, customer_name, client_name):
        """기존 견적서 수정 (견적서 ID는 유지)"""
        from datetime import datetime
        today = datetime.today().strftime('%Y-%m-%d')
        user = self.username
        
        # 기존 견적서 ID로 모든 행 찾기
        all_values = ws.get_all_values()
        existing_sheet_rows = []  # 기존 시트의 행들 (행 번호와 데이터)
        updated_rows = 0
        
        for row_idx, row_data in enumerate(all_values[1:], start=2):  # 헤더 제외
            # None이 아닌 값들만 필터링하여 최대값 계산
            valid_indices = [idx for idx in col_idx.values() if idx is not None]
            if not valid_indices or len(row_data) <= max(valid_indices):
                continue
            
            # 견적서 ID가 일치하는 행 찾기 (ID 또는 견적서ID 컬럼 확인)
            sheet_quote_id = ''
            if col_idx.get('ID') is not None and col_idx['ID'] < len(row_data):
                sheet_quote_id = row_data[col_idx['ID']]
            elif col_idx.get('견적서ID') is not None and col_idx['견적서ID'] < len(row_data):
                sheet_quote_id = row_data[col_idx['견적서ID']]
            
            if sheet_quote_id == existing_quote_id:
                existing_sheet_rows.append((row_idx, row_data))
        
        print(f"기존 견적서에서 찾은 행 개수: {len(existing_sheet_rows)}")
        
        # 삭제된 행 처리: 기존 행 수가 현재 UI 행 수보다 많으면 삭제
        if len(existing_sheet_rows) > len(self.quote_rows):
            rows_to_delete = len(existing_sheet_rows) - len(self.quote_rows)
            print(f"삭제할 행 개수: {rows_to_delete}")
            
            # 뒤에서부터 삭제 (행 번호가 뒤로 갈수록 큰 번호이므로)
            for i in range(rows_to_delete):
                row_to_delete = existing_sheet_rows[-(i+1)]  # 뒤에서부터
                row_idx = row_to_delete[0]
                print(f"행 {row_idx} 삭제")
                ws.delete_rows(row_idx)
            
            # 삭제 후 기존 행들을 다시 찾기
            all_values = ws.get_all_values()
            existing_sheet_rows = []
            for row_idx, row_data in enumerate(all_values[1:], start=2):
                valid_indices = [idx for idx in col_idx.values() if idx is not None]
                if not valid_indices or len(row_data) <= max(valid_indices):
                    continue
                
                sheet_quote_id = ''
                if col_idx.get('ID') is not None and col_idx['ID'] < len(row_data):
                    sheet_quote_id = row_data[col_idx['ID']]
                elif col_idx.get('견적서ID') is not None and col_idx['견적서ID'] < len(row_data):
                    sheet_quote_id = row_data[col_idx['견적서ID']]
                
                if sheet_quote_id == existing_quote_id:
                    existing_sheet_rows.append((row_idx, row_data))
        
        # 기존 행들 업데이트
        for i, (row_idx, row_data) in enumerate(existing_sheet_rows):
            if i < len(self.quote_rows):  # UI에 해당 행이 있는 경우만 업데이트
                print(f"기존 견적서 행 업데이트: 행 {row_idx}, 견적서ID: {existing_quote_id}")
                self._update_changed_values_only(ws, col_idx, row_idx, customer_name, client_name)
                updated_rows += 1
        
        print(f"업데이트된 행 개수: {updated_rows}")
        
        # 새로운 제품이 있으면 기존 견적서에 추가
        if len(self.quote_rows) > updated_rows:
            print(f"새로운 제품 {len(self.quote_rows) - updated_rows}개를 기존 견적서에 추가합니다.")
            self._add_new_products_to_existing_quote(ws, col_idx, existing_quote_id, customer_name, client_name, updated_rows)

    def _add_new_products_to_existing_quote(self, ws, col_idx, quote_id, customer_name, client_name, existing_product_count):
        """기존 견적서에 새로운 제품 추가 (기존 견적서 ID 사용)"""
        from datetime import datetime
        today = datetime.today().strftime('%Y-%m-%d')
        user = self.username
        
        # 기존 제품 수만큼 건너뛰고 새로운 제품들만 추가
        for i, row in enumerate(self.quote_rows[existing_product_count:], start=existing_product_count):
            if not (row['대분류'].get() and row['제품'].get() and row['수량'].get() and row['판매 단가'].get() and row['원가'].get()):
                continue
            
            # 새로운 행 데이터 생성 (기존 견적서와 동일한 ID 사용)
            data = [''] * len(ws.row_values(1))
            
            # 견적서 ID (기존 견적서와 동일)
            if col_idx.get('견적서ID') is not None:
                data[col_idx['견적서ID']] = quote_id
                print(f"새 제품 추가 - 견적서 ID: {quote_id} -> 행 {i+1}")
            elif col_idx.get('ID') is not None:
                data[col_idx['ID']] = quote_id
                print(f"새 제품 추가 - ID: {quote_id} -> 행 {i+1}")
            
            # 기본 정보
            if col_idx['일자'] is not None:
                data[col_idx['일자']] = today
            if col_idx['영업사원'] is not None:
                data[col_idx['영업사원']] = user
            if col_idx['대분류'] is not None:
                data[col_idx['대분류']] = row['대분류'].get()
            if col_idx.get('소분류') is not None:
                data[col_idx['소분류']] = row['소분류'].get()
            if col_idx['제품'] is not None:
                data[col_idx['제품']] = clean_product_name(row['제품'].get())
            if col_idx['수량'] is not None:
                data[col_idx['수량']] = row['수량'].get()
            if col_idx['판매 단가'] is not None:
                data[col_idx['판매 단가']] = row['판매 단가'].get()
            if col_idx['원가'] is not None:
                data[col_idx['원가']] = row['원가'].get()
            if col_idx['고객사명'] is not None:
                data[col_idx['고객사명']] = customer_name
            if col_idx['거래처명'] is not None:
                data[col_idx['거래처명']] = client_name
            if col_idx.get('만료일') is not None:
                data[col_idx['만료일']] = row['만료일'].get()
            
            # 제품 정보 및 Status 저장
            product_name = clean_product_name(row['제품'].get())
            if product_name:
                product_info = self.get_product_info_from_price_list(product_name)
                if product_info:
                    if '구분' in col_idx and col_idx['구분'] is not None:
                        data[col_idx['구분']] = product_info.get('구분', '')
                    if 'USD' in col_idx and col_idx['USD'] is not None:
                        data[col_idx['USD']] = product_info.get('USD', '')
                    if 'EUR' in col_idx and col_idx['EUR'] is not None:
                        data[col_idx['EUR']] = product_info.get('EUR', '')
                    if 'DC' in col_idx and col_idx['DC'] is not None:
                        data[col_idx['DC']] = product_info.get('DC', '')
                    
                    # 매입처: UI 입력값 우선, 없으면 price_list 사용
                    if '매입처' in col_idx and col_idx['매입처'] is not None:
                        ui_supplier = row['매입처'].get().strip()
                        if ui_supplier:
                            data[col_idx['매입처']] = ui_supplier
                            print(f"새 제품 매입처 저장 (UI): {ui_supplier}")
                        else:
                            data[col_idx['매입처']] = product_info.get('매입처', '')
                            print(f"새 제품 매입처 저장 (price_list): {product_info.get('매입처', '')}")
                else:
                    # 제품 정보가 없어도 UI 입력값은 저장
                    if '매입처' in col_idx and col_idx['매입처'] is not None:
                        ui_supplier = row['매입처'].get().strip()
                        if ui_supplier:
                            data[col_idx['매입처']] = ui_supplier
                            print(f"새 제품 매입처 저장 (UI): {ui_supplier}")
            else:
                # 제품명이 없어도 UI 입력값은 저장
                if '매입처' in col_idx and col_idx['매입처'] is not None:
                    ui_supplier = row['매입처'].get().strip()
                    if ui_supplier:
                        data[col_idx['매입처']] = ui_supplier
                        print(f"새 제품 매입처 저장 (UI): {ui_supplier}")
            
            # Status 저장
            if 'Status' in col_idx and col_idx['Status'] is not None:
                selected_status = self.status_var.get()
                if not selected_status:
                    selected_status = "60%"
                data[col_idx['Status']] = selected_status
            
            # 수식 입력
            last_row = len(ws.get_all_values())
            if col_idx['판매 합계'] is not None:
                data[col_idx['판매 합계']] = f'=K{last_row}*L{last_row}'
            if col_idx['원가계'] is not None:
                data[col_idx['원가계']] = f'=N{last_row}*K{last_row}'
            if col_idx['이익액'] is not None:
                data[col_idx['이익액']] = f'=M{last_row}-O{last_row}'
            if col_idx['이익율'] is not None:
                data[col_idx['이익율']] = f'=IF(M{last_row}=0,0,P{last_row}/M{last_row})'
            
            ws.append_row(data, value_input_option='USER_ENTERED')
            
            # 숫자 컬럼들 오른쪽 정렬 설정
            if col_idx['수량'] is not None:
                ws.format(f'J{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['판매 단가'] is not None:
                ws.format(f'K{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['판매 합계'] is not None:
                ws.format(f'L{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['원가'] is not None:
                ws.format(f'M{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['원가계'] is not None:
                ws.format(f'N{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['이익액'] is not None:
                ws.format(f'O{last_row}', {'horizontalAlignment': 'RIGHT'})
            if col_idx['이익율'] is not None:
                ws.format(f'P{last_row}', {'horizontalAlignment': 'RIGHT'})

    def _update_changed_values_only(self, ws, col_idx, row_idx, customer_name, client_name):
        """기존 행에서 변경된 값만 업데이트"""
        try:
            # 현재 시트의 행 데이터 가져오기
            current_row = ws.row_values(row_idx)
            print(f"현재 행 {row_idx} 데이터: {current_row}")
            
            # 변경된 값들만 업데이트
            updated_fields = []
            
            # 고객사명이 변경된 경우
            if col_idx.get('고객사명') is not None:
                current_customer = current_row[col_idx['고객사명']] if col_idx['고객사명'] < len(current_row) else ''
                if current_customer != customer_name:
                    ws.update_cell(row_idx, col_idx['고객사명'] + 1, customer_name)  # 구글 시트는 1부터 시작
                    updated_fields.append(f"고객사명: {current_customer} → {customer_name}")
            
            # 거래처명이 변경된 경우
            if col_idx.get('거래처명') is not None:
                current_client = current_row[col_idx['거래처명']] if col_idx['거래처명'] < len(current_row) else ''
                if current_client != client_name:
                    ws.update_cell(row_idx, col_idx['거래처명'] + 1, client_name)
                    updated_fields.append(f"거래처명: {current_client} → {client_name}")
            
            # Status가 변경된 경우
            if col_idx.get('Status') is not None:
                current_status = current_row[col_idx['Status']] if col_idx['Status'] < len(current_row) else ''
                new_status = self.status_var.get()
                if current_status != new_status:
                    ws.update_cell(row_idx, col_idx['Status'] + 1, new_status)
                    updated_fields.append(f"Status: {current_status} → {new_status}")
            
            # 각 견적 행의 변경사항 확인
            for i, row in enumerate(self.quote_rows):
                if not (row['대분류'].get() and row['제품'].get() and row['수량'].get() and row['판매 단가'].get() and row['원가'].get()):
                    continue
                
                # 현재 행의 데이터와 비교할 원본 데이터 찾기
                original_row = None
                if i < len(self.rows_data):
                    original_row = self.rows_data[i]
                
                if original_row:
                    # 대분류가 변경된 경우
                    if col_idx.get('대분류') is not None:
                        current_category = current_row[col_idx['대분류']] if col_idx['대분류'] < len(current_row) else ''
                        new_category = row['대분류'].get()
                        if current_category != new_category:
                            ws.update_cell(row_idx, col_idx['대분류'] + 1, new_category)
                            updated_fields.append(f"대분류: {current_category} → {new_category}")
                    
                    # 소분류가 변경된 경우
                    if col_idx.get('소분류') is not None:
                        current_subcategory = current_row[col_idx['소분류']] if col_idx['소분류'] < len(current_row) else ''
                        new_subcategory = row['소분류'].get()
                        if current_subcategory != new_subcategory:
                            ws.update_cell(row_idx, col_idx['소분류'] + 1, new_subcategory)
                            updated_fields.append(f"소분류: {current_subcategory} → {new_subcategory}")
                    
                    # 제품이 변경된 경우
                    if col_idx.get('제품') is not None:
                        current_product = current_row[col_idx['제품']] if col_idx['제품'] < len(current_row) else ''
                        new_product = clean_product_name(row['제품'].get())
                        if current_product != new_product:
                            ws.update_cell(row_idx, col_idx['제품'] + 1, new_product)
                            updated_fields.append(f"제품: {current_product} → {new_product}")
                    
                    # 수량이 변경된 경우
                    if col_idx.get('수량') is not None:
                        current_quantity = current_row[col_idx['수량']] if col_idx['수량'] < len(current_row) else ''
                        new_quantity = row['수량'].get()
                        if current_quantity != new_quantity:
                            ws.update_cell(row_idx, col_idx['수량'] + 1, new_quantity)
                            updated_fields.append(f"수량: {current_quantity} → {new_quantity}")
                    
                    # 판매 단가가 변경된 경우
                    if col_idx.get('판매 단가') is not None:
                        current_unit_price = current_row[col_idx['판매 단가']] if col_idx['판매 단가'] < len(current_row) else ''
                        new_unit_price = row['판매 단가'].get()
                        if current_unit_price != new_unit_price:
                            ws.update_cell(row_idx, col_idx['판매 단가'] + 1, new_unit_price)
                            updated_fields.append(f"판매 단가: {current_unit_price} → {new_unit_price}")
                    
                    # 원가가 변경된 경우
                    if col_idx.get('원가') is not None:
                        current_cost = current_row[col_idx['원가']] if col_idx['원가'] < len(current_row) else ''
                        new_cost = row['원가'].get()
                        if current_cost != new_cost:
                            ws.update_cell(row_idx, col_idx['원가'] + 1, new_cost)
                            updated_fields.append(f"원가: {current_cost} → {new_cost}")
                    
                    # 만료일이 변경된 경우
                    if col_idx.get('만료일') is not None:
                        current_expiry = current_row[col_idx['만료일']] if col_idx['만료일'] < len(current_row) else ''
                        new_expiry = row['만료일'].get()
                        if current_expiry != new_expiry:
                            ws.update_cell(row_idx, col_idx['만료일'] + 1, new_expiry)
                            updated_fields.append(f"만료일: {current_expiry} → {new_expiry}")
                    
                    # 예상 발주일이 변경된 경우 (전역 변수 사용)
                    if col_idx.get('예상 발주일') is not None:
                        current_forecast = current_row[col_idx['예상 발주일']] if col_idx['예상 발주일'] < len(current_row) else ''
                        new_forecast = self.forecast_var.get()
                        if current_forecast != new_forecast:
                            ws.update_cell(row_idx, col_idx['예상 발주일'] + 1, new_forecast)
                            updated_fields.append(f"예상 발주일: {current_forecast} → {new_forecast}")
                    
                    # 매입처가 변경된 경우
                    if col_idx.get('매입처') is not None:
                        current_supplier = current_row[col_idx['매입처']] if col_idx['매입처'] < len(current_row) else ''
                        new_supplier = row['매입처'].get()
                        if current_supplier != new_supplier:
                            ws.update_cell(row_idx, col_idx['매입처'] + 1, new_supplier)
                            updated_fields.append(f"매입처: {current_supplier} → {new_supplier}")
            
            # 수식 재계산 (변경된 값이 있으면)
            if updated_fields:
                print(f"행 {row_idx}에서 변경된 필드: {updated_fields}")
                
                # 수식 업데이트 (FCST 열 추가로 인해 한 칸씩 밀림)
                if col_idx.get('판매 합계') is not None:
                    ws.update_cell(row_idx, col_idx['판매 합계'] + 1, f'=K{row_idx}*L{row_idx}')  # 수량(K) × 판매 단가(L)
                if col_idx.get('원가계') is not None:
                    ws.update_cell(row_idx, col_idx['원가계'] + 1, f'=N{row_idx}*K{row_idx}')  # 원가(N) × 수량(K)
                if col_idx.get('이익액') is not None:
                    ws.update_cell(row_idx, col_idx['이익액'] + 1, f'=M{row_idx}-O{row_idx}')  # 판매 합계(M) - 원가계(O)
                if col_idx.get('이익율') is not None:
                    ws.update_cell(row_idx, col_idx['이익율'] + 1, f'=IF(M{row_idx}=0,0,P{row_idx}/M{row_idx})')  # 이익액(P) / 판매 합계(M)
                
                print(f"행 {row_idx} 업데이트 완료")
            else:
                print(f"행 {row_idx}에서 변경된 값이 없습니다.")
                
        except Exception as e:
            print(f"행 {row_idx} 업데이트 중 오류: {e}")
            import traceback
            traceback.print_exc()

    def refresh_parent_window(self):
        try:
            if self.callback and callable(self.callback):
                self.callback()
            elif hasattr(self.parent, 'load_data'):
                self.parent.load_data()
            elif hasattr(self.parent.master, 'load_data'):
                self.parent.master.load_data()
            elif hasattr(self.parent.master.master, 'load_data'):
                self.parent.master.master.load_data()
        except Exception as e:
            print(f"부모 창 새로고침 오류: {e}")
            pass

    def refresh_quote_data(self):
        """견적서 데이터 새로고침"""
        try:
            # 부모 창에서 견적서 데이터 새로고침 메서드 호출
            if hasattr(self.parent, 'refresh_quote_data'):
                self.parent.refresh_quote_data()
            elif hasattr(self.parent.master, 'refresh_quote_data'):
                self.parent.master.refresh_quote_data()
            elif hasattr(self.parent.master.master, 'refresh_quote_data'):
                self.parent.master.master.refresh_quote_data()
            else:
                print("견적서 데이터 새로고침 메서드를 찾을 수 없습니다.")
        except Exception as e:
            print(f"견적서 데이터 새로고침 오류: {e}")

    def create_exchange_tables_side_by_side(self):
        """환율 정보와 가격표 환율을 나란히 배치"""
        # 환율 정보와 가격표 환율을 담을 컨테이너 프레임
        container_frame = tk.Frame(self.win)
        container_frame.pack(pady=(10, 10), padx=10, fill='x')
        
        # 환율 정보 (왼쪽)
        exchange_frame = tk.LabelFrame(container_frame, text="환율 정보", font=("맑은 고딕", 9, "bold"), bg="white", relief="solid", bd=1, labelanchor='nw')
        exchange_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        headers = ['구분', '환율']
        for i, h in enumerate(headers):
            tk.Label(exchange_frame, text=h, relief='solid', width=16, bg='#F1F5F9', fg="#374151", font=("맑은 고딕", 10, "bold"), anchor='center').grid(row=0, column=i, sticky='nsew')
        
        usd_rate = self.get_naver_fx_rate('USD')
        eur_rate = self.get_naver_fx_rate('EUR')
        tk.Label(exchange_frame, text='USD', relief='solid', width=16, bg='white', anchor='center').grid(row=1, column=0, sticky='nsew')
        tk.Label(exchange_frame, text=usd_rate, relief='solid', width=16, bg='white', anchor='e').grid(row=1, column=1, sticky='nsew')
        tk.Label(exchange_frame, text='EURO', relief='solid', width=16, bg='white', anchor='center').grid(row=2, column=0, sticky='nsew')
        tk.Label(exchange_frame, text=eur_rate, relief='solid', width=16, bg='white', anchor='e').grid(row=2, column=1, sticky='nsew')
        
        # 가격표 환율 (오른쪽)
        price_fx_frame = tk.LabelFrame(container_frame, text="가격표 환율", font=("맑은 고딕", 9, "bold"), bg="white", relief="solid", bd=1, labelanchor='nw')
        price_fx_frame.pack(side='right', fill='both', expand=True, padx=(5, 0))
        
        headers = ['구분', '환율', '적용 환율', '']
        for i, h in enumerate(headers):
            tk.Label(price_fx_frame, text=h, relief='solid', width=12, bg='#F1F5F9', fg="#374151", font=("맑은 고딕", 10, "bold"), anchor='center').grid(row=0, column=i, sticky='nsew')
        
        usd_rate, eur_rate = self.get_price_fx_values()
        self.usd_apply_var = tk.StringVar()
        self.eur_apply_var = tk.StringVar()
        
        # USD 행
        tk.Label(price_fx_frame, text='USD', relief='solid', width=12, bg='white', anchor='center').grid(row=1, column=0, sticky='nsew')
        self.usd_rate_var = tk.StringVar(value=usd_rate)
        tk.Entry(price_fx_frame, textvariable=self.usd_rate_var, width=12, state='readonly', justify='right', relief='solid', font=("맑은 고딕", 10), bg='white').grid(row=1, column=1, sticky='nsew')
        usd_apply_entry = tk.Entry(price_fx_frame, textvariable=self.usd_apply_var, width=12, justify='right', relief='solid', font=("맑은 고딕", 10), bg='white')
        usd_apply_entry.grid(row=1, column=2, sticky='nsew')
        tk.Button(price_fx_frame, text='적용', width=6, command=lambda: self.apply_price_fx('USD'), font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white", relief="flat").grid(row=1, column=3, sticky='nsew')
        
        # EUR 행
        tk.Label(price_fx_frame, text='EURO', relief='solid', width=12, bg='white', anchor='center').grid(row=2, column=0, sticky='nsew')
        self.eur_rate_var = tk.StringVar(value=eur_rate)
        tk.Entry(price_fx_frame, textvariable=self.eur_rate_var, width=12, state='readonly', justify='right', relief='solid', font=("맑은 고딕", 10), bg='white').grid(row=2, column=1, sticky='nsew')
        eur_apply_entry = tk.Entry(price_fx_frame, textvariable=self.eur_apply_var, width=12, justify='right', relief='solid', font=("맑은 고딕", 10), bg='white')
        eur_apply_entry.grid(row=2, column=2, sticky='nsew')
        tk.Button(price_fx_frame, text='적용', width=6, command=lambda: self.apply_price_fx('EUR'), font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white", relief="flat").grid(row=2, column=3, sticky='nsew')

    def get_naver_fx_rate(self, currency):
        try:
            import requests
            from bs4 import BeautifulSoup
            url = 'https://finance.naver.com/marketindex/exchangeList.naver'
            resp = requests.get(url, timeout=3)
            soup = BeautifulSoup(resp.text, 'html.parser')
            rows = soup.select('table.tbl_exchange tbody tr')
            for row in rows:
                tds = row.find_all('td')
                if len(tds) >= 2:
                    name = tds[0].get_text(strip=True)
                    rate = tds[1].get_text(strip=True).replace(',', '')
                    if currency == 'USD' and 'USD' in name:
                        return f'{float(rate):,.2f}'
                    if currency == 'EUR' and ('EUR' in name or '유로' in name):
                        return f'{float(rate):,.2f}'
            return ''
        except Exception as e:
            print(f'환율 정보 오류: {e}')
            return ''

    def get_product_info_from_price_list(self, product_name):
        """price_list에서 제품 정보 가져오기 (캐시 최적화)"""
        try:
            # 캐시 키 생성
            cache_key = f"product_info_{hashlib.md5(product_name.encode()).hexdigest()}"
            
            # 캐시에서 먼저 확인
            from core.data_loader import get_cached_data, set_cached_data
            cached_info = get_cached_data(cache_key)
            if cached_info is not None:
                print(f"제품 정보를 캐시에서 로드: {product_name}")
                return cached_info
            
            # API 호출 최소화를 위해 safe_api_call 사용
            gc = get_global_auth()
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet1')
            
            # 모든 데이터 가져오기 (안전한 API 호출)
            from core.data_loader import safe_api_call
            all_data = safe_api_call(lambda: ws.get_all_records())
            
            if all_data:
                print(f"제품 검색 중: '{product_name}'")
                print(f"price_list 데이터 개수: {len(all_data)}")
                
                # 제품명으로 검색
                for i, row in enumerate(all_data):
                    if '제품' in row and row['제품'] == product_name:
                        product_info = {
                            '구분': row.get('구분', ''),
                            'USD': row.get('USD', ''),
                            'EUR': row.get('EUR', ''),
                            'DC': row.get('DC', '')
                        }
                        print(f"제품 정보 찾음 (행 {i+1}): {product_info}")
                        # 캐시에 저장
                        set_cached_data(cache_key, product_info)
                        return product_info
                
                print(f"제품 '{product_name}'을 찾을 수 없습니다.")
                # 디버깅을 위해 처음 몇 개 제품명 출력
                for i, row in enumerate(all_data[:5]):
                    if '제품' in row:
                        print(f"  행 {i+1}: '{row['제품']}'")
            
            # 제품을 찾지 못한 경우 빈 정보를 캐시에 저장 (API 호출 방지)
            empty_info = {'구분': '', 'USD': '', 'EUR': '', 'DC': ''}
            set_cached_data(cache_key, empty_info)
            return None
            
        except Exception as e:
            print(f"price_list에서 제품 정보 읽기 오류: {e}")
            # 오류 발생 시 빈 정보 반환 (API 호출 방지)
            return {'구분': '', 'USD': '', 'EUR': '', 'DC': ''}

    def get_price_fx_values(self):
        # price_list Sheet2 A2(USD), B2(EUR) 값 읽기
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet2')
            usd = ws.acell('A2').value
            eur = ws.acell('B2').value
            
            # 값이 None이거나 빈 문자열인 경우 기본값 사용
            if not usd or usd == '' or usd == 'None':
                usd = '1,390.00'  # 기본 USD 환율
                print("견적 수정: USD 환율이 비어있어 기본값을 사용합니다.")
            else:
                usd = self.format_fx_value(usd)
                
            if not eur or eur == '' or eur == 'None':
                eur = '1,630.00'  # 기본 EUR 환율
                print("견적 수정: EUR 환율이 비어있어 기본값을 사용합니다.")
            else:
                eur = self.format_fx_value(eur)
            
            print(f"견적 수정: 환율 데이터 로드 완료 - USD: {usd}, EUR: {eur}")
            return usd, eur
        except Exception as e:
            print(f'가격표 환율 읽기 오류: {e}')
            print("견적 수정: 기본값으로 계속 진행합니다.")
            return '1,390.00', '1,630.00'  # 기본값 반환

    def format_fx_value(self, val):
        try:
            if val is None or val == '' or str(val).strip() == '':
                return ''
            # 쉼표 제거 후 숫자로 변환
            clean_val = str(val).replace(',', '').strip()
            fval = float(clean_val)
            return f"{fval:,.2f}"
        except (ValueError, TypeError) as e:
            print(f"견적 수정: 환율 값 형식 변환 오류 - {val}, 오류: {e}")
            return str(val) if val else ''



    def apply_price_fx(self, currency=None):
        # 적용 환율을 price_list Sheet2 A2(USD), B2(EUR)에 업데이트
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet2')
            
            if currency == 'USD':
                value = self.usd_apply_var.get().strip()
                if not value:
                    print("견적 수정: USD 환율 값이 비어있습니다.")
                    return
                try:
                    # 숫자로 변환하여 유효성 검사
                    num_value = float(value.replace(',', ''))
                    value_fmt = f"{num_value:,.2f}"
                    ws.update_acell('A2', value_fmt)
                    self.usd_rate_var.set(value_fmt)
                    self.usd_apply_var.set(value_fmt)
                    print(f"견적 수정: USD 환율 업데이트 완료 - {value_fmt}")
                except ValueError as ve:
                    print(f"견적 수정: USD 환율 값 형식 오류 - {value}")
                    messagebox.showerror("오류", f"USD 환율 값이 올바르지 않습니다: {value}")
            elif currency == 'EUR':
                value = self.eur_apply_var.get().strip()
                if not value:
                    print("견적 수정: EUR 환율 값이 비어있습니다.")
                    return
                try:
                    # 숫자로 변환하여 유효성 검사
                    num_value = float(value.replace(',', ''))
                    value_fmt = f"{num_value:,.2f}"
                    ws.update_acell('B2', value_fmt)
                    self.eur_rate_var.set(value_fmt)
                    self.eur_apply_var.set(value_fmt)
                    print(f"견적 수정: EUR 환율 업데이트 완료 - {value_fmt}")
                except ValueError as ve:
                    print(f"견적 수정: EUR 환율 값 형식 오류 - {value}")
                    messagebox.showerror("오류", f"EUR 환율 값이 올바르지 않습니다: {value}")
            else:
                print(f"견적 수정: 알 수 없는 통화 - {currency}")
                return
                
            self.load_price_data()
            for row in self.quote_rows:
                self.on_product_change(row)
        except Exception as e:
            print(f'가격표 환율 적용 오류: {e}')
            messagebox.showerror("오류", f"환율 적용 중 오류가 발생했습니다:\n{e}")