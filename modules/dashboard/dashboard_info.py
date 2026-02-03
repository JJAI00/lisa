import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import calendar
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES

# matplotlib이 없어도 작동하도록 선택적 import
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib
    matplotlib.use('TkAgg')
    
    # 한글 폰트 설정 (개선된 방식)
    import matplotlib.font_manager as fm
    import warnings
    
    # matplotlib 경고 숨기기
    warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')
    warnings.filterwarnings('ignore', category=UserWarning, message='.*Glyph.*missing from font.*')
    warnings.filterwarnings('ignore', category=UserWarning, message='.*findfont.*')
    
    # 나눔고딕 폰트 우선 설정
    korean_fonts = ['나눔고딕', 'NanumGothic', '맑은 고딕', 'Malgun Gothic', 'Dotum', '돋움']
    font_found = False
    
    # 폰트 캐시 초기화 (버전 호환성 고려)
    try:
        fm._rebuild()
    except AttributeError:
        try:
            fm.fontManager.rebuild()
        except AttributeError:
            # 최신 버전에서는 자동으로 관리되므로 무시
            pass
    
    for font_name in korean_fonts:
        try:
            # 간단한 폰트 설정 방식
            font_prop = fm.FontProperties(family=font_name)
            if font_prop.get_name() != 'DejaVu Sans':
                plt.rcParams['font.family'] = font_name
                plt.rcParams['font.sans-serif'] = [font_name] + [f for f in plt.rcParams['font.sans-serif'] if f != font_name]
                font_found = True
                print(f"한글 폰트 설정 완료: {font_name}")
                break
        except Exception as e:
            print(f"폰트 {font_name} 설정 실패: {e}")
            continue
    
    if not font_found:
        # 시스템 폰트 중 한글 지원 폰트 찾기
        try:
            system_fonts = [f.name for f in fm.fontManager.ttflist]
            for font in system_fonts:
                if any(korean_char in font for korean_char in ['고딕', 'Gothic', 'Dotum', 'Malgun', 'Nanum']):
                    plt.rcParams['font.family'] = font
                    plt.rcParams['font.sans-serif'] = [font] + [f for f in plt.rcParams['font.sans-serif'] if f != font]
                    font_found = True
                    print(f"시스템 한글 폰트 설정 완료: {font}")
                    break
        except Exception as e:
            print(f"시스템 폰트 검색 실패: {e}")
            
    if not font_found:
        print("한글 폰트를 찾을 수 없어 기본 폰트를 사용합니다.")
        # 경고 메시지 완전히 억제
        plt.rcParams['font.family'] = 'DejaVu Sans'
        warnings.filterwarnings('ignore', category=UserWarning)
    
    # 한글 폰트가 설정되었는지 확인
    if font_found:
        print(f"최종 설정된 폰트: {plt.rcParams['font.family']}")
        print(f"폰트 목록: {plt.rcParams['font.sans-serif']}")
    
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("matplotlib이 설치되지 않았습니다. 차트 기능이 비활성화됩니다.")
    print("설치 명령어: pip install matplotlib")

class DashboardDialog:
    """월별 영업사원별 대시보드"""
    
    def __init__(self, parent, username=None):
        self.parent = parent
        self.username = username or '테스트유저'
        self.frame = parent  # Toplevel 대신 parent frame 사용
        
        # 데이터 로드
        self.load_quote_data()
        
        # UI 생성
        self.create_widgets()
        
    def load_quote_data(self):
        """quote_list에서 데이터 로드"""
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
            ws = sh.worksheet('Sheet1')
            
            # 모든 데이터 가져오기
            all_data = ws.get_all_values()
            if len(all_data) < 2:
                self.quote_df = pd.DataFrame()
                return
                
            # DataFrame 생성
            headers = all_data[0]
            data = all_data[1:]
            self.quote_df = pd.DataFrame(data, columns=headers)
            
            # 날짜 컬럼 변환
            if '일자' in self.quote_df.columns:
                self.quote_df['일자'] = pd.to_datetime(self.quote_df['일자'], errors='coerce')
                
        except Exception as e:
            print(f"데이터 로드 오류: {e}")
            self.quote_df = pd.DataFrame()
    
    def create_widgets(self):
        """UI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 상단 필터 프레임 (LISA 스타일과 통일)
        filter_frame = tk.Frame(main_frame, bg="white", relief="solid", bd=1)
        filter_frame.pack(fill='x', pady=(0, 10))
        
        # 필터 헤더
        filter_header = tk.Frame(filter_frame, bg="#F1F5F9", height=40)
        filter_header.pack(fill='x')
        filter_header.pack_propagate(False)
        
        tk.Label(filter_header, text="필터", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 필터 내용
        filter_content = tk.Frame(filter_frame, bg="white")
        filter_content.pack(fill='x', padx=15, pady=15)
        
        # 월 선택
        tk.Label(filter_content, text="월:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=0, padx=(0, 5), pady=5)
        self.month_var = tk.StringVar(value=datetime.now().strftime('%Y-%m'))
        month_entry = ttk.Entry(filter_content, textvariable=self.month_var, width=12, font=("맑은 고딕", 10))
        month_entry.grid(row=0, column=1, padx=(0, 15), pady=5)
        
        # 영업사원 선택
        tk.Label(filter_content, text="영업사원:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=2, padx=(0, 5), pady=5)
        self.salesperson_var = tk.StringVar(value="전체")
        salesperson_combo = ttk.Combobox(filter_content, textvariable=self.salesperson_var, 
                                        width=15, font=("맑은 고딕", 10))
        salesperson_combo['values'] = self.get_salespeople()
        salesperson_combo.grid(row=0, column=3, padx=(0, 15), pady=5)
        
        # 새로고침 버튼
        refresh_btn = tk.Button(filter_content, text="🔄 새로고침", command=self.refresh_dashboard,
                               font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                               relief="flat", padx=15, pady=5)
        refresh_btn.grid(row=0, column=4, padx=(0, 10), pady=5)
        
        # Status 변경 버튼
        status_btn = tk.Button(filter_content, text="📝 Status 변경", command=self.change_status,
                              font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white",
                              relief="flat", padx=15, pady=5)
        status_btn.grid(row=0, column=5, pady=5)
        
        # 통계 카드 프레임
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill='x', pady=(0, 10))
        
        # Status별 통계 카드
        self.create_status_cards(stats_frame)
        
        # 차트 프레임 (LISA 스타일과 통일)
        chart_frame = tk.Frame(main_frame, bg="white", relief="solid", bd=1)
        chart_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # 차트 헤더
        chart_header = tk.Frame(chart_frame, bg="#F1F5F9", height=40)
        chart_header.pack(fill='x')
        chart_header.pack_propagate(False)
        
        tk.Label(chart_header, text="진행률 차트", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 차트 내용
        chart_content = tk.Frame(chart_frame, bg="white")
        chart_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 차트 생성
        self.create_chart(chart_content)
        
        # 데이터 테이블 프레임 (LISA 스타일과 통일)
        table_frame = tk.Frame(main_frame, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # 테이블 헤더
        table_header = tk.Frame(table_frame, bg="#F1F5F9", height=40)
        table_header.pack(fill='x')
        table_header.pack_propagate(False)
        
        tk.Label(table_header, text="상세 데이터", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 테이블 내용
        table_content = tk.Frame(table_frame, bg="white")
        table_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 테이블 생성
        self.create_table(table_content)
        
        # 초기 데이터 로드
        self.refresh_dashboard()
    
    def get_salespeople(self):
        """영업사원 목록 반환"""
        if self.quote_df.empty or '영업사원' not in self.quote_df.columns:
            return ["전체"]
        
        salespeople = self.quote_df['영업사원'].dropna().unique().tolist()
        return ["전체"] + sorted(salespeople)
    
    def create_status_cards(self, parent):
        """Status별 통계 카드 생성"""
        # Status별 색상 정의 (LISA 디자인과 통일)
        status_colors = {
            '20% - 초기 단계': '#6B7280',         # 회색
            '40% - 영업 진행 중': '#3B82F6',      # 파란색
            '60% - 고객사 검토 중': '#F59E0B',    # 주황색
            '80% - 발주서 수령': '#10B981',       # 초록색
            '100% - 세금계산서 발행': '#EF4444',  # 빨간색
            'Drop - 중단': '#DC2626'              # 진한 빨간색
        }
        
        self.status_cards = {}
        
        for i, (status, color) in enumerate(status_colors.items()):
            # 카드 프레임 (LISA 스타일과 통일)
            card_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
            card_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            # 상단 색상 바
            color_bar = tk.Frame(card_frame, bg=color, height=4)
            color_bar.pack(fill='x')
            
            # 내용 프레임
            content_frame = tk.Frame(card_frame, bg="white")
            content_frame.pack(fill='both', expand=True, padx=15, pady=15)
            
            # Status 이름
            tk.Label(content_frame, text=status, fg="#374151", bg="white", 
                    font=("맑은 고딕", 11, "bold")).pack(pady=(0, 10))
            
            # 건수
            count_label = tk.Label(content_frame, text="0건", fg=color, bg="white", 
                                 font=("맑은 고딕", 18, "bold"))
            count_label.pack()
            
            self.status_cards[status] = count_label
    
    def create_chart(self, parent):
        """진행률 차트 생성"""
        if MATPLOTLIB_AVAILABLE:
            # matplotlib 차트
            self.fig, self.ax = plt.subplots(figsize=(10, 6))
            self.canvas = FigureCanvasTkAgg(self.fig, parent)
            self.canvas.get_tk_widget().pack(fill='both', expand=True)
        else:
            # matplotlib이 없을 때 대체 텍스트
            self.chart_label = tk.Label(parent, text="차트를 보려면 matplotlib을 설치하세요\n\npip install matplotlib", 
                                      font=("맑은 고딕", 12), fg="#6B7280", bg="white")
            self.chart_label.pack(fill='both', expand=True)
    
    def create_table(self, parent):
        """데이터 테이블 생성"""
        # Treeview 생성 (LISA 스타일과 통일)
        columns = ('일자', '영업사원', '고객사명', '거래처명', '제품', 'Status')
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', height=12)
        
        # 스타일 설정
        style = ttk.Style()
        style.configure("Treeview", font=("맑은 고딕", 10))
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"))
        
        # 컬럼 설정
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor='center')
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 배치
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def refresh_dashboard(self):
        """대시보드 새로고침"""
        try:
            # 필터 적용
            filtered_df = self.apply_filters()
            
            # Status별 통계 업데이트
            self.update_status_stats(filtered_df)
            
            # 차트 업데이트
            self.update_chart(filtered_df)
            
            # 테이블 업데이트
            self.update_table(filtered_df)
            
        except Exception as e:
            messagebox.showerror("오류", f"대시보드 새로고침 중 오류가 발생했습니다:\n{str(e)}")
    
    def apply_filters(self):
        """필터 적용"""
        if self.quote_df.empty:
            return pd.DataFrame()
        
        df = self.quote_df.copy()
        
        # 월 필터
        selected_month = self.month_var.get()
        if selected_month:
            try:
                year, month = selected_month.split('-')
                start_date = datetime(int(year), int(month), 1)
                if int(month) == 12:
                    end_date = datetime(int(year) + 1, 1, 1)
                else:
                    end_date = datetime(int(year), int(month) + 1, 1)
                
                df = df[(df['일자'] >= start_date) & (df['일자'] < end_date)]
            except:
                pass
        
        # 영업사원 필터
        selected_salesperson = self.salesperson_var.get()
        if selected_salesperson and selected_salesperson != "전체":
            df = df[df['영업사원'] == selected_salesperson]
        
        return df
    
    def update_status_stats(self, df):
        """Status별 통계 업데이트"""
        if df.empty:
            for status in self.status_cards:
                self.status_cards[status].config(text="0건")
            return
        
        # Status별 건수 계산
        status_counts = df['Status'].value_counts()
        
        for status in self.status_cards:
            count = status_counts.get(status, 0)
            self.status_cards[status].config(text=f"{count}건")
    
    def update_chart(self, df):
        """차트 업데이트"""
        if MATPLOTLIB_AVAILABLE:
            self.ax.clear()
            
            # 차트 그리기 전에 폰트 재설정
            try:
                import matplotlib.font_manager as fm
                import warnings
                
                # 폰트 관련 경고 메시지 억제
                warnings.filterwarnings('ignore', category=UserWarning, message='.*Glyph.*missing from font.*')
                warnings.filterwarnings('ignore', category=UserWarning, message='.*findfont.*')
                
                # 나눔고딕 폰트 재설정
                korean_fonts = ['나눔고딕', 'NanumGothic', '맑은 고딕', 'Malgun Gothic']
                font_found = False
                
                for font_name in korean_fonts:
                    try:
                        font_prop = fm.FontProperties(family=font_name)
                        if font_prop.get_name() != 'DejaVu Sans':
                            plt.rcParams['font.family'] = font_name
                            font_found = True
                            break
                    except:
                        continue
                
                if not font_found:
                    # 기본 폰트 사용
                    plt.rcParams['font.family'] = 'DejaVu Sans'
            except:
                # 폰트 설정 실패 시 기본값 사용
                plt.rcParams['font.family'] = 'DejaVu Sans'
            
            if df.empty:
                self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                           transform=self.ax.transAxes, fontsize=12)
            else:
                # Status별 건수
                status_counts = df['Status'].value_counts()
                
                # 원형 차트
                colors = ['#3B82F6', '#F59E0B', '#10B981', '#EF4444']
                wedges, texts, autotexts = self.ax.pie(status_counts.values, 
                                                       labels=status_counts.index,
                                                       colors=colors[:len(status_counts)],
                                                       autopct='%1.1f%%',
                                                       textprops={'fontsize': 10})
                
                self.ax.set_title('Status별 분포', fontsize=14, pad=20)
            
            self.canvas.draw()
        else:
            # matplotlib이 없을 때 차트 대신 통계 텍스트 표시
            if hasattr(self, 'chart_label'):
                if df.empty:
                    self.chart_label.config(text="데이터가 없습니다")
                else:
                    status_counts = df['Status'].value_counts()
                    chart_text = "Status별 분포:\n\n"
                    for status, count in status_counts.items():
                        percentage = (count / len(df)) * 100
                        chart_text += f"{status}: {count}건 ({percentage:.1f}%)\n"
                    self.chart_label.config(text=chart_text)
    
    def update_table(self, df):
        """테이블 업데이트"""
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if df.empty:
            return
        
        # 데이터 추가
        for _, row in df.iterrows():
            values = (
                row.get('일자', ''),
                row.get('영업사원', ''),
                row.get('고객사명', ''),
                row.get('거래처명', ''),
                row.get('제품', ''),
                row.get('Status', '')
            )
            self.tree.insert('', 'end', values=values)
    
    def change_status(self):
        """Status 변경 다이얼로그"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 필요", "Status를 변경할 견적서를 선택하세요.")
            return
        
        # 선택된 항목 정보
        item = self.tree.item(selected[0])
        row_values = item['values']
        
        # Status 선택 다이얼로그
        status_dialog = tk.Toplevel(self.frame)
        status_dialog.title("Status 변경")
        status_dialog.geometry("400x200")
        status_dialog.grab_set()
        
        ttk.Label(status_dialog, text="새로운 Status를 선택하세요:", font=("맑은 고딕", 12, "bold")).pack(pady=20)
        
        status_var = tk.StringVar()
        status_options = [
            "20% - 초기 단계",
            "40% - 영업 진행 중",
            "60% - 고객사 검토 중",
            "80% - 발주서 수령",
            "100% - 세금계산서 발행",
            "Drop - 중단"
        ]
        
        for status in status_options:
            ttk.Radiobutton(status_dialog, text=status, variable=status_var, value=status).pack(pady=5)
        
        def update_status():
            new_status = status_var.get()
            if not new_status:
                messagebox.showwarning("선택 필요", "Status를 선택하세요.")
                return
            
            # Google Sheets에서 Status 업데이트
            try:
                self.update_status_in_sheet(row_values, new_status)
                messagebox.showinfo("완료", "Status가 성공적으로 변경되었습니다.")
                status_dialog.destroy()
                self.refresh_dashboard()
            except Exception as e:
                messagebox.showerror("오류", f"Status 변경 중 오류가 발생했습니다:\n{str(e)}")
        
        ttk.Button(status_dialog, text="변경", command=update_status).pack(pady=20)
    
    def update_status_in_sheet(self, row_values, new_status):
        """Google Sheets에서 Status 업데이트"""
        creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
        ws = sh.worksheet('Sheet1')
        
        # 헤더 가져오기
        headers = ws.row_values(1)
        status_col_idx = headers.index('Status') + 1  # 1-based index
        
        # 전체 데이터에서 해당 행 찾기
        all_data = ws.get_all_values()
        target_row = None
        
        for row_idx, row_data in enumerate(all_data[1:], start=2):  # 헤더 제외
            if (len(row_data) > 0 and
                row_data[headers.index('일자')] == str(row_values[0]) and
                row_data[headers.index('영업사원')] == row_values[1] and
                row_data[headers.index('고객사명')] == row_values[2] and
                row_data[headers.index('거래처명')] == row_values[3] and
                row_data[headers.index('제품')] == row_values[4]):
                target_row = row_idx
                break
        
        if target_row:
            ws.update_cell(target_row, status_col_idx, new_status)
        else:
            raise Exception("해당 행을 찾을 수 없습니다.") 