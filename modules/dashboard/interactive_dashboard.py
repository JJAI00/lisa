import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import calendar
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES
from integrations.organization_manager import get_organization_manager, OrganizationFilterWidget
import numpy as np # matplotlib 고급 기능을 위한 import

# matplotlib 고급 기능을 위한 import
try:
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    from matplotlib.figure import Figure
    import matplotlib.dates as mdates
    from matplotlib.widgets import Button, RadioButtons, CheckButtons
    import matplotlib
    matplotlib.use('TkAgg')
    
    # 한글 폰트 설정
    import matplotlib.font_manager as fm
    import warnings
    warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')
    
    korean_fonts = ['나눔고딕', 'NanumGothic', '맑은 고딕', 'Malgun Gothic', 'Dotum', '돋움']
    font_found = False
    
    for font_name in korean_fonts:
        try:
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
        plt.rcParams['font.family'] = 'DejaVu Sans'
        print("한글 폰트를 찾을 수 없어 기본 폰트를 사용합니다.")
    
    print("matplotlib 로드 성공")
    MATPLOTLIB_AVAILABLE = True
except ImportError as e:
    MATPLOTLIB_AVAILABLE = False
    print(f"matplotlib이 설치되지 않았습니다: {e}")
    print("설치 명령어: pip install matplotlib")
except Exception as e:
    MATPLOTLIB_AVAILABLE = False
    print(f"matplotlib 로드 중 오류 발생: {e}")

class InteractiveDashboard:
    """인터랙티브 대시보드 클래스 (Tkinter + 고급 matplotlib)"""
    
    def __init__(self, parent, username=None, data_loader=None):
        self.parent = parent
        self.username = username or '테스트유저'
        self.data_loader = data_loader
        self.org_manager = get_organization_manager()
        
        # 데이터 저장
        self.quote_df = pd.DataFrame()
        self.renewal_df = pd.DataFrame()
        self.current_filtered_data = pd.DataFrame()
        
        # 차트 상태
        self.current_chart_type = 'status_pie'
        self.selected_data_point = None
        self.chart_history = []  # 차트 히스토리 (드릴다운용)
        
        # matplotlib 관련
        self.fig = None
        self.ax = None
        self.canvas = None
        self.toolbar = None
        
        # UI 생성
        self.create_widgets()
        self.load_data()
        
    def create_widgets(self):
        """UI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 상단 컨트롤 패널
        self.create_control_panel(main_frame)
        
        # 중간 통계 카드
        self.create_statistics_cards(main_frame)
        
        # 하단 차트 영역
        self.create_chart_area(main_frame)
        
    def create_control_panel(self, parent):
        """상단 컨트롤 패널 생성"""
        control_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        control_frame.pack(fill='x', pady=(0, 10))
        
        # 컨트롤 헤더
        control_header = tk.Frame(control_frame, bg="#F1F5F9", height=40)
        control_header.pack(fill='x')
        control_header.pack_propagate(False)
        
        tk.Label(control_header, text="대시보드 컨트롤", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 컨트롤 내용
        control_content = tk.Frame(control_frame, bg="white")
        control_content.pack(fill='x', padx=15, pady=15)
        
        # 조직 필터 위젯을 별도 프레임에 배치
        org_filter_frame = tk.Frame(control_content, bg="white")
        org_filter_frame.pack(fill='x', pady=(0, 10))
        self.org_filter = OrganizationFilterWidget(org_filter_frame, self.org_manager, self.username)
        
        # 차트 타입 선택을 별도 프레임에 배치
        chart_type_frame = tk.Frame(control_content, bg="white")
        chart_type_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(chart_type_frame, text="차트 타입:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=0, padx=(0, 5), pady=10)
        
        self.chart_type_var = tk.StringVar(value="status_pie")
        chart_types = [
            ("Status 분포", "status_pie"),
            ("월별 추이", "monthly_trend"),
            ("부서별 성과", "department_performance"),
            ("팀별 성과", "team_performance"),
            ("영업사원별 성과", "salesperson_performance"),
            ("제품별 분석", "product_analysis")
        ]
        
        chart_frame = tk.Frame(chart_type_frame, bg="white")
        chart_frame.grid(row=0, column=1, columnspan=3, padx=(0, 15), pady=10)
        
        for i, (text, value) in enumerate(chart_types):
            tk.Radiobutton(chart_frame, text=text, variable=self.chart_type_var, 
                          value=value, command=self.on_chart_type_changed,
                          font=("맑은 고딕", 9), bg="white").grid(row=0, column=i, padx=5)
        
        # 버튼들을 별도 프레임에 배치
        button_frame = tk.Frame(control_content, bg="white")
        button_frame.pack(fill='x', pady=(0, 10))
        
        # 새로고침 버튼
        refresh_btn = tk.Button(button_frame, text="🔄 새로고침", command=self.refresh_dashboard,
                               font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                               relief="flat", padx=15, pady=5)
        refresh_btn.pack(side='left', padx=(0, 10))
        
        # 뒤로가기 버튼 (드릴다운용)
        self.back_btn = tk.Button(button_frame, text="⬅ 뒤로가기", command=self.go_back,
                                 font=("맑은 고딕", 10, "bold"), bg="#6B7280", fg="white",
                                 relief="flat", padx=15, pady=5, state='disabled')
        self.back_btn.pack(side='left')
    
    def create_statistics_cards(self, parent):
        """통계 카드 생성"""
        stats_frame = ttk.Frame(parent)
        stats_frame.pack(fill='x', pady=(0, 10))
        
        # Status별 통계 카드
        self.status_cards = {}
        status_colors = {
            '20% - 초기 단계': '#6B7280',
            '40% - 영업 진행 중': '#3B82F6',
            '60% - 고객사 검토 중': '#F59E0B',
            '80% - 발주서 수령': '#10B981',
            '100% - 세금계산서 발행': '#EF4444',
            'Drop - 중단': '#DC2626'
        }
        
        for i, (status, color) in enumerate(status_colors.items()):
            card_frame = tk.Frame(stats_frame, bg="white", relief="solid", bd=1)
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
        
        # 전체 통계 카드
        total_frame = tk.Frame(stats_frame, bg="white", relief="solid", bd=1)
        total_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        tk.Frame(total_frame, bg="#1E293B", height=4).pack(fill='x')
        
        content_frame = tk.Frame(total_frame, bg="white")
        content_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        tk.Label(content_frame, text="전체", fg="#374151", bg="white", 
                font=("맑은 고딕", 11, "bold")).pack(pady=(0, 10))
        
        self.total_count_label = tk.Label(content_frame, text="0건", fg="#1E293B", bg="white", 
                                         font=("맑은 고딕", 18, "bold"))
        self.total_count_label.pack()
    
    def create_chart_area(self, parent):
        """차트 영역 생성"""
        chart_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        chart_frame.pack(fill='both', expand=True)
        
        # 차트 헤더
        chart_header = tk.Frame(chart_frame, bg="#F1F5F9", height=40)
        chart_header.pack(fill='x')
        chart_header.pack_propagate(False)
        
        tk.Label(chart_header, text="인터랙티브 차트", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 차트 내용
        self.chart_content = tk.Frame(chart_frame, bg="white")
        self.chart_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 초기 차트 생성 (데이터 로드 후)
        self.parent.after(100, self.initialize_chart)
    
    def load_data(self):
        """데이터 로드"""
        try:
            if self.data_loader:
                self.quote_df = self.data_loader.get_quote_data()
                self.renewal_df = self.data_loader.get_renewal_data()
            else:
                # 직접 로드
                self.load_data_directly()
            
            # 조직 필터 적용
            self.apply_organization_filters()
            
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로드 중 오류가 발생했습니다:\n{str(e)}")
    
    def load_data_directly(self):
        """직접 데이터 로드"""
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            
            # 견적서 데이터
            sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
            ws = sh.worksheet('Sheet1')
            quote_data = ws.get_all_records()
            self.quote_df = pd.DataFrame(quote_data)
            
            # 리뉴얼 데이터
            sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
            ws = sh.worksheet('Sheet1')
            renewal_data = ws.get_all_records()
            self.renewal_df = pd.DataFrame(renewal_data)
            
        except Exception as e:
            print(f"데이터 직접 로드 오류: {e}")
    
    def initialize_chart(self):
        """차트 초기화 (데이터 로드 후 호출)"""
        try:
            # 데이터 로드 상태 확인
            print(f"견적서 데이터 크기: {len(self.quote_df)}행")
            print(f"리뉴얼 데이터 크기: {len(self.renewal_df)}행")
            print(f"필터링된 데이터 크기: {len(self.current_filtered_data)}행")
            
            # 통계 업데이트
            self.update_statistics()
            
            # 차트 생성
            self.create_chart()
            
        except Exception as e:
            print(f"차트 초기화 오류: {e}")
            messagebox.showerror("오류", f"차트 초기화 중 오류가 발생했습니다:\n{str(e)}")
    
    def apply_organization_filters(self):
        """조직 필터 적용"""
        if self.quote_df.empty:
            self.current_filtered_data = pd.DataFrame()
            return
        
        # 조직 필터 적용
        self.current_filtered_data = self.org_filter.apply_filters_to_data(self.quote_df)
        
        # 날짜 필터 적용 (이번달)
        if '일자' in self.current_filtered_data.columns:
            self.current_filtered_data['일자'] = pd.to_datetime(self.current_filtered_data['일자'], errors='coerce')
            today = datetime.now()
            start_of_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            end_of_month = (start_of_month + timedelta(days=32)).replace(day=1) - timedelta(seconds=1)
            
            mask = (self.current_filtered_data['일자'] >= start_of_month) & (self.current_filtered_data['일자'] <= end_of_month)
            self.current_filtered_data = self.current_filtered_data[mask]
    
    def on_chart_type_changed(self):
        """차트 타입 변경 시 호출"""
        self.current_chart_type = self.chart_type_var.get()
        self.chart_history = []  # 히스토리 초기화
        self.back_btn.config(state='disabled')
        self.create_chart()
    
    def go_back(self):
        """뒤로가기 (드릴다운)"""
        if len(self.chart_history) > 1:
            self.chart_history.pop()  # 현재 상태 제거
            previous_state = self.chart_history[-1]
            
            # 이전 상태로 복원
            self.current_chart_type = previous_state['chart_type']
            self.current_filtered_data = previous_state['filtered_data']
            self.selected_data_point = previous_state.get('selected_point')
            
            # UI 업데이트
            self.chart_type_var.set(self.current_chart_type)
            self.create_chart()
            
            # 뒤로가기 버튼 상태 업데이트
            if len(self.chart_history) <= 1:
                self.back_btn.config(state='disabled')
    
    def create_chart(self):
        """차트 생성"""
        print(f"차트 생성 시작 - matplotlib 사용 가능: {MATPLOTLIB_AVAILABLE}")
        
        if not MATPLOTLIB_AVAILABLE:
            print("matplotlib을 사용할 수 없습니다.")
            no_chart_label = tk.Label(self.chart_content, 
                                    text="matplotlib이 설치되지 않았습니다.\n\npip install matplotlib",
                                    font=("맑은 고딕", 12), fg="#6B7280", bg="white")
            no_chart_label.pack(expand=True)
            return
        
        # 기존 차트 위젯 제거
        for widget in self.chart_content.winfo_children():
            widget.destroy()
        
        print(f"필터링된 데이터 크기: {len(self.current_filtered_data)}행")
        if self.current_filtered_data.empty:
            print("데이터가 비어있습니다.")
            # 데이터가 없을 때
            no_data_label = tk.Label(self.chart_content, text="표시할 데이터가 없습니다.", 
                                   font=("맑은 고딕", 14), fg="#6B7280", bg="white")
            no_data_label.pack(expand=True)
            return
        
        try:
            print("Figure 생성 중...")
            # Figure 생성
            self.fig = Figure(figsize=(12, 8), facecolor='white')
            self.ax = self.fig.add_subplot(111)
            
            print(f"차트 타입: {self.current_chart_type}")
            # 차트 타입에 따른 생성
            if self.current_chart_type == 'status_pie':
                print("Status 파이 차트 생성 중...")
                self.create_status_pie_chart()
            elif self.current_chart_type == 'monthly_trend':
                print("월별 추이 차트 생성 중...")
                self.create_monthly_trend_chart()
            elif self.current_chart_type == 'department_performance':
                print("부서별 성과 차트 생성 중...")
                self.create_department_performance_chart()
            elif self.current_chart_type == 'team_performance':
                print("팀별 성과 차트 생성 중...")
                self.create_team_performance_chart()
            elif self.current_chart_type == 'salesperson_performance':
                print("영업사원별 성과 차트 생성 중...")
                self.create_salesperson_performance_chart()
            elif self.current_chart_type == 'product_analysis':
                print("제품별 분석 차트 생성 중...")
                self.create_product_analysis_chart()
            else:
                print("기본 Status 파이 차트 생성 중...")
                self.create_status_pie_chart()
            
            print("Canvas 생성 중...")
            # Canvas 생성
            self.canvas = FigureCanvasTkAgg(self.fig, self.chart_content)
            self.canvas.draw()
            
            print("Navigation Toolbar 추가 중...")
            # Navigation Toolbar 추가
            self.toolbar = NavigationToolbar2Tk(self.canvas, self.chart_content)
            self.toolbar.update()
            
            print("Canvas 배치 중...")
            # Canvas 배치
            self.canvas.get_tk_widget().pack(fill='both', expand=True)
            
            # 클릭 이벤트 연결
            self.canvas.mpl_connect('button_press_event', self.on_chart_click)
            
            print("차트 생성 완료!")
            
        except Exception as e:
            print(f"차트 생성 중 오류 발생: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"차트 생성 중 오류가 발생했습니다:\n{str(e)}")
    
    def create_status_pie_chart(self):
        """Status 분포 파이 차트"""
        print("Status 파이 차트 생성 시작")
        print(f"데이터 컬럼: {list(self.current_filtered_data.columns)}")
        
        if self.current_filtered_data.empty:
            print("데이터가 비어있습니다")
            self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        if 'Status' not in self.current_filtered_data.columns:
            print("Status 컬럼이 없습니다. 테스트 차트를 생성합니다.")
            # 테스트용 간단한 파이 차트 생성
            labels = ['테스트1', '테스트2', '테스트3']
            sizes = [30, 40, 30]
            colors = ['#3B82F6', '#F59E0B', '#10B981']
            
            wedges, texts, autotexts = self.ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%')
            self.ax.set_title('테스트 차트', fontsize=14, pad=20, fontweight='bold')
            
            # 데이터 저장 (드릴다운용)
            self.chart_data = {
                'type': 'pie',
                'data': pd.Series(sizes, index=labels),
                'wedges': wedges
            }
            return
        
        status_counts = self.current_filtered_data['Status'].value_counts()
        colors = ['#3B82F6', '#F59E0B', '#10B981', '#EF4444', '#6B7280', '#DC2626']
        
        wedges, texts, autotexts = self.ax.pie(
            status_counts.values, 
            labels=status_counts.index, 
            colors=colors[:len(status_counts)], 
            autopct='%1.1f%%',
            startangle=90,
            explode=[0.05] * len(status_counts)  # 약간의 간격
        )
        
        # 텍스트 스타일 설정
        for text in texts:
            text.set_fontsize(10)
        for autotext in autotexts:
            autotext.set_fontsize(9)
            autotext.set_color('white')
            autotext.set_weight('bold')
        
        self.ax.set_title('Status별 분포', fontsize=14, pad=20, fontweight='bold')
        
        # 범례 추가
        self.ax.legend(wedges, status_counts.index, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'pie',
            'data': status_counts,
            'wedges': wedges
        }
    
    def create_monthly_trend_chart(self):
        """월별 추이 차트"""
        if self.current_filtered_data.empty or '일자' not in self.current_filtered_data.columns:
            self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 월별 데이터 집계
        monthly_data = self.current_filtered_data.copy()
        monthly_data['월'] = monthly_data['일자'].dt.to_period('M')
        monthly_counts = monthly_data.groupby('월').size()
        
        # x축 라벨 생성
        x_labels = [str(x) for x in monthly_counts.index]
        x_positions = range(len(monthly_counts))
        
        # 바 차트 생성
        bars = self.ax.bar(x_positions, monthly_counts.values, color='#3B82F6', alpha=0.7, edgecolor='#1E40AF', linewidth=1)
        
        # x축 설정
        self.ax.set_xlabel('월', fontsize=12, fontweight='bold')
        self.ax.set_ylabel('건수', fontsize=12, fontweight='bold')
        self.ax.set_title('월별 견적서 추이', fontsize=14, pad=20, fontweight='bold')
        
        # x축 라벨 설정
        self.ax.set_xticks(x_positions)
        self.ax.set_xticklabels(x_labels, rotation=45, ha='right')
        
        # 그리드 추가
        self.ax.grid(True, axis='y', alpha=0.3)
        
        # 값 라벨 추가
        for bar, value in zip(bars, monthly_counts.values):
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{value}', ha='center', va='bottom', fontweight='bold')
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'bar',
            'data': monthly_counts,
            'bars': bars,
            'x_labels': x_labels
        }
    
    def create_department_performance_chart(self):
        """부서별 성과 차트"""
        if self.current_filtered_data.empty or '부서' not in self.current_filtered_data.columns:
            self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        dept_counts = self.current_filtered_data['부서'].value_counts()
        
        # 바 차트 생성
        bars = self.ax.bar(range(len(dept_counts)), dept_counts.values, color='#10B981', alpha=0.7, edgecolor='#059669', linewidth=1)
        
        # x축 설정
        self.ax.set_xlabel('부서', fontsize=12, fontweight='bold')
        self.ax.set_ylabel('건수', fontsize=12, fontweight='bold')
        self.ax.set_title('부서별 성과', fontsize=14, pad=20, fontweight='bold')
        
        # x축 라벨 설정
        self.ax.set_xticks(range(len(dept_counts)))
        self.ax.set_xticklabels(dept_counts.index, rotation=45, ha='right')
        
        # 그리드 추가
        self.ax.grid(True, axis='y', alpha=0.3)
        
        # 값 라벨 추가
        for bar, value in zip(bars, dept_counts.values):
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{value}', ha='center', va='bottom', fontweight='bold')
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'bar',
            'data': dept_counts,
            'bars': bars,
            'x_labels': dept_counts.index.tolist()
        }
    
    def create_team_performance_chart(self):
        """팀별 성과 차트"""
        if self.current_filtered_data.empty or '팀' not in self.current_filtered_data.columns:
            self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        team_counts = self.current_filtered_data['팀'].value_counts()
        
        # 바 차트 생성
        bars = self.ax.bar(range(len(team_counts)), team_counts.values, color='#F59E0B', alpha=0.7, edgecolor='#D97706', linewidth=1)
        
        # x축 설정
        self.ax.set_xlabel('팀', fontsize=12, fontweight='bold')
        self.ax.set_ylabel('건수', fontsize=12, fontweight='bold')
        self.ax.set_title('팀별 성과', fontsize=14, pad=20, fontweight='bold')
        
        # x축 라벨 설정
        self.ax.set_xticks(range(len(team_counts)))
        self.ax.set_xticklabels(team_counts.index, rotation=45, ha='right')
        
        # 그리드 추가
        self.ax.grid(True, axis='y', alpha=0.3)
        
        # 값 라벨 추가
        for bar, value in zip(bars, team_counts.values):
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{value}', ha='center', va='bottom', fontweight='bold')
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'bar',
            'data': team_counts,
            'bars': bars,
            'x_labels': team_counts.index.tolist()
        }
    
    def create_salesperson_performance_chart(self):
        """영업사원별 성과 차트"""
        if self.current_filtered_data.empty or '영업사원' not in self.current_filtered_data.columns:
            self.ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        salesperson_counts = self.current_filtered_data['영업사원'].value_counts()
        
        # 상위 10개만 표시
        top_salespeople = salesperson_counts.head(10)
        
        # 바 차트 생성
        bars = self.ax.bar(range(len(top_salespeople)), top_salespeople.values, color='#8B5CF6', alpha=0.7, edgecolor='#7C3AED', linewidth=1)
        
        # x축 설정
        self.ax.set_xlabel('영업사원', fontsize=12, fontweight='bold')
        self.ax.set_ylabel('건수', fontsize=12, fontweight='bold')
        self.ax.set_title('영업사원별 성과 (상위 10명)', fontsize=14, pad=20, fontweight='bold')
        
        # x축 라벨 설정
        self.ax.set_xticks(range(len(top_salespeople)))
        self.ax.set_xticklabels(top_salespeople.index, rotation=45, ha='right')
        
        # 그리드 추가
        self.ax.grid(True, axis='y', alpha=0.3)
        
        # 값 라벨 추가
        for bar, value in zip(bars, top_salespeople.values):
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{value}', ha='center', va='bottom', fontweight='bold')
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'bar',
            'data': top_salespeople,
            'bars': bars,
            'x_labels': top_salespeople.index.tolist()
        }
    
    def create_product_analysis_chart(self):
        """제품별 분석 차트 (원형 차트 + 막대 차트)"""
        if self.current_filtered_data.empty or '제품' not in self.current_filtered_data.columns:
            self.ax.text(0.5, 0.5, '제품 데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 제품별 건수 집계
        product_counts = self.current_filtered_data['제품'].value_counts()
        
        if product_counts.empty:
            self.ax.text(0.5, 0.5, '제품 데이터가 없습니다', ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 상위 10개만 표시 (나머지는 '기타'로)
        if len(product_counts) > 10:
            top_products = product_counts.head(9)
            other_count = product_counts.iloc[9:].sum()
            top_products = pd.concat([top_products, pd.Series({'기타': other_count})])
        else:
            top_products = product_counts
        
        # 서브플롯 생성 (1행 2열)
        self.fig.clear()
        
        # 원형 차트 (왼쪽)
        ax1 = self.fig.add_subplot(1, 2, 1)
        
        # 색상 설정
        colors = plt.cm.Set3(np.linspace(0, 1, len(top_products)))
        
        # 원형 차트 생성
        wedges, texts, autotexts = ax1.pie(top_products.values, 
                                          labels=top_products.index, 
                                          autopct='%1.1f%%',
                                          colors=colors,
                                          startangle=90)
        
        ax1.set_title('제품별 비율', fontsize=14, fontweight='bold', pad=20)
        
        # 막대 차트 (오른쪽)
        ax2 = self.fig.add_subplot(1, 2, 2)
        
        # 막대 차트 생성
        bars = ax2.bar(range(len(top_products)), top_products.values, 
                      color='#3B82F6', alpha=0.7, edgecolor='#2563EB', linewidth=1)
        
        # x축 설정
        ax2.set_xlabel('제품', fontsize=12, fontweight='bold')
        ax2.set_ylabel('건수', fontsize=12, fontweight='bold')
        ax2.set_title('제품별 건수', fontsize=14, pad=20, fontweight='bold')
        
        # x축 라벨 설정
        ax2.set_xticks(range(len(top_products)))
        ax2.set_xticklabels(top_products.index, rotation=45, ha='right')
        
        # 그리드 추가
        ax2.grid(True, axis='y', alpha=0.3)
        
        # 값 라벨 추가
        for bar, value in zip(bars, top_products.values):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{value}', ha='center', va='bottom', fontweight='bold')
        
        # 레이아웃 조정
        self.fig.tight_layout()
        
        # 데이터 저장 (드릴다운용)
        self.chart_data = {
            'type': 'product_analysis',
            'data': top_products,
            'wedges': wedges,
            'bars': bars,
            'x_labels': top_products.index.tolist(),
            'ax1': ax1,
            'ax2': ax2
        }
        
        # 현재 축을 원형 차트로 설정 (클릭 이벤트용)
        self.ax = ax1
    
    def on_chart_click(self, event):
        """차트 클릭 이벤트 처리"""
        if event.inaxes != self.ax:
            return
        
        # 현재 상태를 히스토리에 저장
        current_state = {
            'chart_type': self.current_chart_type,
            'filtered_data': self.current_filtered_data.copy(),
            'selected_point': self.selected_data_point
        }
        
        if not self.chart_history or self.chart_history[-1] != current_state:
            self.chart_history.append(current_state)
        
        # 클릭된 위치 확인
        if self.chart_data['type'] == 'pie':
            self.handle_pie_click(event)
        elif self.chart_data['type'] == 'bar':
            self.handle_bar_click(event)
        elif self.chart_data['type'] == 'product_analysis':
            self.handle_product_analysis_click(event)
        
        # 뒤로가기 버튼 활성화
        if len(self.chart_history) > 1:
            self.back_btn.config(state='normal')
    
    def handle_pie_click(self, event):
        """파이 차트 클릭 처리"""
        if 'wedges' not in self.chart_data:
            return
        
        wedges = self.chart_data['wedges']
        data = self.chart_data['data']
        
        for i, wedge in enumerate(wedges):
            if wedge.contains_point([event.x, event.y]):
                # 선택된 Status의 데이터만 필터링
                selected_status = data.index[i]
                filtered_data = self.current_filtered_data[self.current_filtered_data['Status'] == selected_status]
                
                # 드릴다운: 해당 Status의 영업사원별 분석
                self.drill_down_to_salesperson(filtered_data, selected_status)
                break
    
    def handle_bar_click(self, event):
        """바 차트 클릭 처리"""
        if 'bars' not in self.chart_data:
            return
        
        bars = self.chart_data['bars']
        data = self.chart_data['data']
        x_labels = self.chart_data['x_labels']
        
        for i, bar in enumerate(bars):
            if bar.contains(event)[0]:
                # 선택된 항목의 데이터만 필터링
                selected_item = x_labels[i]
                
                if self.current_chart_type == 'department_performance':
                    # 부서 선택 시 해당 부서의 팀별 분석
                    self.drill_down_to_team(selected_item)
                elif self.current_chart_type == 'team_performance':
                    # 팀 선택 시 해당 팀의 영업사원별 분석
                    self.drill_down_to_salesperson_by_team(selected_item)
                elif self.current_chart_type == 'salesperson_performance':
                    # 영업사원 선택 시 해당 영업사원의 상세 분석
                    self.drill_down_to_salesperson_detail(selected_item)
                elif self.current_chart_type == 'product_analysis':
                    # 제품 선택 시 해당 제품의 상세 분석
                    self.drill_down_to_product_detail(selected_item)
                break
    
    def handle_product_analysis_click(self, event):
        """제품별 분석 차트 클릭 처리"""
        if 'wedges' not in self.chart_data:
            return
        
        wedges = self.chart_data['wedges']
        data = self.chart_data['data']
        
        for i, wedge in enumerate(wedges):
            if wedge.contains_point([event.x, event.y]):
                # 선택된 제품의 데이터만 필터링
                selected_product = data.index[i]
                filtered_data = self.current_filtered_data[self.current_filtered_data['제품'] == selected_product]
                
                # 드릴다운: 해당 제품의 상세 분석
                self.drill_down_to_product_detail(selected_product)
                break
    
    def drill_down_to_salesperson(self, filtered_data, status):
        """Status에서 영업사원별로 드릴다운"""
        if filtered_data.empty or '영업사원' not in filtered_data.columns:
            return
        
        # 데이터 업데이트
        self.current_filtered_data = filtered_data
        self.selected_data_point = f"Status: {status}"
        
        # 차트 타입 변경
        self.current_chart_type = 'salesperson_performance'
        self.chart_type_var.set('salesperson_performance')
        
        # 차트 다시 생성
        self.create_chart()
    
    def drill_down_to_team(self, department):
        """부서에서 팀별로 드릴다운"""
        if self.current_filtered_data.empty or '팀' not in self.current_filtered_data.columns:
            return
        
        # 해당 부서의 데이터만 필터링
        filtered_data = self.current_filtered_data[self.current_filtered_data['부서'] == department]
        
        # 데이터 업데이트
        self.current_filtered_data = filtered_data
        self.selected_data_point = f"부서: {department}"
        
        # 차트 타입 변경
        self.current_chart_type = 'team_performance'
        self.chart_type_var.set('team_performance')
        
        # 차트 다시 생성
        self.create_chart()
    
    def drill_down_to_salesperson_by_team(self, team):
        """팀에서 영업사원별로 드릴다운"""
        if self.current_filtered_data.empty or '영업사원' not in self.current_filtered_data.columns:
            return
        
        # 해당 팀의 데이터만 필터링
        filtered_data = self.current_filtered_data[self.current_filtered_data['팀'] == team]
        
        # 데이터 업데이트
        self.current_filtered_data = filtered_data
        self.selected_data_point = f"팀: {team}"
        
        # 차트 타입 변경
        self.current_chart_type = 'salesperson_performance'
        self.chart_type_var.set('salesperson_performance')
        
        # 차트 다시 생성
        self.create_chart()
    
    def drill_down_to_salesperson_detail(self, salesperson):
        """영업사원 상세 분석"""
        if self.current_filtered_data.empty:
            return
        
        # 해당 영업사원의 데이터만 필터링
        filtered_data = self.current_filtered_data[self.current_filtered_data['영업사원'] == salesperson]
        
        # 상세 정보 표시 다이얼로그
        self.show_salesperson_detail(salesperson, filtered_data)
    
    def drill_down_to_product_detail(self, selected_product):
        """제품별 상세 분석으로 드릴다운"""
        # 선택된 제품의 데이터만 필터링
        product_data = self.current_filtered_data[self.current_filtered_data['제품'] == selected_product]
        
        if product_data.empty:
            messagebox.showinfo("알림", f"'{selected_product}' 제품의 상세 데이터가 없습니다.")
            return
        
        # 차트 제목 업데이트
        self.ax.clear()
        
        # 서브플롯 생성 (2행 2열)
        self.fig.clear()
        
        # 1. 고객사별 분포 (원형 차트)
        ax1 = self.fig.add_subplot(2, 2, 1)
        if '고객사명' in product_data.columns:
            customer_counts = product_data['고객사명'].value_counts().head(8)
            if not customer_counts.empty:
                colors = plt.cm.Set3(np.linspace(0, 1, len(customer_counts)))
                wedges, texts, autotexts = ax1.pie(customer_counts.values, 
                                                  labels=customer_counts.index, 
                                                  autopct='%1.1f%%',
                                                  colors=colors)
                ax1.set_title(f'{selected_product} - 고객사별 분포', fontsize=12, fontweight='bold')
        
        # 2. 영업사원별 분포 (막대 차트)
        ax2 = self.fig.add_subplot(2, 2, 2)
        if '영업사원' in product_data.columns:
            sales_counts = product_data['영업사원'].value_counts().head(8)
            if not sales_counts.empty:
                bars = ax2.bar(range(len(sales_counts)), sales_counts.values, 
                              color='#10B981', alpha=0.7)
                ax2.set_xlabel('영업사원')
                ax2.set_ylabel('건수')
                ax2.set_title(f'{selected_product} - 영업사원별 분포', fontsize=12, fontweight='bold')
                ax2.set_xticks(range(len(sales_counts)))
                ax2.set_xticklabels(sales_counts.index, rotation=45, ha='right')
        
        # 3. 월별 추이 (선 차트)
        ax3 = self.fig.add_subplot(2, 2, 3)
        if '일자' in product_data.columns:
            product_data['월'] = pd.to_datetime(product_data['일자']).dt.to_period('M')
            monthly_counts = product_data['월'].value_counts().sort_index().head(12)
            if not monthly_counts.empty:
                ax3.plot(range(len(monthly_counts)), monthly_counts.values, 
                        marker='o', color='#3B82F6', linewidth=2)
                ax3.set_xlabel('월')
                ax3.set_ylabel('건수')
                ax3.set_title(f'{selected_product} - 월별 추이', fontsize=12, fontweight='bold')
                ax3.grid(True, alpha=0.3)
        
        # 4. Status별 분포 (막대 차트)
        ax4 = self.fig.add_subplot(2, 2, 4)
        if 'Status' in product_data.columns:
            status_counts = product_data['Status'].value_counts()
            if not status_counts.empty:
                bars = ax4.bar(range(len(status_counts)), status_counts.values, 
                              color='#F59E0B', alpha=0.7)
                ax4.set_xlabel('Status')
                ax4.set_ylabel('건수')
                ax4.set_title(f'{selected_product} - Status별 분포', fontsize=12, fontweight='bold')
                ax4.set_xticks(range(len(status_counts)))
                ax4.set_xticklabels(status_counts.index, rotation=45, ha='right')
        
        # 레이아웃 조정
        self.fig.tight_layout()
        
        # 차트 업데이트
        self.canvas.draw()
        
        # 데이터 저장
        self.chart_data = {
            'type': 'product_detail',
            'product': selected_product,
            'data': product_data
        }
        
        # 현재 축 설정
        self.ax = ax1
    
    def show_salesperson_detail(self, salesperson, data):
        """영업사원 상세 정보 표시"""
        detail_window = tk.Toplevel(self.parent)
        detail_window.title(f"영업사원 상세 정보: {salesperson}")
        detail_window.geometry("600x400")
        detail_window.configure(bg="white")
        
        # 상세 정보 표시
        info_frame = tk.Frame(detail_window, bg="white")
        info_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 기본 정보
        tk.Label(info_frame, text=f"영업사원: {salesperson}", 
                font=("맑은 고딕", 16, "bold"), fg="#374151", bg="white").pack(pady=(0, 20))
        
        # 통계 정보
        stats_frame = tk.Frame(info_frame, bg="white")
        stats_frame.pack(fill='x', pady=(0, 20))
        
        total_count = len(data)
        status_counts = data['Status'].value_counts() if 'Status' in data.columns else pd.Series()
        
        tk.Label(stats_frame, text=f"총 견적서: {total_count}건", 
                font=("맑은 고딕", 12), fg="#374151", bg="white").pack()
        
        # Status별 통계
        if not status_counts.empty:
            status_text = "Status별 분포:\n"
            for status, count in status_counts.items():
                percentage = (count / total_count) * 100
                status_text += f"  {status}: {count}건 ({percentage:.1f}%)\n"
            
            tk.Label(stats_frame, text=status_text, 
                    font=("맑은 고딕", 10), fg="#6B7280", bg="white", justify='left').pack(pady=(10, 0))
    
    def show_product_detail(self, product, data):
        """제품 상세 정보 표시"""
        detail_window = tk.Toplevel(self.parent)
        detail_window.title(f"제품 상세 정보: {product}")
        detail_window.geometry("600x400")
        detail_window.configure(bg="white")
        
        # 상세 정보 표시
        info_frame = tk.Frame(detail_window, bg="white")
        info_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 기본 정보
        tk.Label(info_frame, text=f"제품: {product}", 
                font=("맑은 고딕", 16, "bold"), fg="#374151", bg="white").pack(pady=(0, 20))
        
        # 통계 정보
        stats_frame = tk.Frame(info_frame, bg="white")
        stats_frame.pack(fill='x', pady=(0, 20))
        
        total_count = len(data)
        status_counts = data['Status'].value_counts() if 'Status' in data.columns else pd.Series()
        
        tk.Label(stats_frame, text=f"총 견적서: {total_count}건", 
                font=("맑은 고딕", 12), fg="#374151", bg="white").pack()
        
        # Status별 통계
        if not status_counts.empty:
            status_text = "Status별 분포:\n"
            for status, count in status_counts.items():
                percentage = (count / total_count) * 100
                status_text += f"  {status}: {count}건 ({percentage:.1f}%)\n"
            
            tk.Label(stats_frame, text=status_text, 
                    font=("맑은 고딕", 10), fg="#6B7280", bg="white", justify='left').pack(pady=(10, 0))
    
    def refresh_dashboard(self):
        """대시보드 새로고침"""
        try:
            # 데이터 다시 로드
            self.load_data()
            
            # 통계 업데이트
            self.update_statistics()
            
            # 차트 다시 생성
            self.create_chart()
            
        except Exception as e:
            messagebox.showerror("오류", f"대시보드 새로고침 중 오류가 발생했습니다:\n{str(e)}")
    
    def update_statistics(self):
        """통계 업데이트"""
        if self.current_filtered_data.empty:
            # 모든 카드를 0으로 설정
            for status in self.status_cards:
                self.status_cards[status].config(text="0건")
            self.total_count_label.config(text="0건")
            return
        
        # Status별 통계
        if 'Status' in self.current_filtered_data.columns:
            status_counts = self.current_filtered_data['Status'].value_counts()
            
            for status in self.status_cards:
                count = status_counts.get(status, 0)
                self.status_cards[status].config(text=f"{count}건")
        
        # 전체 통계
        total_count = len(self.current_filtered_data)
        self.total_count_label.config(text=f"{total_count}건")

def create_interactive_dashboard(parent, username=None, data_loader=None):
    """인터랙티브 대시보드 생성 함수"""
    return InteractiveDashboard(parent, username, data_loader) 