import tkinter as tk
from tkinter import messagebox, ttk
import hashlib
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES
from modules.renewal.renewal_preview import RenewalMailerGUI
from modules.customer.customer_info import CustomerInfoGUI   # ← 이미 존재
from modules.purchase.purchase_sales_info import PurchaseSalesInfo   # ← 새로 추가
from certificates.license_certificate_info import open_dialog   # ← 새로 추가
from modules.dashboard.dashboard_foundry import create_foundry_dashboard   # ← Foundry 대시보드 추가
import requests
import sys, os
import socket
# 로딩 스크린 관련 코드 제거됨
from modules.quote.quote_info import QuoteInfoGUI  # 견적서 탭용 GUI 클래스 추가
from core.data_loader import DataLoader
from ui.loading_indicator import BackgroundLoader
from modules.renewal.expiry_alarm import show_expiry_alarm
from integrations.lisa_logging import setup_logger

# 모듈 로거 설정
logger = setup_logger('renewal_gui')

# PIL 라이브러리 import (이미지 처리용)
try:
    from PIL import Image, ImageTk, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("PIL 라이브러리가 설치되지 않았습니다. 이미지 기능이 제한됩니다.")
    print("설치 명령어: pip install Pillow")


# — 로그인용 구글 시트 정보 —
LOGIN_SHEET_ID   = '1KEcLicMV6p-RSXl2kg1AHyD2CmU8wABeXDlMTk0gqZY'
LOGIN_SHEET_NAME = 'Sheet1'

# --- 버전 읽기 함수 추가 ---
def get_version():
    try:
        with open(resource_path('VERSION.txt'), encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return "버전정보없음"

def hash_password(pw):
    """SHA-256 해시화"""
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def fetch_user_db():
    """구글시트에서 사용자 DB 로드"""
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    ws    = gc.open_by_key(LOGIN_SHEET_ID).worksheet(LOGIN_SHEET_NAME)
    return pd.DataFrame(ws.get_all_records())

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


class LoginWindow:
    """로그인 창"""
    def __init__(self, win, root):
        self.win = win
        self.root = root
        win.title(f"{get_version()} 로그인")
        win.geometry("350x250")
        win.configure(bg="#F7F9FB")
        win.resizable(False, False)
        
        # 로그인 창 아이콘 설정
        try:
            # 패키징된 환경과 개발 환경 모두에서 작동하도록 수정
            icon_path = resource_path('design/lisa.ico')
            if os.path.exists(icon_path):
                win.iconbitmap(icon_path)
            else:
                # 패키징된 환경에서는 절대 경로로 시도
                import sys
                if getattr(sys, 'frozen', False):
                    # PyInstaller로 패키징된 경우
                    base_path = sys._MEIPASS
                    icon_path = os.path.join(base_path, 'design', 'lisa.ico')
                    if os.path.exists(icon_path):
                        win.iconbitmap(icon_path)
        except Exception as e:
            print(f"로그인 창 아이콘 로드 실패: {e}")
            pass  # 아이콘 파일이 없거나 오류가 나도 무시

        # 카드 프레임
        card = tk.Frame(win, bg="white", relief="solid", bd=1)
        card.place(relx=0.5, rely=0.5, anchor="c", width=300, height=170)

        # ID
        tk.Label(card, text="ID:", bg="white", fg="#374151").place(x=30, y=25)
        self.id_entry = tk.Entry(card, bg="#F1F5F9", relief="flat")
        self.id_entry.place(x=80, y=25, width=180, height=28)

        # PW
        tk.Label(card, text="PW:", bg="white", fg="#374151").place(x=30, y=70)
        self.pw_entry = tk.Entry(card, bg="#F1F5F9", relief="flat", show="*")
        self.pw_entry.place(x=80, y=70, width=180, height=28)

        # 로그인 버튼
        login_btn = tk.Button(card, text="로그인",
                              bg="#3B82F6", fg="white", relief="flat", activebackground="#2563EB", activeforeground="white",
                              command=self.try_login)
        login_btn.place(x=30, y=120, width=230, height=32)
        login_btn.bind("<Enter>", lambda e: login_btn.config(bg="#2563EB"))
        login_btn.bind("<Leave>", lambda e: login_btn.config(bg="#3B82F6"))

        self.id_entry.focus()
        win.bind('<Return>', lambda e: self.try_login())

    def try_login(self):
        user_id = self.id_entry.get().strip()
        pw_hash = hash_password(self.pw_entry.get())
        try:
            df = fetch_user_db()
        except Exception as e:
            messagebox.showerror("오류", f"계정 정보를 불러오는 중 오류가 발생했습니다:\n{e}")
            return

        if '아이디' not in df.columns or '패스워드' not in df.columns:
            messagebox.showerror("오류", "계정 시트의 컬럼명이 '아이디', '패스워드'인지 확인해주세요.")
            return

        user_row = df[(df['아이디']==user_id)&(df['패스워드']==pw_hash)]
        if user_row.empty:
            messagebox.showerror("로그인 실패","아이디 또는 비밀번호가 올바르지 않습니다.")
            self.pw_entry.delete(0, tk.END)
            return

        username = user_row.iloc[0].get('이름', user_id)
        self.win.destroy()  # 로그인창만 닫기
        self.root.deiconify()  # 메인 윈도우 보이기
        
        # MainApp 초기화 및 Renewal 관리 화면으로 시작
        app = MainApp(self.root, username)
        app.show_frame("renewal")  # 명시적으로 Renewal 관리 화면 표시

class MainApp:
    """전체 앱 UI (sidebar + main_frame 구조)"""
    def __init__(self, root, username=None):
        self.current_frame = None
        self.root     = root
        self.username = username or '테스트유저'
        self.root.title(f"{get_version()} - {self.username}")
        self.root.geometry("1300x850")
        self.root.configure(bg="#F7F9FB")

        # 데이터 로더 초기화
        self.data_loader = DataLoader(KEY_FILE, SCOPES)
        self.background_loader = None
        
        # 데이터 매니저 초기화 및 백그라운드 동기화 시작
        try:
            from core.data_manager import DataManager
            self.data_manager = DataManager()
            self.data_manager.start_background_sync()
            logger.info("백그라운드 데이터 동기화 시작됨")
        except Exception as e:
            logger.error(f"데이터 매니저 초기화 실패: {e}")
        
        # 프로그램 종료 시 정리 함수 등록
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # --- sidebar ---
        self.sidebar = tk.Frame(self.root, width=200, bg="#1E293B")  # 더 어두운 색상으로 변경
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # 로고 섹션
        logo_frame = tk.Frame(self.sidebar, bg="#1E293B", height=80)
        logo_frame.pack(fill='x', pady=(20, 30))
        logo_frame.pack_propagate(False)

        # 로고 (이미지 + LISA 텍스트)
        try:
            from PIL import Image, ImageTk
            # 패키징된 환경과 개발 환경 모두에서 작동하도록 수정
            icon_path = resource_path('design/lisa.ico')
            if os.path.exists(icon_path):
                logo_img = Image.open(icon_path)
            else:
                # 패키징된 환경에서는 절대 경로로 시도
                import sys
                if getattr(sys, 'frozen', False):
                    # PyInstaller로 패키징된 경우
                    base_path = sys._MEIPASS
                    icon_path = os.path.join(base_path, 'design', 'lisa.ico')
                    if os.path.exists(icon_path):
                        logo_img = Image.open(icon_path)
                    else:
                        raise FileNotFoundError("아이콘 파일을 찾을 수 없습니다")
                else:
                    raise FileNotFoundError("아이콘 파일을 찾을 수 없습니다")
            
            logo_img = logo_img.resize((48, 48), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_img)
            logo_icon = tk.Label(logo_frame, image=logo_photo, bg="#1E293B")
            logo_icon.image = logo_photo  # 참조 유지
            logo_icon.pack()
        except Exception as e:
            print(f"로고 이미지 로드 실패: {e}")
            logo_icon = tk.Label(logo_frame, text="LISA", fg="white", bg="#1E293B")
            logo_icon.pack()
        
        logo_text = tk.Label(logo_frame, text="LISA", fg="white", bg="#1E293B")
        logo_text.pack()

        # 메뉴 정의 (아이콘, 텍스트, Frame key)
        self.menu_defs = [
            ("🔄", "Renewal 관리", "renewal"),
            ("👥", "고객사 관리", "customer"),
            ("📋", "견적서", "quote"),
            ("💰", "매입/매출", "purchase_sales"),
            ("📊", "대시보드", "dashboard"),
            ("📜", "인증서", "cert")
        ]

        # 서브메뉴 정의
        self.submenu_defs = {
            "dashboard": [
                ("📊", "팀별", "dashboard_team"),
                ("👤", "담당자별", "dashboard_person"),
                ("🏭", "제조사별", "dashboard_manufacturer"),
                ("🔧", "Foundry", "dashboard_foundry")
            ]
        }

        self.menu_buttons = {}
        self.submenu_frames = {}  # 서브메뉴 프레임들을 저장
        self.menu_frames = {}  # 메뉴 프레임들을 저장
        
        for icon, text, key in self.menu_defs:
            # 메뉴 버튼 프레임
            btn_frame = tk.Frame(self.sidebar, bg="#1E293B", height=50)
            btn_frame.pack(fill='x', pady=2, padx=4)
            btn_frame.pack_propagate(False)
            self.menu_frames[key] = btn_frame  # 메뉴 프레임 저장

            # 메뉴 버튼
            btn = tk.Button(
                btn_frame,
                text=f"  {icon}  {text}",
                anchor="w",
                bg="#1E293B",
                fg="#94A3B8",  # 연한 회색
                relief="flat",
                bd=0,
                activebackground="#334155",
                activeforeground="white",
                command=lambda k=key: self.handle_menu_click(k)
            )
            btn.pack(fill="both", expand=True, padx=4, pady=5)

            # hover 효과
            btn.bind('<Enter>', lambda e, b=btn: self.on_menu_hover(b, True))
            btn.bind('<Leave>', lambda e, b=btn: self.on_menu_hover(b, False))

            self.menu_buttons[key] = btn
            
            # 서브메뉴가 있는 경우 서브메뉴 프레임 생성
            if key in self.submenu_defs:
                self.submenu_frames[key] = tk.Frame(self.sidebar, bg="#1E293B")
                # 서브메뉴는 처음에는 숨김
                self.submenu_frames[key].pack_forget()
                
                # 서브메뉴 버튼들 생성
                for sub_icon, sub_text, sub_key in self.submenu_defs[key]:
                    sub_btn_frame = tk.Frame(self.submenu_frames[key], bg="#1E293B", height=40)
                    sub_btn_frame.pack(fill='x', pady=1, padx=8)
                    sub_btn_frame.pack_propagate(False)
                    
                    sub_btn = tk.Button(
                        sub_btn_frame,
                        text=f"    {sub_icon}  {sub_text}",
                        anchor="w",
                        bg="#1E293B",
                        fg="#64748B",  # 더 연한 회색
                        relief="flat",
                        bd=0,
                        activebackground="#334155",
                        activeforeground="white",
                        command=lambda k=sub_key: self.handle_submenu_click(k)
                    )
                    sub_btn.pack(fill="both", expand=True, padx=4, pady=3)
                    
                    # 서브메뉴 hover 효과
                    sub_btn.bind('<Enter>', lambda e, b=sub_btn: self.on_submenu_hover(b, True))
                    sub_btn.bind('<Leave>', lambda e, b=sub_btn: self.on_submenu_hover(b, False))
                    
                    self.menu_buttons[sub_key] = sub_btn

        # --- main_frame ---
        self.main_frame = tk.Frame(self.root, bg="#F7F9FB")
        self.main_frame.pack(side="right", fill="both", expand=True)
        
        # 메인 프레임이 제대로 표시되도록 강제 업데이트
        self.main_frame.update_idletasks()

        # 각 기능별 Frame 생성 및 main_frame에 배치
        self.frames = {}
        self.frames["renewal"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["customer"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["quote"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["purchase_sales"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["dashboard"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["cert"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        
        # 서브메뉴용 프레임들 추가
        self.frames["dashboard_team"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["dashboard_person"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["dashboard_manufacturer"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        self.frames["dashboard_foundry"] = tk.Frame(self.main_frame, bg="#F7F9FB")
        
        # 초기 데이터 로드 (즉시 실행)
        initial_data = self.data_loader.load_current_month_data()
        
        # 각 Frame에 기존 기능별 UI 삽입 (데이터 로더 전달)
        self.renewal_gui = RenewalMailerGUI(self.frames["renewal"], username=self.username, data_loader=self.data_loader)
        self.quote_gui = QuoteInfoGUI(self.frames["quote"], username=self.username, data_loader=self.data_loader)
        self.customer_gui = CustomerInfoGUI(self.frames["customer"], username=self.username, data_loader=self.data_loader, quote_gui=self.quote_gui)
        self.purchase_sales_gui = PurchaseSalesInfo(self.frames["purchase_sales"], username=self.username, data_loader=self.data_loader)
        
        # 대시보드 UI 생성
        from modules.dashboard.interactive_dashboard import InteractiveDashboard
        self.dashboard_gui = InteractiveDashboard(self.frames["dashboard"], username=self.username, data_loader=self.data_loader)
        
        # 서브메뉴용 기본 UI 생성
        self.create_submenu_ui()
        
        # 팀별 대시보드 UI 생성 (메인 GUI에 통합)
        try:
            from modules.dashboard.team_performance_dashboard import create_team_performance_dashboard
            self.team_dashboard_gui = create_team_performance_dashboard(self.frames["dashboard_team"])
        except Exception as e:
            # 기본 팀별 대시보드 생성
            self.create_basic_team_dashboard()
        
        from certificates.license_certificate_info import open_dialog
        open_dialog(self.frames["cert"], username=self.username)

        # Renewal 관리 화면을 기본으로 설정
        self.current_frame = "renewal"
        
        # 메인 프레임 강제 업데이트
        self.root.update_idletasks()
        
        self.show_frame("renewal")  # 명시적으로 Renewal 관리 화면 표시
        
        # 백그라운드 로딩 시작
        # 로딩 창을 먼저 표시한 후 백그라운드 로딩 시작
        self.root.after(100, self.start_background_loading)

    def on_menu_hover(self, button, entering):
        """메뉴 hover 효과"""
        if entering:
            button.configure(bg="#334155", fg="white")
        else:
            # 현재 선택된 메뉴가 아닌 경우에만 기본 색상으로
            current_key = None
            for key, btn in self.menu_buttons.items():
                if btn == button:
                    current_key = key
                    break
            if current_key and self.current_frame != current_key:
                button.configure(bg="#1E293B", fg="#94A3B8")
    
    def on_submenu_hover(self, button, entering):
        """서브메뉴 hover 효과"""
        if entering:
            button.configure(bg="#334155", fg="white")
        else:
            button.configure(bg="#1E293B", fg="#64748B")
    
    def handle_menu_click(self, menu_key):
        """메뉴 클릭 처리"""
        if menu_key == "dashboard":
            # 대시보드 메뉴 클릭 시 서브메뉴 표시하고 대시보드 화면도 표시
            if menu_key in self.submenu_frames:
                if not self.submenu_frames[menu_key].winfo_ismapped():
                    # 서브메뉴가 숨겨져 있으면 보이기
                    self.submenu_frames[menu_key].pack_forget()
                    self.submenu_frames[menu_key].pack(fill='x', pady=0, padx=0, after=self.menu_frames[menu_key])
            # 대시보드 화면 표시
            self.show_frame(menu_key)
        elif menu_key in self.submenu_defs:
            # 다른 서브메뉴가 있는 메뉴의 경우 기존 토글 동작
            self.toggle_submenu(menu_key)
        else:
            # 일반 메뉴의 경우 화면만 표시
            # 대시보드 외의 메뉴 클릭 시 대시보드 서브메뉴 접기
            if "dashboard" in self.submenu_frames and self.submenu_frames["dashboard"].winfo_ismapped():
                self.submenu_frames["dashboard"].pack_forget()
            self.show_frame(menu_key)
    
    def handle_submenu_click(self, submenu_key):
        """서브메뉴 클릭 처리"""
        # 서브메뉴 클릭 시 해당 화면 표시
        self.show_frame(submenu_key)
    
    def toggle_submenu(self, menu_key):
        """서브메뉴 토글"""
        if menu_key in self.submenu_frames:
            if self.submenu_frames[menu_key].winfo_ismapped():
                # 서브메뉴가 보이면 숨기기
                self.submenu_frames[menu_key].pack_forget()
            else:
                # 서브메뉴가 숨겨져 있으면 해당 메뉴 바로 다음에 보이기
                # 먼저 서브메뉴를 제거하고 다시 삽입
                self.submenu_frames[menu_key].pack_forget()
                # 해당 메뉴 프레임 다음에 서브메뉴 삽입
                self.submenu_frames[menu_key].pack(fill='x', pady=0, padx=0, after=self.menu_frames[menu_key])
    
    def create_submenu_ui(self):
        """서브메뉴용 기본 UI 생성"""
        # 팀별 대시보드는 별도로 생성되므로 여기서는 제거
        
        # 담당자별 대시보드
        person_frame = self.frames["dashboard_person"]
        tk.Label(person_frame, text="담당자별 대시보드", 
                bg="#F7F9FB", fg="#1E293B").pack(pady=50)
        tk.Label(person_frame, text="담당자별 통계 및 분석 기능이 구현될 예정입니다.", 
                bg="#F7F9FB", fg="#64748B").pack()
        
        # 제조사별 대시보드
        manufacturer_frame = self.frames["dashboard_manufacturer"]
        tk.Label(manufacturer_frame, text="제조사별 대시보드", 
                bg="#F7F9FB", fg="#1E293B").pack(pady=50)
        tk.Label(manufacturer_frame, text="제조사별 통계 및 분석 기능이 구현될 예정입니다.", 
                bg="#F7F9FB", fg="#64748B").pack()
        
        # Foundry 대시보드
        foundry_frame = self.frames["dashboard_foundry"]
        self.foundry_dashboard = create_foundry_dashboard(foundry_frame)

    def show_frame(self, key):
        self.current_frame = key
        for k, f in self.frames.items():
            if k == key:
                f.pack(fill="both", expand=True)
                # 선택된 메뉴 스타일
                self.menu_buttons[k].configure(bg="#3B82F6", fg="white")
                # 프레임 강제 업데이트
                f.update_idletasks()
            else:
                f.pack_forget()
                # 선택되지 않은 메뉴 스타일
                self.menu_buttons[k].configure(bg="#1E293B", fg="#94A3B8")
        
        # 메인 프레임 강제 업데이트
        self.main_frame.update_idletasks()

    def start_background_loading(self):
        """백그라운드 데이터 로딩 시작"""
        
        # 로딩 인디케이터 생성 및 표시
        from ui.loading_indicator import LoadingIndicator
        self.loading_indicator = LoadingIndicator(self.root)
        self.loading_indicator.show("데이터를 로딩 중입니다...")
        
        # 로딩 창이 표시되도록 강제 업데이트
        self.root.update()
        
        # 백그라운드에서 데이터 로딩
        import threading
        self.background_loader = threading.Thread(target=self._load_data_background, daemon=True)
        self.background_loader.start()
        
        # 로딩 모니터링 시작
        self._monitor_loading()
    
    def _load_data_background(self):
        """백그라운드에서 데이터 로딩"""
        try:
            print("백그라운드에서 전체 데이터를 로딩합니다...")
            # 전체 데이터 로딩
            self.data_loader.load_all_data()
            print("백그라운드 데이터 로딩 완료")
        except Exception as e:
            print(f"백그라운드 데이터 로딩 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def _monitor_loading(self):
        """로딩 상태 모니터링"""
        if hasattr(self, 'background_loader') and self.background_loader.is_alive():
            # 아직 로딩 중이면 100ms 후 다시 확인
            self.root.after(100, self._monitor_loading)
        else:
            # 로딩 완료
            if hasattr(self, 'loading_indicator'):
                self.loading_indicator.hide()
            self.on_data_loading_complete()
    
    def on_data_loading_complete(self):
        """데이터 로딩 완료 시 호출되는 콜백"""
        print("전체 데이터 로딩이 완료되어 각 GUI 모듈을 업데이트합니다.")
        
        # 각 GUI 모듈에 전체 데이터가 로드되었음을 알림
        if hasattr(self, 'renewal_gui'):
            print("Renewal 관리 탭 업데이트 중...")
            self.renewal_gui.on_full_data_loaded()
        if hasattr(self, 'customer_gui'):
            print("고객사 관리 탭 업데이트 중...")
            self.customer_gui.on_full_data_loaded()
        if hasattr(self, 'quote_gui'):
            print("견적서 탭 업데이트 중...")
            self.quote_gui.on_full_data_loaded()
        if hasattr(self, 'purchase_sales_gui'):
            print("매입/매출 탭 업데이트 중...")
            self.purchase_sales_gui.on_full_data_loaded()
        
        print("모든 데이터 로딩이 완료되었습니다.")
        
        # 만료일 알람 팝업 표시 (3초 후)
        self.root.after(3000, self._show_expiry_alarm)

    def create_basic_team_dashboard(self):
        """기본 팀별 대시보드 생성"""
        team_frame = self.frames["dashboard_team"]
        
        # 기존 위젯 제거
        for widget in team_frame.winfo_children():
            widget.destroy()
        
        # 기본 UI 생성
        title_label = tk.Label(team_frame, text="📊 팀별 성과 대시보드", 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(pady=50)
        
        status_label = tk.Label(team_frame, text="팀별 성과 대시보드가 준비 중입니다.\n곧 정상적으로 표시될 예정입니다.", 
                               fg="#6B7280", bg="#F7F9FB")
        status_label.pack(pady=20)
        
        # 데이터 로딩 상태 표시
        self.team_loading_label = tk.Label(team_frame, text="🔄 데이터 로딩 중...", 
                                          fg="#6B7280", bg="#F7F9FB")
        self.team_loading_label.pack(pady=10)
    
    def _show_expiry_alarm(self):
        """만료일 알람 팝업 표시"""
        def go_to_renewal():
            self.show_frame("renewal")
        
        show_expiry_alarm(
            self.root,
            self.data_loader,
            self.username,
            on_detail_click=go_to_renewal
        )

    def on_closing(self):
        """프로그램 종료 시 정리 작업"""
        try:
            print("LISA 프로그램 종료 중...")
            
            # 백그라운드 로더가 있으면 정리
            if hasattr(self, 'background_loader') and self.background_loader:
                if hasattr(self.background_loader, 'loading_indicator'):
                    self.background_loader.loading_indicator.hide()
                print("백그라운드 로더 정리 완료")
            
            # 데이터 로더 정리
            if hasattr(self, 'data_loader') and self.data_loader:
                # 백그라운드 로딩 중단
                if hasattr(self.data_loader, 'stop_background_loading'):
                    self.data_loader.stop_background_loading()
                    print("백그라운드 로딩 중단")
                
                # 캐시 정리
                from core.data_loader import clear_cache
                clear_cache()
                print("캐시 정리 완료")
            
            # 모든 스레드 종료 대기 (최대 5초로 증가)
            import threading
            import time
            
            # 활성 스레드 확인
            active_threads = [t for t in threading.enumerate() if t != threading.main_thread()]
            if active_threads:
                print(f"활성 스레드 {len(active_threads)}개 종료 대기 중...")
                
                # 각 스레드에 종료 신호 전송
                for thread in active_threads:
                    if hasattr(thread, '_stop'):
                        thread._stop()
                    elif hasattr(thread, 'stop'):
                        thread.stop()
                
                # 최대 5초 대기 (3초에서 증가)
                start_time = time.time()
                while active_threads and (time.time() - start_time) < 5:
                    active_threads = [t for t in threading.enumerate() if t != threading.main_thread() and t.is_alive()]
                    time.sleep(0.1)
                
                if active_threads:
                    print(f"강제 종료: {len(active_threads)}개 스레드")
                    # 남은 스레드 정보 출력
                    for i, thread in enumerate(active_threads):
                        print(f"  스레드 {i+1}: {thread.name} (daemon: {thread.daemon})")
                else:
                    print("모든 스레드 정상 종료")
            
            print("LISA 프로그램 정리 완료")
            
        except Exception as e:
            print(f"프로그램 종료 중 오류: {e}")
        finally:
            # 창 종료
            self.root.destroy()

    def run(self):
        pass  # mainloop는 여기서 실행하지 않음

if __name__ == '__main__':
    root = tk.Tk()
    try:
        # 패키징된 환경과 개발 환경 모두에서 작동하도록 수정
        icon_path = resource_path('design/lisa.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
        else:
            # 패키징된 환경에서는 절대 경로로 시도
            import sys
            if getattr(sys, 'frozen', False):
                # PyInstaller로 패키징된 경우
                base_path = sys._MEIPASS
                icon_path = os.path.join(base_path, 'design', 'lisa.ico')
                if os.path.exists(icon_path):
                    root.iconbitmap(icon_path)
    except Exception as e:
        print(f"아이콘 로드 실패: {e}")
        pass  # 아이콘 파일이 없거나 오류가 나도 무시
    
    # 모던한 ttk 스타일 설정
    try:
        pass
    except Exception:
        pass
    
    style = ttk.Style()
    style.theme_use('default')
    
    
    # 버튼 스타일
    style.configure('TButton', 
                   padding=(10, 6),
                   background="#3B82F6",
                   foreground="white")
    style.map('TButton',
              background=[('active', '#2563EB'), ('pressed', '#1D4ED8')],
              foreground=[('active', 'white'), ('pressed', 'white')])
    
    # 라벨 스타일
    style.configure('TLabel', font=("맑은 고딕", 10))
    
    # 엔트리 스타일
    style.configure('TEntry', 
                   font=("맑은 고딕", 10),
                   padding=(6, 4),
                   fieldbackground="white",
                   borderwidth=1,
                   relief="solid")
    
    # Combobox 스타일
    style.configure('TCombobox', 
                   font=("맑은 고딕", 10),
                   padding=(6, 4),
                   fieldbackground="white",
                   borderwidth=1,
                   relief="solid",
                   arrowcolor="#6B7280",
                   background="white")
    style.map('TCombobox',
              fieldbackground=[('readonly', 'white'), ('disabled', '#F3F4F6')],
              selectbackground=[('readonly', '#DBEAFE')],
              selectforeground=[('readonly', '#1E40AF')])
    
    # Treeview 헤더 스타일
    style.configure('Treeview.Heading', 
                   background="#F1F5F9",
                   foreground="#374151",
                   padding=(6, 4))
    style.map('Treeview.Heading',
              background=[('active', '#E2E8F0')])
    
    # Treeview 본문 스타일
    style.configure('Treeview', 
                   font=("맑은 고딕", 9), 
                   rowheight=28,
                   background="white",
                   foreground="#374151",
                   fieldbackground="white",
                   borderwidth=0)
    style.map('Treeview', 
              background=[('selected', '#DBEAFE')],
              foreground=[('selected', '#1E40AF')])
    
    # Treeview 레이아웃 설정
    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
    
    # 프레임 스타일
    style.configure('Card.TFrame', 
                   background="white",
                   relief="flat",
                   borderwidth=0)
    
    # 라운드 처리된 프레임을 위한 커스텀 스타일
    style.configure('Rounded.TFrame', 
                   background="white",
                   relief="flat",
                   borderwidth=0)
    
    # 라디오버튼 스타일
    style.configure('TRadiobutton', 
                   font=("맑은 고딕", 11),
                   background="#F7F9FB")
    
    # 체크박스 스타일
    style.configure('TCheckbutton', 
                   font=("맑은 고딕", 11),
                   background="#F7F9FB")
    
    root.withdraw()  # 메인 창 숨김
    login_win = tk.Toplevel(root)
    login_win.lift()  # 창을 최상위로
    login_win.focus_force()  # 포커스 강제 설정
    login_win.attributes('-topmost', True)  # 항상 최상위
    login_win.after(100, lambda: login_win.attributes('-topmost', False))  # 100ms 후 해제
    
    # 창 크기와 위치 강제 설정
    login_win.geometry("400x300")
    login_win.resizable(False, False)
    
    # 창이 보이도록 강제 업데이트
    login_win.update()
    login_win.deiconify()
    
    try:
        LoginWindow(login_win, root)
        root.mainloop()
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        input("엔터를 눌러서 종료하세요...")
