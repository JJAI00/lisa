
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import traceback

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 로깅 설정
try:
    from integrations.lisa_logging import setup_logger
    logger = setup_logger('main')
except ImportError:
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger('main')


# 설정 파일 로드
try:
    from config import KEY_FILE, SCOPES
except ImportError:
    KEY_FILE = 'resources/keys/service_account.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# --- UI 상수 ---
COLOR_SIDEBAR_BG = "#1e2129"  # 더 어두운 남색
COLOR_SIDEBAR_FG = "#ffffff"
COLOR_MENU_HOVER = "#2c313c"
COLOR_MENU_ACTIVE = "#3a7bd5" # 밝은 파랑
COLOR_MAIN_BG    = "#f5f6fa"

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("LISA Login")
        self.root.geometry("300x200")
        
        # 화면 중앙 배치
        self.center_window(300, 200)
        
        tk.Label(root, text="LISA v2.01", font=("Arial", 16, "bold")).pack(pady=20)
        
        frame = tk.Frame(root)
        frame.pack(pady=10)
        
        tk.Label(frame, text="사용자 이름:").pack(side='left', padx=5)
        self.entry = tk.Entry(frame)
        self.entry.insert(0, "김한경") # 기본값
        self.entry.pack(side='left', padx=5)
        self.entry.bind('<Return>', lambda e: self.login())
        self.entry.focus_set()
        
        tk.Button(root, text="로그인", command=self.login, width=20, bg="#007bff", fg="white").pack(pady=20)
        
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def login(self):
        username = self.entry.get()
        if not username:
            messagebox.showwarning("경고", "이름을 입력하세요")
            return
            
        # 데이터 로더 초기화 (로그인 시점)
        try:
            from core.data_loader import DataLoader
            self.data_loader = DataLoader(KEY_FILE, SCOPES)
            logger.info("DataLoader initialized")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로더 초기화 실패:\n{e}")
            logger.error(f"DataLoader init failed: {traceback.format_exc()}")
            return

        # 메인 앱 전환
        for widget in self.root.winfo_children():
            widget.destroy()
            
        MainApp(self.root, username, self.data_loader)

class MainApp:
    def __init__(self, root, username, data_loader):
        self.root = root
        self.root.title(f"LISA v2.01 - {username}")
        self.root.geometry("1400x900")
        self.root.configure(bg=COLOR_MAIN_BG)
        
        # 화면 중앙 배치
        self.center_window(1400, 900)
        
        self.username = username
        self.data_loader = data_loader
        
        # 레이아웃 구성
        self._setup_layout()
        
        # 메뉴 로드
        self._load_menus()
        
        # 초기 화면 (첫 번째 메뉴)
        self.switch_page("Renewal 관리")

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_layout(self):
        # 1. 사이드바 (왼쪽)
        self.sidebar = tk.Frame(self.root, bg=COLOR_SIDEBAR_BG, width=250)
        self.sidebar.pack(side='left', fill='y')
        self.sidebar.pack_propagate(False) # 너비 고정
        
        # 로고 영역
        logo_frame = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG, height=80)
        logo_frame.pack(fill='x', pady=20)
        tk.Label(logo_frame, text="LISA", font=("Arial", 24, "bold"), 
                 bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG).pack()
        tk.Label(logo_frame, text=f"v2.01 - {self.username}", font=("Arial", 10), 
                 bg=COLOR_SIDEBAR_BG, fg="#888888").pack()
        
        # 메뉴 컨테이너
        self.menu_container = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG)
        self.menu_container.pack(fill='both', expand=True, pady=10)
        
        # 2. 메인 컨텐츠 (오른쪽)
        self.main_area = tk.Frame(self.root, bg=COLOR_MAIN_BG)
        self.main_area.pack(side='right', fill='both', expand=True)
        
        # 페이지 저장소
        self.pages = {}
        self.buttons = {}

    def _load_menus(self):
        # 메뉴 정의 (이름, 아이콘(생략), 모듈경로, 클래스명, DataLoader필요여부)
        menus = [
            ("Renewal 관리", "modules.renewal.renewal_preview", "RenewalMailerGUI", True),
            ("고객사 관리",   "modules.customer.customer_info",   "CustomerInfoGUI", False),
            ("견적서",       "modules.quote.quote_info",         "QuoteInfoGUI", True),
            ("매입/매출",     "modules.purchase.purchase_sales_info", "PurchaseSalesInfo", False),
            ("대시보드",     "modules.dashboard.dashboard_foundry", "create_foundry_dashboard", False), # 함수형
            ("인증서",       "certificates.license_certificate_info", "LicenseCertificateInfo", False)
        ]
        
        for title, mod_path, cls_name, use_dl in menus:
            self._add_menu_button(title)
            # 페이지 지연 로드 (클릭 시 로드하거나, 여기서 미리 로드하거나)
            # 여기서는 미리 로드하지만 에러 시 빈 프레임 유지
            self._preload_page(title, mod_path, cls_name, use_dl)

    def _add_menu_button(self, title):
        btn = tk.Button(self.menu_container, text=title, font=("Arial", 11),
                        bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG,
                        activebackground=COLOR_MENU_HOVER, activeforeground=COLOR_SIDEBAR_FG,
                        bd=0, relief="flat", anchor="w", padx=20, pady=12,
                        command=lambda t=title: self.switch_page(t))
        btn.pack(fill='x')
        self.buttons[title] = btn
        
    def _preload_page(self, title, mod_path, cls_name, use_dl):
        # 컨테이너 프레임
        page_frame = tk.Frame(self.main_area, bg=COLOR_MAIN_BG)
        
        # 로딩 라벨 (나중에 덮어씌워짐)
        loading_lbl = tk.Label(page_frame, text=f"{title} 로딩 중...", bg=COLOR_MAIN_BG)
        loading_lbl.pack(expand=True)
        
        self.pages[title] = page_frame
        
        # 실제 로딩 (즉시 실행)
        try:
            # 대시보드는 함수형일 수 있음
            if cls_name == "create_foundry_dashboard":
                module = __import__(mod_path, fromlist=[cls_name])
                func = getattr(module, cls_name)
                # 기존 컨텐츠 제거
                for child in page_frame.winfo_children(): child.destroy()
                func(page_frame)
            else:
                module = __import__(mod_path, fromlist=[cls_name])
                cls = getattr(module, cls_name)
                # 기존 컨텐츠 제거
                for child in page_frame.winfo_children(): child.destroy()
                
                if use_dl:
                    try:
                        cls(page_frame, username=self.username, data_loader=self.data_loader)
                    except TypeError:
                        cls(page_frame, username=self.username)
                else:
                    cls(page_frame, username=self.username)
                    
            logger.info(f"Loaded page: {title}")
            
        except Exception as e:
            for child in page_frame.winfo_children(): child.destroy()
            err_msg = f"{title} 로드 실패\n\n{e}\n\n{traceback.format_exc()}"
            tk.Label(page_frame, text=err_msg, fg="red", bg=COLOR_MAIN_BG, justify="left").pack(padx=20, pady=20)
            logger.error(f"Failed to load {title}: {traceback.format_exc()}")

    def switch_page(self, title):
        # 1. 모든 페이지 숨김
        for frame in self.pages.values():
            frame.pack_forget()
            
        # 2. 선택된 페이지 표시
        if title in self.pages:
            self.pages[title].pack(fill='both', expand=True)
            
        # 3. 버튼 스타일 업데이트
        for t, btn in self.buttons.items():
            if t == title:
                btn.configure(bg=COLOR_MENU_ACTIVE, font=("Arial", 11, "bold"))
            else:
                btn.configure(bg=COLOR_SIDEBAR_BG, font=("Arial", 11))

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = LoginWindow(root)
        root.mainloop()
    except Exception as e:
        with open("crash_log.txt", "w") as f:
            f.write(traceback.format_exc())
        messagebox.showerror("치명적 오류", f"앱 실행 중 오류가 발생했습니다.\n{e}")
