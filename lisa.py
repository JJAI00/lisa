
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import traceback
import hashlib
import gspread
from google.oauth2.service_account import Credentials

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
    from config import KEY_FILE, SCOPES, SHEET_IDS, resource_path
    LOGIN_SHEET_ID = SHEET_IDS.get('login_list')
except ImportError:
    # Fallback config
    def resource_path(relative_path):
        if hasattr(sys, 'frozen') and hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    KEY_FILE = resource_path('resources/keys/renewal-bot-463605-b8f41de8fbdb.json')
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    LOGIN_SHEET_ID = '1KEcLicMV6p-RSXl2kg1AHyD2CmU8wABeXDlMTk0gqZY'

# --- UI 상수 ---
COLOR_SIDEBAR_BG = "#2c3e50"  # Midnight Blue
COLOR_SIDEBAR_FG = "#ecf0f1"  # Clouds
COLOR_MENU_HOVER = "#34495e"  # Wet Asphalt
COLOR_MENU_ACTIVE = "#3498db" # Peter River (밝은 파랑)
COLOR_MAIN_BG    = "#ecf0f1"

# --- 유틸리티 ---
def hash_password(pw):
    """SHA-256 해시화"""
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("LISA Login")
        self.root.geometry("400x380")
        self.root.configure(bg=COLOR_MAIN_BG)
        # 윈도우 크기 조절 불가
        self.root.resizable(False, False)
        
        # 화면 중앙 배치
        self.center_window(400, 380)
        
        # 로고 영역 (텍스트)
        tk.Label(root, text="LISA", font=("Arial", 36, "bold"), bg=COLOR_MAIN_BG, fg=COLOR_SIDEBAR_BG).pack(pady=(50, 10))
        tk.Label(root, text="License & Sales Automation", font=("Arial", 10), bg=COLOR_MAIN_BG, fg="#7f8c8d").pack(pady=(0, 30))
        
        # 입력 폼 프레임
        frame = tk.Frame(root, bg=COLOR_MAIN_BG)
        frame.pack(pady=10)
        
        # Username
        tk.Label(frame, text="Username", font=("Arial", 10, "bold"), bg=COLOR_MAIN_BG, fg=COLOR_SIDEBAR_BG).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_id = tk.Entry(frame, font=("Arial", 12), width=25, bd=1, relief="solid")
        self.entry_id.insert(0, "김한경") # 편의상 기본값
        self.entry_id.grid(row=1, column=0, padx=5, pady=(0, 10), ipady=3)
        self.entry_id.bind('<Return>', lambda e: self.entry_pw.focus())
        
        # Password
        tk.Label(frame, text="Password", font=("Arial", 10, "bold"), bg=COLOR_MAIN_BG, fg=COLOR_SIDEBAR_BG).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_pw = tk.Entry(frame, font=("Arial", 12), width=25, bd=1, relief="solid", show="*")
        self.entry_pw.grid(row=3, column=0, padx=5, pady=(0, 10), ipady=3)
        self.entry_pw.bind('<Return>', lambda e: self.login())
        
        # Login Button
        btn = tk.Button(root, text="LOGIN", command=self.login, width=25, 
                        bg=COLOR_SIDEBAR_BG, fg="white", font=("Arial", 11, "bold"),
                        bd=0, relief="flat", cursor="hand2")
        btn.pack(pady=20, ipady=8)
        
        # 상태 메시지
        self.status_lbl = tk.Label(root, text="", bg=COLOR_MAIN_BG, fg="red", font=("Arial", 9))
        self.status_lbl.pack()
        
        self.entry_id.focus_set()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def login(self):
        username = self.entry_id.get().strip()
        password = self.entry_pw.get().strip()
        
        if not username or not password:
            self.status_lbl.config(text="아이디와 비밀번호를 모두 입력하세요.")
            return

        self.status_lbl.config(text="로그인 중...", fg="blue")
        self.root.update()

        # 인증 시도
        try:
            if self.authenticate_user(username, password):
                # 성공 시 메인 앱으로
                logger.info(f"Login success: {username}")
                self.launch_main_app(username)
            else:
                self.status_lbl.config(text="로그인 실패: 아이디 또는 비밀번호가 틀렸습니다.", fg="red")
                messagebox.showerror("로그인 실패", "아이디 또는 비밀번호가 올바르지 않습니다.")
        except Exception as e:
            self.status_lbl.config(text=f"오류 발생: {str(e)}", fg="red")
            logger.error(f"Login error: {traceback.format_exc()}")
            messagebox.showerror("오류", f"로그인 중 오류가 발생했습니다.\n{e}")

    def authenticate_user(self, username, password):
        """구글 시트 연동하여 사용자 인증"""
        # 개발용 백도어 (필요 시 주석 해제)
        # if username == "admin" and password == "0000": return True
        
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(LOGIN_SHEET_ID)
            ws = sh.sheet1 # 첫 번째 시트 가정
            
            # 모든 데이터 가져오기 (헤더 제외)
            data = ws.get_all_values()
            if not data: return False
            
            # 헤더가 있을 수 있으므로 2번째 줄부터 체크하거나, 내용을 보고 판단
            # 가정: A열=이름, B열=해시비번
            
            input_hash = hash_password(password)
            
            for row in data:
                if len(row) < 2: continue
                db_user = row[0].strip()
                db_pw   = row[1].strip()
                
                if db_user == username:
                    if db_pw == input_hash:
                        return True
                    else:
                        logger.warning(f"Password mismatch for user: {username}")
                        return False
            
            logger.warning(f"User not found: {username}")
            return False
            
        except Exception as e:
            logger.error(f"Authentication failed: {e}")
            raise e

    def launch_main_app(self, username):
        # 데이터 로더 초기화
        try:
            from core.data_loader import DataLoader
            data_loader = DataLoader(KEY_FILE, SCOPES)
            logger.info("DataLoader initialized")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로더 초기화 실패:\n{e}")
            return

        # 메인 앱 전환
        for widget in self.root.winfo_children():
            widget.destroy()
            
        MainApp(self.root, username, data_loader)


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
        self.sidebar = tk.Frame(self.root, bg=COLOR_SIDEBAR_BG, width=220)
        self.sidebar.pack(side='left', fill='y')
        self.sidebar.pack_propagate(False) # 너비 고정
        
        # 로고 영역
        logo_frame = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG, height=100)
        logo_frame.pack(fill='x', pady=30)
        
        # 로고 이미지 로드 시도
        try:
            from PIL import Image, ImageTk
            logo_path = resource_path("resources/images/logo.png") # 경로 추정
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((150, 50), Image.Resampling.LANCZOS) # 비율에 맞게 조절 필요
                self.logo_img = ImageTk.PhotoImage(img) # 참조 유지
                tk.Label(logo_frame, image=self.logo_img, bg=COLOR_SIDEBAR_BG).pack()
            else:
                tk.Label(logo_frame, text="LISA", font=("Arial", 28, "bold"), 
                         bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG).pack()
        except Exception:
            tk.Label(logo_frame, text="LISA", font=("Arial", 28, "bold"), 
                     bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG).pack()
            
        tk.Label(logo_frame, text=f"v2.01\n{self.username}", font=("Arial", 9), 
                 bg=COLOR_SIDEBAR_BG, fg="#bdc3c7").pack(pady=5)
        
        # 메뉴 컨테이너
        self.menu_container = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG)
        self.menu_container.pack(fill='both', expand=True, pady=20)
        
        # 2. 메인 컨텐츠 (오른쪽)
        self.main_area = tk.Frame(self.root, bg=COLOR_MAIN_BG)
        self.main_area.pack(side='right', fill='both', expand=True)
        
        # 페이지 저장소
        self.pages = {}
        self.buttons = {}

    def _load_menus(self):
        # 메뉴 정의
        menus = [
            ("Renewal 관리", "modules.renewal.renewal_preview", "RenewalMailerGUI", "class", True),
            ("고객사 관리",   "modules.customer.customer_info",   "CustomerInfoGUI", "class", False),
            ("견적서",       "modules.quote.quote_info",         "QuoteInfoGUI", "class", True),
            ("매입/매출",     "modules.purchase.purchase_sales_info", "PurchaseSalesInfo", "class", False),
            ("대시보드",     "modules.dashboard.dashboard_foundry", "create_foundry_dashboard", "func", False),
            ("인증서",       "certificates.license_certificate_info", "open_dialog", "func", False)
        ]
        
        for title, mod_path, obj_name, obj_type, use_dl in menus:
            self._add_menu_button(title)
            self._preload_page(title, mod_path, obj_name, obj_type, use_dl)

    def _add_menu_button(self, title):
        btn = tk.Button(self.menu_container, text=title, font=("Arial", 11),
                        bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG,
                        activebackground=COLOR_MENU_HOVER, activeforeground=COLOR_SIDEBAR_FG,
                        bd=0, relief="flat", anchor="w", padx=25, pady=12,
                        command=lambda t=title: self.switch_page(t))
        btn.pack(fill='x', pady=2)
        self.buttons[title] = btn
        
    def _preload_page(self, title, mod_path, obj_name, obj_type, use_dl):
        page_frame = tk.Frame(self.main_area, bg=COLOR_MAIN_BG)
        # 로딩 라벨 제거 (깔끔하게 빈 화면으로 시작)
        self.pages[title] = page_frame
        
        try:
            module = __import__(mod_path, fromlist=[obj_name])
            obj = getattr(module, obj_name)
            
            if obj_type == "func":
                obj(page_frame)
            else:
                if use_dl:
                    try:
                        obj(page_frame, username=self.username, data_loader=self.data_loader)
                    except TypeError:
                         obj(page_frame, username=self.username)
                else:
                    obj(page_frame, username=self.username)
                    
            logger.info(f"Loaded page: {title}")
            
        except Exception as e:
            for child in page_frame.winfo_children(): child.destroy()
            err_msg = f"{title} 로드 실패\n\n{e}"
            tk.Label(page_frame, text=err_msg, fg="red", bg=COLOR_MAIN_BG).pack(expand=True)
            logger.error(f"Failed to load {title}: {traceback.format_exc()}")

    def switch_page(self, title):
        for frame in self.pages.values():
            frame.pack_forget()
            
        if title in self.pages:
            self.pages[title].pack(fill='both', expand=True)
            
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
