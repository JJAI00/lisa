
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

# --- UI 상수 (Previous UI Colors) ---
COLOR_SIDEBAR_BG = "#1E293B"  # Slate 900
COLOR_SIDEBAR_FG = "#ffffff"
COLOR_MENU_HOVER = "#334155"  # Slate 700
COLOR_MENU_ACTIVE = "#3B82F6" # Blue 500 (Login button color match)
COLOR_MAIN_BG    = "#F7F9FB"  # Slate 50
COLOR_TEXT_MAIN  = "#1E293B"
COLOR_TEXT_SUB   = "#64748B"

# --- 유틸리티 ---
def hash_password(pw):
    """SHA-256 해시화"""
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("LISA Login")
        self.root.geometry("500x400")
        self.root.configure(bg="#F1F5F9") # Login background
        self.root.resizable(False, False)
        
        self.center_window(500, 400)
        
        # 상단 타이틀
        tk.Label(root, text="LISA", font=("Arial", 32, "bold"), 
                 bg="#F1F5F9", fg="#3B82F6").pack(pady=(40, 5))
        tk.Label(root, text="License & Sales Automation", font=("Arial", 10), 
                 bg="#F1F5F9", fg="#64748B").pack(pady=(0, 20))
        
        # 카드 프레임 (Shadow effect simulation by borders)
        card = tk.Frame(root, bg="white", relief="solid", bd=1)
        card.pack(pady=10, padx=20, ipadx=20, ipady=20)
        
        # ID
        tk.Label(card, text="ID", font=("Arial", 10, "bold"), bg="white", fg="#334155").pack(anchor="w", pady=(0,5))
        self.entry_id = tk.Entry(card, font=("Arial", 11), width=30, bd=1, relief="solid", bg="#F8FAFC")
        self.entry_id.insert(0, "김한경")
        self.entry_id.pack(pady=(0, 15), ipady=5)
        self.entry_id.bind('<Return>', lambda e: self.entry_pw.focus())

        # PW
        tk.Label(card, text="PW", font=("Arial", 10, "bold"), bg="white", fg="#334155").pack(anchor="w", pady=(0,5))
        self.entry_pw = tk.Entry(card, font=("Arial", 11), width=30, bd=1, relief="solid", bg="#F8FAFC", show="*")
        self.entry_pw.pack(pady=(0, 20), ipady=5)
        self.entry_pw.bind('<Return>', lambda e: self.login())
        
        # Login Button
        btn = tk.Button(card, text="Login", command=self.login, width=30, 
                        bg="#3B82F6", fg="white", font=("Arial", 10, "bold"),
                        bd=0, relief="flat", cursor="hand2")
        btn.pack(ipady=8)
        
        self.status_lbl = tk.Label(root, text="", bg="#F1F5F9", fg="red", font=("Arial", 9))
        self.status_lbl.pack(pady=10)
        
        self.entry_pw.focus_set()

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
            self.status_lbl.config(text="아이디와 비밀번호를 입력하세요.")
            return

        self.status_lbl.config(text="로그인 중...", fg="#3B82F6")
        self.root.update()

        try:
            if self.authenticate_user(username, password):
                self.launch_main_app(username)
            else:
                self.status_lbl.config(text="로그인 실패: 정보가 올바르지 않습니다.", fg="red")
        except Exception as e:
            self.status_lbl.config(text=f"오류: {str(e)}", fg="red")

    def authenticate_user(self, username, password):
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            # sh = gc.open_by_key(LOGIN_SHEET_ID) # Sheet ID가 없다면 이름으로 열 수도 있음
            # 일단 ID 사용
            try:
                sh = gc.open_by_key(LOGIN_SHEET_ID)
            except:
                # ID가 틀렸거나 권한 문제 시 예외 처리
                 logger.error("Login sheet not found by ID")
                 # 개발용 백도어 for demo (실제로는 에러를 띄워야 함)
                 # if username == "김한경": return True 
                 raise Exception("Login DB access failed")

            ws = sh.sheet1
            data = ws.get_all_values()
            input_hash = hash_password(password)
            
            # 헤더 제외하고 2번째 줄부터
            for row in data[1:]:
                if len(row) < 2: continue
                # A열: 이름, B열: 해시
                if row[0].strip() == username:
                    if row[1].strip() == input_hash:
                        return True
            return False
        except Exception as e:
            # 개발용: 로컬 테스트 시 인증 패스 (실제 배포 전엔 제거 필요하지만, 복구 우선)
            # 만약 키 파일 에러라면?
            if "No such file" in str(e):
                 self.status_lbl.config(text="키 파일을 찾을 수 없습니다.")
                 return False
            raise e

    def launch_main_app(self, username):
        try:
            from core.data_loader import DataLoader
            data_loader = DataLoader(KEY_FILE, SCOPES)
        except Exception as e:
            messagebox.showerror("초기화 오류", f"데이터 로더 오류:\n{e}")
            return

        for widget in self.root.winfo_children(): widget.destroy()
        MainApp(self.root, username, data_loader)


class MainApp:
    def __init__(self, root, username, data_loader):
        self.root = root
        self.root.title(f"LISA v2.01 - {username}")
        self.root.geometry("1400x900")
        self.root.configure(bg=COLOR_MAIN_BG)
        self.username = username
        self.data_loader = data_loader
        self.center_window(1400, 900)
        self._setup_layout()
        self._load_menus()
        self._initial_load()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_layout(self):
        # Sidebar
        self.sidebar = tk.Frame(self.root, bg=COLOR_SIDEBAR_BG, width=220)
        self.sidebar.pack(side='left', fill='y')
        self.sidebar.pack_propagate(False)
        
        # Logo
        logo_frame = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG, height=100)
        logo_frame.pack(fill='x', pady=30)
        tk.Label(logo_frame, text="LISA", font=("Arial", 28, "bold"), 
                 bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG).pack()
        tk.Label(logo_frame, text=f"v2.01\n{self.username}", font=("Arial", 9), 
                 bg=COLOR_SIDEBAR_BG, fg="#94A3B8").pack(pady=5)
        
        self.menu_container = tk.Frame(self.sidebar, bg=COLOR_SIDEBAR_BG)
        self.menu_container.pack(fill='both', expand=True, pady=20)
        
        # Main Area
        self.main_area = tk.Frame(self.root, bg=COLOR_MAIN_BG)
        self.main_area.pack(side='right', fill='both', expand=True)
        
        self.pages = {}
        self.buttons = {}

    def _load_menus(self):
        menus = [
            ("Renewal 관리", "modules.renewal.renewal_preview", "RenewalMailerGUI", "class", True),
            ("고객사 관리",   "modules.customer.customer_info",   "CustomerInfoGUI", "class", False),
            ("견적서",       "modules.quote.quote_info",         "QuoteInfoGUI", "class", True),
            ("매입/매출",     "modules.purchase.purchase_sales_info", "PurchaseSalesInfo", "class", False),
            ("대시보드",     "modules.dashboard.dashboard_foundry", "create_foundry_dashboard", "func", False),
            ("인증서",       "certificates.license_certificate_info", "open_dialog", "func", False)
        ]
        
        for title, mod_path, obj_name, obj_type, use_dl in menus:
            btn = tk.Button(self.menu_container, text=title, font=("Arial", 11),
                            bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_FG,
                            activebackground=COLOR_MENU_HOVER, activeforeground=COLOR_SIDEBAR_FG,
                            bd=0, relief="flat", anchor="w", padx=25, pady=12,
                            command=lambda t=title: self.switch_page(t))
            btn.pack(fill='x', pady=2)
            self.buttons[title] = btn
            
            # Preload container
            page_frame = tk.Frame(self.main_area, bg=COLOR_MAIN_BG)
            self.pages[title] = {
                "frame": page_frame,
                "mod_path": mod_path,
                "obj_name": obj_name,
                "obj_type": obj_type,
                "use_dl": use_dl,
                "loaded": False
            }

    def _initial_load(self):
        # 첫 페이지 로드
        self.switch_page("Renewal 관리")

    def switch_page(self, title):
        # UI 업데이트
        for t, btn in self.buttons.items():
            if t == title:
                btn.configure(bg=COLOR_MENU_ACTIVE, font=("Arial", 11, "bold"))
            else:
                btn.configure(bg=COLOR_SIDEBAR_BG, font=("Arial", 11))
                
        # 화면 전환
        for t, page_info in self.pages.items():
            page_info["frame"].pack_forget()
            
        if title in self.pages:
            page_info = self.pages[title]
            frame = page_info["frame"]
            frame.pack(fill='both', expand=True)
            
            # 지연 로딩
            if not page_info["loaded"]:
                self.load_content(title, page_info)

    def load_content(self, title, page_info):
        frame = page_info["frame"]
        # 로딩 표시
        lbl = tk.Label(frame, text=f"{title} 로딩 중...", bg=COLOR_MAIN_BG, font=("Arial", 12))
        lbl.pack(expand=True)
        self.root.update()
        
        try:
            mod_path = page_info["mod_path"]
            obj_name = page_info["obj_name"]
            obj_type = page_info["obj_type"]
            use_dl   = page_info["use_dl"]
            
            module = __import__(mod_path, fromlist=[obj_name])
            obj = getattr(module, obj_name)
            
            lbl.destroy() # 로딩 라벨 제거
            
            if obj_type == "func":
                obj(frame)
            else:
                if use_dl:
                    try:
                        obj(frame, username=self.username, data_loader=self.data_loader)
                    except TypeError:
                         obj(frame, username=self.username)
                else:
                    obj(frame, username=self.username)
            
            page_info["loaded"] = True
            logger.info(f"Loaded {title}")
            
        except Exception as e:
            lbl.destroy()
            err_msg = f"{title} 로드 오류\n\n{e}\n{traceback.format_exc()}"
            tk.Label(frame, text=err_msg, fg="red", bg=COLOR_MAIN_BG, justify="left").pack(padx=20, pady=20)
            logger.error(f"Failed to load {title}: {traceback.format_exc()}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = LoginWindow(root)
        root.mainloop()
    except Exception as e:
        with open("crash_log.txt", "w") as f:
            f.write(traceback.format_exc())
        messagebox.showerror("치명적 오류", f"앱 실행 중 오류가 발생했습니다.\n{e}")
