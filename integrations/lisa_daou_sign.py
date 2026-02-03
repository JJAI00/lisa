import tkinter as tk
from tkinter import messagebox
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import logging
import os

class LisaDaouSign:
    def __init__(self):
        self.driver = None
        self.setup_logging()
        self.config_file = 'lisa_daou_selectors.json'
        self.selectors = self.load_selectors()
        self.credentials_file = 'daou_credentials.json'
        self.credentials = self.load_credentials()
        
    def setup_logging(self):
        """로깅 설정"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('lisa_daou_sign.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def load_selectors(self):
        """저장된 선택자 로드"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    selectors = json.load(f)
                    self.logger.info(f"저장된 선택자 로드 완료: {len(selectors)}개")
                    return selectors
            else:
                self.logger.info("저장된 선택자 파일이 없습니다. 기본 선택자를 사용합니다.")
                return {}
        except Exception as e:
            self.logger.error(f"선택자 로드 실패: {e}")
            return {}
    
    def save_selectors(self):
        """현재 선택자 저장"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.selectors, f, ensure_ascii=False, indent=2)
            self.logger.info("선택자 저장 완료")
        except Exception as e:
            self.logger.error(f"선택자 저장 실패: {e}")
    
    def update_selector(self, field_name, selector):
        """특정 필드의 선택자 업데이트"""
        self.selectors[field_name] = selector
        self.logger.info(f"선택자 업데이트: {field_name} = {selector}")
        self.save_selectors()
    
    def get_selector(self, field_name, fallback_selectors):
        """저장된 선택자 반환, 없으면 기본값 사용"""
        if field_name in self.selectors:
            self.logger.info(f"저장된 선택자 사용: {field_name} = {self.selectors[field_name]}")
            return [self.selectors[field_name]] + fallback_selectors
        else:
            self.logger.info(f"기본 선택자 사용: {field_name}")
            return fallback_selectors
    
    def load_credentials(self):
        """저장된 로그인 정보 로드"""
        try:
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, 'r', encoding='utf-8') as f:
                    credentials = json.load(f)
                    self.logger.info("저장된 로그인 정보 로드 완료")
                    return credentials
            else:
                self.logger.info("저장된 로그인 정보가 없습니다.")
                return {}
        except Exception as e:
            self.logger.error(f"로그인 정보 로드 실패: {e}")
            return {}
    
    def save_credentials(self, username, password):
        """로그인 정보 저장"""
        try:
            credentials = {
                'username': username,
                'password': password,
                'saved_at': time.strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(self.credentials_file, 'w', encoding='utf-8') as f:
                json.dump(credentials, f, ensure_ascii=False, indent=2)
            self.credentials = credentials
            self.logger.info("로그인 정보 저장 완료")
            return True
        except Exception as e:
            self.logger.error(f"로그인 정보 저장 실패: {e}")
            return False
    
    def update_from_debug_log(self, debug_info):
        """디버깅 로그에서 찾은 선택자들로 업데이트"""
        try:
            for field_name, selector in debug_info.items():
                self.update_selector(field_name, selector)
                self.logger.info(f"디버그 로그에서 업데이트: {field_name} = {selector}")
            
            self.logger.info("디버그 로그에서 선택자 업데이트 완료")
            return True
            
        except Exception as e:
            self.logger.error(f"디버그 로그 업데이트 실패: {e}")
            return False

    def setup_driver(self):
        """헤드리스 브라우저 설정"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            
            # 기존 크롬 프로필 사용 (로그인 세션 공유)
            import os
            chrome_options.add_argument("--user-data-dir=C:\\Users\\HK\\AppData\\Local\\Google\\Chrome\\User Data")
            chrome_options.add_argument("--profile-directory=Default")
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.logger.info("헤드리스 브라우저 설정 완료 (기존 프로필 사용)")
            return True
        except Exception as e:
            self.logger.error(f"브라우저 설정 실패: {e}")
            return False

    def setup_driver_visible(self):
        """화면 표시 브라우저 설정"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-plugins")
            chrome_options.add_argument("--remote-debugging-port=9222")
            
            # 기존 크롬 프로필 사용 (충돌 방지를 위해 포트 지정)
            import os
            chrome_options.add_argument("--user-data-dir=C:\\Users\\HK\\AppData\\Local\\Google\\Chrome\\User Data")
            chrome_options.add_argument("--profile-directory=Default")
            chrome_options.add_argument("--remote-debugging-port=9223")  # 다른 포트 사용
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.logger.info("화면 표시 브라우저 설정 완료 (기존 프로필 사용)")
            return True
        except Exception as e:
            self.logger.error(f"브라우저 설정 실패: {e}")
            # 실패 시 임시 프로필로 재시도
            try:
                self.logger.info("임시 프로필로 재시도합니다")
                chrome_options = Options()
                chrome_options.add_argument("--no-sandbox")
                chrome_options.add_argument("--disable-dev-shm-usage")
                chrome_options.add_argument("--window-size=1920,1080")
                chrome_options.add_argument("--disable-gpu")
                
                import tempfile
                temp_dir = tempfile.mkdtemp()
                chrome_options.add_argument(f"--user-data-dir={temp_dir}")
                
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
                self.logger.info("화면 표시 브라우저 설정 완료 (임시 프로필 사용)")
                return True
            except Exception as e2:
                self.logger.error(f"임시 프로필 설정도 실패: {e2}")
                return False

    def auto_login(self):
        """저장된 로그인 정보로 자동 로그인"""
        try:
            if not self.credentials or 'username' not in self.credentials or 'password' not in self.credentials:
                self.logger.info("저장된 로그인 정보가 없습니다.")
                return False
            
            self.logger.info("자동 로그인 시도 중...")
            
            # 전자결재 양식 페이지로 직접 이동 (권한 없으면 로그인 화면으로 리다이렉트)
            self.driver.get("https://qvrex.daouoffice.com/app/approval/document/new/17923/215127")
            time.sleep(1)
            
            # 사용자명 필드 찾기 및 입력
            username_selectors = [
                "input[name='username']",
                "input[name='userid']", 
                "input[name='id']",
                "input[type='text']"
            ]
            
            username_field = None
            for selector in username_selectors:
                try:
                    username_field = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"사용자명 필드 찾음: {selector}")
                    break
                except:
                    continue
            
            if not username_field:
                self.logger.error("사용자명 필드를 찾을 수 없습니다")
                return False
            
            # 사용자명 입력
            username_field.clear()
            username_field.send_keys(self.credentials['username'])
            self.logger.info("사용자명 입력 완료")
            
            # 비밀번호 필드 찾기 및 입력
            password_selectors = [
                "input[name='password']",
                "input[name='pwd']",
                "input[type='password']"
            ]
            
            password_field = None
            for selector in password_selectors:
                try:
                    password_field = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"비밀번호 필드 찾음: {selector}")
                    break
                except:
                    continue
            
            if not password_field:
                self.logger.error("비밀번호 필드를 찾을 수 없습니다")
                return False
            
            # 비밀번호 입력
            password_field.clear()
            password_field.send_keys(self.credentials['password'])
            self.logger.info("비밀번호 입력 완료")
            
            # 로그인 버튼 찾기 및 클릭
            login_button_selectors = [
                "//button[contains(text(), '로그인')]",
                "//button[contains(text(), 'Login')]",
                "//input[@type='submit']",
                "button[type='submit']"
            ]
            
            login_button = None
            for selector in login_button_selectors:
                try:
                    if selector.startswith("//"):
                        login_button = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        login_button = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    self.logger.info(f"로그인 버튼 찾음: {selector}")
                    break
                except:
                    continue
            
            if login_button:
                login_button.click()
                self.logger.info("로그인 버튼 클릭")
            else:
                # Enter 키로 로그인 시도
                password_field.send_keys(Keys.RETURN)
                self.logger.info("Enter 키로 로그인 시도")
            
            # 로그인 성공 확인
            time.sleep(1)
            current_url = self.driver.current_url
            if "login" not in current_url.lower() and "auth" not in current_url.lower():
                self.logger.info("자동 로그인 성공")
                return True
            else:
                self.logger.error("자동 로그인 실패")
                return False
                
        except Exception as e:
            self.logger.error(f"자동 로그인 실패: {e}")
            return False

    def extract_lisa_data(self):
        """LISA 시스템에서 데이터 추출 (시뮬레이션)"""
        # 실제로는 LISA 시스템에서 데이터를 가져와야 함
        lisa_data = {
            'account_name': '테스트거래처',
            'customer_name': '테스트고객사',
            'category': '소프트웨어',
            'product': '테스트제품',
            'total_sales': '1000000',
            'total_cost': '600000',
            'profit_amount': '400000',
            'profit_margin': '40',
            'purchasing_vendor': '테스트매입처'
        }
        self.logger.info("LISA 데이터 추출 완료")
        return lisa_data

    def fill_form_with_tab_navigation(self, lisa_data):
        """탭 키를 사용한 순차적 폼 입력"""
        try:
            self.logger.info("탭 기반 폼 입력 시작")
            
            # 1. 제목 필드 찾기 및 입력
            title_selectors = [
                "input[name='subject']",
                "input[id='subject']",
                "input[data-name='subject']",
                "input.ipt_editor[data-dsl*='subject']"
            ]
            
            title_field = None
            for selector in title_selectors:
                try:
                    title_field = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self.logger.info(f"제목 필드 찾음: {selector}")
                    break
                except:
                    continue
            
            if not title_field:
                self.logger.error("제목 필드를 찾을 수 없습니다")
                return False
            
            # 제목 입력
            title_text = f"{lisa_data['account_name']}_{lisa_data['customer_name']}_{lisa_data['category']}"
            title_field.clear()
            title_field.send_keys(title_text)
            self.logger.info(f"제목 입력 완료: {title_text}")
            
            # 1번 탭: 거래처명(상호) 필드로 이동
            title_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 거래처명(상호) 필드로 이동")
            
            # 거래처명 입력
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['account_name'])
            self.logger.info(f"거래처명 입력 완료: {lisa_data['account_name']}")
            
            # 1번 더 탭: 사업자번호 필드로 이동
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 더 탭: 사업자번호 필드로 이동")
            
            # 사업자번호 필드 비우기
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            self.logger.info("사업자번호 필드 비움")
            
            # 1번 더 탭: 다음 필드로 이동
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 더 탭: 다음 필드로 이동")
            
            # 화살표 오른쪽 키 입력 (거래처 신규유무 라디오박스에서 "기존" 선택)
            current_field = self.driver.switch_to.active_element
            current_field.send_keys(Keys.RIGHT)
            self.logger.info("화살표 오른쪽 키 입력 완료 (거래처 신규유무 라디오박스에서 '기존' 선택)")
            
            # 8번 탭: 고객사명 필드로 이동
            current_field.send_keys(Keys.TAB * 8)
            time.sleep(1)
            self.logger.info("8번 탭: 고객사명 필드로 이동")
            
            # 고객사명 입력
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['customer_name'])
            self.logger.info(f"고객사명 입력 완료: {lisa_data['customer_name']}")
            
            # 10번 탭 후 화살표 우클릭
            current_field.send_keys(Keys.TAB * 10)
            time.sleep(1)
            self.logger.info("10번 탭: 다음 필드로 이동")
            
            # 화살표 우클릭
            current_field = self.driver.switch_to.active_element
            current_field.send_keys(Keys.RIGHT)
            self.logger.info("화살표 우클릭 입력 완료")
            
            # 1번 더 탭 후 0.5초 대기
            current_field.send_keys(Keys.TAB)
            time.sleep(0.5)
            self.logger.info("1번 더 탭 후 0.5초 대기")
            
            # 엔터 입력
            current_field = self.driver.switch_to.active_element
            current_field.send_keys(Keys.RETURN)
            self.logger.info("엔터 입력 완료")
            
            # 2번 탭 후 품목 입력
            current_field.send_keys(Keys.TAB * 2)
            time.sleep(1)
            self.logger.info("2번 탭: 품목 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['product'])
            self.logger.info(f"품목 입력 완료: {lisa_data['product']}")
            
            # 1번 탭 후 판매 합계 입력
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 판매 합계 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['total_sales'])
            self.logger.info(f"판매 합계 입력 완료: {lisa_data['total_sales']}")
            
            # 1번 탭 후 원가계 입력
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 원가계 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['total_cost'])
            self.logger.info(f"원가계 입력 완료: {lisa_data['total_cost']}")
            
            # 1번 탭 후 이익액 입력
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 이익액 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['profit_amount'])
            self.logger.info(f"이익액 입력 완료: {lisa_data['profit_amount']}")
            
            # 1번 탭 후 이익율 입력
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 이익율 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['profit_margin'])
            self.logger.info(f"이익율 입력 완료: {lisa_data['profit_margin']}")
            
            # 1번 탭 후 매입처 입력
            current_field.send_keys(Keys.TAB)
            time.sleep(1)
            self.logger.info("1번 탭: 매입처 필드로 이동")
            
            current_field = self.driver.switch_to.active_element
            current_field.clear()
            current_field.send_keys(lisa_data['purchasing_vendor'])
            self.logger.info(f"매입처 입력 완료: {lisa_data['purchasing_vendor']}")
            
            self.logger.info("탭 기반 폼 입력 완료")
            return True
            
        except Exception as e:
            self.logger.error(f"탭 기반 폼 입력 실패: {e}")
            return False

    def process_approval(self, skip_login=False, auto_login=False):
        """전자결재 처리 메인 함수 (기본 데이터 사용)"""
        lisa_data = self.extract_lisa_data()
        return self.process_approval_with_data(lisa_data, skip_login, auto_login)
    
    def process_approval_with_data(self, lisa_data, skip_login=False, auto_login=False):
        """전자결재 처리 메인 함수 (외부 데이터 사용)"""
        try:
            self.logger.info("전자결재 양식 입력 프로세스 시작")
            
            # 1. 브라우저 설정 (화면 표시 모드로 진행)
            self.setup_driver_visible()
            
            # 2. 로그인 처리
            if auto_login and self.credentials:
                # 자동 로그인 시도
                self.logger.info("자동 로그인 시도")
                if self.auto_login():
                    self.logger.info("자동 로그인 성공")
                else:
                    self.logger.warning("자동 로그인 실패, 수동 로그인 필요")
                    skip_login = False
            
            # 3. 다우오피스 접속
            self.driver.get("https://qvrex.daouoffice.com/app/approval/document/new/17923/215127")
            self.logger.info("다우오피스 전자결재 양식 페이지 접속")
            
            # 4. 로그인 상태 확인 (skip_login이 True인 경우)
            if skip_login:
                self.logger.info("로그인 스킵 모드로 진행")
                time.sleep(1)
                
                # URL에서 로그인 필요 여부 확인
                current_url = self.driver.current_url
                self.logger.info(f"현재 URL: {current_url}")
                
                if "login" in current_url.lower() or "auth" in current_url.lower():
                    self.logger.info("로그인이 필요합니다. 로그인 페이지로 이동됨")
                    # 로그인 페이지인 경우 승인 메인 페이지로 이동
                    self.driver.get("https://qvrex.daouoffice.com/app/approval")
                    time.sleep(1)
                    # 다시 양식 페이지로 이동
                    self.driver.get("https://qvrex.daouoffice.com/app/approval/document/new/17923/215127")
                    time.sleep(1)
                else:
                    self.logger.info("이미 로그인된 상태입니다")
                    # 로그인된 상태에서 바로 양식 페이지로 이동
                    self.driver.get("https://qvrex.daouoffice.com/app/approval/document/new/17923/215127")
                    time.sleep(1)
            
            # 5. 페이지 로딩 대기
            time.sleep(1)
            
            # 6. 탭 기반 폼 입력
            if not self.fill_form_with_tab_navigation(lisa_data):
                raise Exception("탭 기반 폼 입력 실패")
            
            self.logger.info("양식 입력 완료 - 브라우저를 유지합니다")
            self.logger.info("사용자가 직접 결재를 제출할 수 있습니다")
            
            return True
            
        except Exception as e:
            self.logger.error(f"전자결재 양식 입력 실패: {e}")
            if self.driver:
                self.driver.quit()
            return False

def show_context_menu(event, tree, lisa_daou=None):
    """우클릭 컨텍스트 메뉴 표시"""
    try:
        # 선택된 항목 확인
        selected_item = tree.selection()
        if not selected_item:
            return
        
        # 선택된 항목의 데이터 가져오기
        item_data = tree.item(selected_item[0])
        values = item_data['values']
        
        # Status 확인 (동적으로 Status 컬럼 찾기)
        columns = tree['columns']
        try:
            status_idx = list(columns).index('Status')
            status = values[status_idx] if status_idx < len(values) else "N/A"
        except ValueError:
            # Status 컬럼을 찾을 수 없는 경우 첫 번째 컬럼 사용
            status = values[0] if len(values) > 0 else "N/A"
        
        # Status 값 정규화 (숫자, 문자열, 퍼센트 등 다양한 형식 지원)
        normalized_status = str(status).strip()
        
        # 컨텍스트 메뉴 생성 (항상 표시)
        context_menu = tk.Menu(event.widget, tearoff=0)
        context_menu.add_command(
            label="전자결재 보내기",
            command=lambda: send_to_approval_with_status_check(tree, selected_item[0], normalized_status)
        )
        
        # 메뉴 표시
        context_menu.tk_popup(event.x_root, event.y_root)
        
    except Exception as e:
        messagebox.showerror("오류", f"컨텍스트 메뉴 생성 실패: {e}")

def send_to_approval_with_status_check(tree, selected_item, current_status):
    """Status를 확인하고 전자결재 보내기 처리"""
    try:
        # Status가 80%인지 확인
        if current_status in ["80", "80%", "80.0", "80.0%"]:
            # 80%인 경우 바로 전자결재 실행
            send_to_approval(tree, selected_item)
        else:
            # 80%가 아닌 경우 안내 메시지 표시
            messagebox.showwarning(
                "Status 확인", 
                f"현재 Status가 '{current_status}'입니다.\n\n"
                "전자결재를 진행하려면 Status를 80%로 수정 후 다시 시도해 주세요.\n\n"
                "Status 메뉴에서 80%를 선택하여 변경할 수 있습니다."
            )
    except Exception as e:
        messagebox.showerror("오류", f"전자결재 전송 실패: {e}")

def send_to_approval(tree, selected_item):
    """전자결재 보내기 함수"""
    try:
        # 선택된 항목의 데이터 가져오기
        item_data = tree.item(selected_item)
        values = item_data['values']
        
        # 트리뷰의 컬럼 정보 가져오기
        columns = tree['columns']
        
        # 컬럼 인덱스 찾기 함수
        def find_column_index(column_name):
            try:
                return list(columns).index(column_name)
            except ValueError:
                return -1
        
        # ID 컬럼 인덱스 찾기
        id_idx = find_column_index('ID')
        if id_idx == -1:
            messagebox.showerror("오류", "ID 컬럼을 찾을 수 없습니다.")
            return
        
        # 선택된 행의 ID 값 가져오기
        selected_id = values[id_idx] if id_idx < len(values) else None
        if not selected_id:
            messagebox.showerror("오류", "선택된 행의 ID 값이 없습니다.")
            return
        
        print(f"선택된 ID: {selected_id}")
        
        # 해당 ID의 모든 행 데이터 수집
        all_rows_with_id = []
        for item in tree.get_children():
            item_values = tree.item(item)['values']
            if id_idx < len(item_values) and item_values[id_idx] == selected_id:
                all_rows_with_id.append(item_values)
        
        print(f"ID {selected_id}에 해당하는 행 개수: {len(all_rows_with_id)}")
        
        if not all_rows_with_id:
            messagebox.showerror("오류", f"ID {selected_id}에 해당하는 데이터를 찾을 수 없습니다.")
            return
        
        # 각 필드의 인덱스 찾기
        account_idx = find_column_index('거래처명')
        customer_idx = find_column_index('고객사명')
        category_idx = find_column_index('대분류')
        product_idx = find_column_index('제품')
        sales_idx = find_column_index('판매 합계')
        cost_idx = find_column_index('원가계')
        profit_amount_idx = find_column_index('이익액')
        profit_margin_idx = find_column_index('이익율')
        vendor_idx = find_column_index('매입처')
        
        # 첫 번째 행에서 기본 정보 가져오기
        first_row = all_rows_with_id[0]
        
        # 숫자 데이터 합계 계산
        total_sales_sum = 0
        total_cost_sum = 0
        profit_amount_sum = 0
        
        for row_values in all_rows_with_id:
            # 판매 합계 합계
            if sales_idx >= 0 and sales_idx < len(row_values):
                try:
                    sales_value = str(row_values[sales_idx]).replace(',', '').replace('₩', '').strip()
                    if sales_value and sales_value != '':
                        total_sales_sum += float(sales_value)
                except (ValueError, TypeError):
                    pass
            
            # 원가계 합계
            if cost_idx >= 0 and cost_idx < len(row_values):
                try:
                    cost_value = str(row_values[cost_idx]).replace(',', '').replace('₩', '').strip()
                    if cost_value and cost_value != '':
                        total_cost_sum += float(cost_value)
                except (ValueError, TypeError):
                    pass
            
            # 이익액 합계
            if profit_amount_idx >= 0 and profit_amount_idx < len(row_values):
                try:
                    profit_value = str(row_values[profit_amount_idx]).replace(',', '').replace('₩', '').strip()
                    if profit_value and profit_value != '':
                        profit_amount_sum += float(profit_value)
                except (ValueError, TypeError):
                    pass
        
        # 이익율 계산 (이익액 / 판매 합계 * 100)
        profit_margin = 0
        if total_sales_sum > 0:
            profit_margin = (profit_amount_sum / total_sales_sum) * 100
        
        # 첫 번째 제품명 가져오기
        first_product_name = '테스트제품'
        if product_idx >= 0 and len(first_row) > product_idx:
            first_product_name = str(first_row[product_idx]).strip()
        
        # 제품명 구성 (첫 번째 제품명 + 개수에 따라 "외" 추가)
        if len(all_rows_with_id) > 1:
            product_display = f"{first_product_name} 외"
        else:
            product_display = first_product_name

        # 데이터 구성
        lisa_data = {
            'account_name': first_row[account_idx] if account_idx >= 0 and len(first_row) > account_idx else '테스트거래처',
            'customer_name': first_row[customer_idx] if customer_idx >= 0 and len(first_row) > customer_idx else '테스트고객사',
            'category': first_row[category_idx] if category_idx >= 0 and len(first_row) > category_idx else '소프트웨어',
            'product': product_display,  # 첫 번째 제품명 + 필요시 "외" 추가
            'total_sales': str(int(total_sales_sum)),
            'total_cost': str(int(total_cost_sum)),
            'profit_amount': str(int(profit_amount_sum)),
            'profit_margin': f"{profit_margin:.2f}",
            'purchasing_vendor': first_row[vendor_idx] if vendor_idx >= 0 and len(first_row) > vendor_idx else '테스트매입처'
        }
        
        print(f"집계된 데이터: {lisa_data}")
        print(f"사용 가능한 컬럼: {list(columns)}")
        
        lisa_daou = LisaDaouSign()
        
        # 로그인 옵션 선택 (로그인 스킵 제거)
        login_options = [
            "1. 자동 로그인 (저장된 정보 사용)",
            "2. 수동 로그인 (로그인 정보 입력)"
        ]
        
        if lisa_daou.credentials:
            login_options[0] += f" - {lisa_daou.credentials.get('username', 'Unknown')}"
        
        # 커스텀 로그인 옵션 다이얼로그 생성
        login_choice = show_login_options_dialog(login_options)
        
        if login_choice == "auto":  # 자동 로그인
            success = lisa_daou.process_approval_with_data(lisa_data, auto_login=True)
        elif login_choice == "manual":  # 수동 로그인
            # 로그인 정보 입력 다이얼로그 표시
            login_dialog = LoginDialog()
            if login_dialog.result:
                username, password = login_dialog.result
                # 로그인 정보 저장 여부 확인
                save_credentials = messagebox.askyesno("로그인 정보 저장", 
                                                     "로그인 정보를 저장하시겠습니까?\n\n"
                                                     "예: 저장 (다음에 자동 로그인 가능)\n"
                                                     "아니오: 저장하지 않음")
                if save_credentials:
                    lisa_daou.save_credentials(username, password)
                
                success = lisa_daou.process_approval_with_data(lisa_data, skip_login=False)
        else:
            return  # 사용자가 취소한 경우
        
        if success:
            messagebox.showinfo("성공", "전자결재 양식이 자동으로 입력되었습니다.\n브라우저에서 직접 결재를 제출하실 수 있습니다.")
        else:
            messagebox.showerror("실패", "전자결재 양식 입력에 실패했습니다. 로그를 확인해주세요.")
        
    except Exception as e:
        messagebox.showerror("오류", f"전자결재 전송 실패: {e}")

class LoginDialog:
    """로그인 정보 입력 다이얼로그"""
    def __init__(self):
        self.result = None
        
        self.dialog = tk.Toplevel()
        self.dialog.title("다우오피스 로그인")
        self.dialog.geometry("300x150")
        self.dialog.resizable(False, False)
        
        # 중앙 정렬
        self.dialog.transient()
        self.dialog.grab_set()
        
        # 위젯 생성
        tk.Label(self.dialog, text="아이디:").pack(pady=5)
        self.username_entry = tk.Entry(self.dialog)
        self.username_entry.pack(pady=5)
        
        tk.Label(self.dialog, text="비밀번호:").pack(pady=5)
        self.password_entry = tk.Entry(self.dialog, show="*")
        self.password_entry.pack(pady=5)
        
        # 버튼
        button_frame = tk.Frame(self.dialog)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="확인", command=self.confirm).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="취소", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        # 포커스 설정
        self.username_entry.focus()
        
        # 다이얼로그가 닫힐 때까지 대기
        self.dialog.wait_window()
    
    def confirm(self):
        """확인 버튼 클릭"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if username and password:
            self.result = (username, password)
            self.dialog.destroy()
        else:
            messagebox.showwarning("경고", "아이디와 비밀번호를 모두 입력해주세요.")
    
    def cancel(self):
        """취소 버튼 클릭"""
        self.dialog.destroy()

def show_login_options_dialog(login_options):
    """커스텀 로그인 옵션 다이얼로그"""
    result = None
    
    dialog = tk.Toplevel()
    dialog.title("로그인 옵션")
    dialog.geometry("400x300")
    dialog.resizable(False, False)
    
    # 중앙 정렬
    dialog.transient()
    dialog.grab_set()
    
    # 메인 프레임
    main_frame = tk.Frame(dialog, bg="white")
    main_frame.pack(fill='both', expand=True, padx=20, pady=20)
    
    # 제목
    title_label = tk.Label(main_frame, text="로그인 방법을 선택하세요:", 
                          font=("맑은 고딕", 12, "bold"), bg="white", fg="#374151")
    title_label.pack(pady=(0, 20))
    
    # 옵션 버튼들
    def select_auto():
        nonlocal result
        result = "auto"
        dialog.destroy()
    
    def select_manual():
        nonlocal result
        result = "manual"
        dialog.destroy()
    
    def cancel():
        nonlocal result
        result = None
        dialog.destroy()
    
    # 자동 로그인 버튼
    auto_btn = tk.Button(main_frame, text=login_options[0], command=select_auto,
                         font=("맑은 고딕", 11), bg="#3B82F6", fg="white",
                         relief="flat", padx=20, pady=10, width=30)
    auto_btn.pack(pady=5)
    auto_btn.bind('<Enter>', lambda e: auto_btn.configure(bg="#2563EB"))
    auto_btn.bind('<Leave>', lambda e: auto_btn.configure(bg="#3B82F6"))
    
    # 수동 로그인 버튼
    manual_btn = tk.Button(main_frame, text=login_options[1], command=select_manual,
                          font=("맑은 고딕", 11), bg="#10B981", fg="white",
                          relief="flat", padx=20, pady=10, width=30)
    manual_btn.pack(pady=5)
    manual_btn.bind('<Enter>', lambda e: manual_btn.configure(bg="#059669"))
    manual_btn.bind('<Leave>', lambda e: manual_btn.configure(bg="#10B981"))
    
    # 취소 버튼 (창 닫기)
    cancel_btn = tk.Button(main_frame, text="취소", command=cancel,
                          font=("맑은 고딕", 11), bg="#6B7280", fg="white",
                          relief="flat", padx=20, pady=8, width=15)
    cancel_btn.pack(pady=(20, 0))
    cancel_btn.bind('<Enter>', lambda e: cancel_btn.configure(bg="#4B5563"))
    cancel_btn.bind('<Leave>', lambda e: cancel_btn.configure(bg="#6B7280"))
    
    # 다이얼로그가 닫힐 때까지 대기
    dialog.wait_window()
    
    return result

if __name__ == "__main__":
    # 테스트용 코드
    lisa_daou = LisaDaouSign()
    print("LISA-Daou 전자결재 연동 모듈이 준비되었습니다.")
    
    # 디버그 로그에서 찾은 선택자들로 업데이트 (예시)
    debug_selectors = {
        "title": "input[name='subject']",
        "customer_name": "input[name='editorForm_4']",
        "existing_radio": "//input[@name='editorForm_6' and @value='기존']",
        "credit_rating": "//input[@name='editorForm_7' and @value='B']",
        "company_name": "input[name='editorForm_9']",
        "item": "input[name='editorForm_27']",
        "sales_amount": "input[name='editorForm_28']",
        "cost_of_sales": "input[name='editorForm_29']",
        "sales_profit": "input[name='editorForm_30']",
        "profit_margin": "input[name='editorForm_31']",
        "purchasing_vendor": "input[name='editorForm_32']",
        "purchasing_month": "input[name='editorForm_33']"
    }
    
    lisa_daou.update_from_debug_log(debug_selectors)
    print("디버그 선택자 업데이트 완료") 