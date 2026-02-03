import threading
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
import time
import hashlib
from integrations.lisa_logging import setup_logger
from config import KEY_FILE, SCOPES, SHEET_IDS, CACHE_TTL_SECONDS

# 모듈 로거 설정
logger = setup_logger('data_loader')

# 전역 변수
_global_auth = None
_cache = {}
_cache_timestamps = {}
CACHE_DURATION = CACHE_TTL_SECONDS  # config에서 가져옴


def get_global_auth():
    """전역 인증 객체 반환 (싱글톤 패턴)"""
    global _global_auth
    if _global_auth is None:
        creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
        _global_auth = gspread.authorize(creds)
    return _global_auth

def get_cached_data(key):
    """캐시된 데이터 반환"""
    if key in _cache and key in _cache_timestamps:
        if time.time() - _cache_timestamps[key] < CACHE_DURATION:
            return _cache[key]
        else:
            # 캐시 만료
            del _cache[key]
            del _cache_timestamps[key]
    return None

def set_cached_data(key, data):
    """데이터를 캐시에 저장"""
    _cache[key] = data
    _cache_timestamps[key] = time.time()

def clear_cache():
    """캐시 초기화"""
    global _cache, _cache_timestamps
    _cache.clear()
    _cache_timestamps.clear()

def safe_api_call(func, max_retries=3, base_delay=1):
    """안전한 API 호출 (재시도 로직 포함)"""
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "Quota exceeded" in error_str:
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)  # 지수 백오프
                    logger.warning(f"API 호출 제한 도달. {delay}초 후 재시도... (시도 {attempt + 1}/{max_retries})")
                    time.sleep(delay)
                    continue
                else:
                    logger.error("API 호출 제한으로 인해 요청을 중단합니다. 캐시된 데이터를 사용하세요.")
                    return None
            elif "503" in error_str or "service is currently unavailable" in error_str:
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)  # 지수 백오프
                    logger.warning(f"Google Sheets 서비스 일시적 사용 불가. {delay}초 후 재시도... (시도 {attempt + 1}/{max_retries})")
                    time.sleep(delay)
                    continue
                else:
                    logger.error("Google Sheets 서비스 일시적 사용 불가. 캐시된 데이터를 사용하세요.")
                    return None
            elif "500" in error_str or "502" in error_str:
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)  # 지수 백오프
                    logger.warning(f"Google Sheets 서버 오류. {delay}초 후 재시도... (시도 {attempt + 1}/{max_retries})")
                    time.sleep(delay)
                    continue
                else:
                    logger.error("Google Sheets 서버 오류. 캐시된 데이터를 사용하세요.")
                    return None
            raise e
    return None

class DataLoader:
    """데이터 로딩 최적화 클래스"""
    
    def __init__(self, key_file, scopes):
        self.key_file = key_file
        self.scopes = scopes
        self.creds = None
        self.gc = None
        
        # 데이터 캐시
        self.renewal_data = None
        self.customer_data = None
        self.price_data = None
        self.quote_data = None
        self.log_data = None
        
        # 로딩 상태
        self.is_loading = False
        self.loading_progress = 0
        self.loading_status = ""
        
        # 이번달 데이터만 먼저 로드
        self.current_month_data = None
        
        # 백그라운드 스레드 관리
        self.background_thread = None
        
    def _authorize(self):
        """전역 인증 객체 사용"""
        if not self.gc:
            self.gc = get_global_auth()
        return self.gc
    
    def load_current_month_data(self):
        """이번달 데이터 로드 (최적화된 방식)"""
        try:
            # 캐시 확인
            cache_key = "current_month_data"
            cached_data = get_cached_data(cache_key)
            if cached_data is not None:
                logger.info("이번달 데이터를 캐시에서 로드했습니다.")
                return cached_data

            logger.info("캐시된 데이터가 없어 Google Sheets에서 새로 로드합니다...")
            self.is_loading = True
            self.loading_progress = 0
            self.loading_status = "이번달 데이터 로딩 중..."

            # 전역 인증 사용
            gc = self._authorize()
            
            # 1. renewal_list 이번달 데이터
            def load_renewal_data():
                sh = gc.open_by_key(SHEET_IDS['renewal_list'])
                try:
                    ws = sh.worksheet('renewal_list')
                except:
                    ws = sh.worksheet('Sheet1')
                return ws.get_all_records()

            all_data = safe_api_call(load_renewal_data)
            if all_data:
                df = pd.DataFrame(all_data)
                df.columns = df.columns.astype(str).str.strip()
                
                # 날짜 컬럼 변환
                if '계산서 발행일' in df.columns:
                    df['계산서_날짜'] = pd.to_datetime(df['계산서 발행일'], errors='coerce')
                if '만료일' in df.columns:
                    df['만료_날짜'] = pd.to_datetime(df['만료일'], errors='coerce')
                
                # 이번달 필터링
                today = datetime.now()
                start_of_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                end_of_month = (start_of_month + timedelta(days=32)).replace(day=1) - timedelta(seconds=1)
                
                if '계산서_날짜' in df.columns:
                    mask = (df['계산서_날짜'] >= start_of_month) & (df['계산서_날짜'] <= end_of_month)
                    self.current_month_data = df[mask].copy()
                else:
                    self.current_month_data = df.copy()

            self.loading_progress = 50

            # 2. quote_list 데이터
            def load_quote_data():
                sh = gc.open_by_key(SHEET_IDS['quote_list'])
                try:
                    ws = sh.worksheet('quote_list')
                except:
                    ws = sh.worksheet('Sheet1')
                return ws.get_all_records()

            quote_data = safe_api_call(load_quote_data)
            if quote_data:
                self.quote_data = pd.DataFrame(quote_data)
                self.quote_data.columns = self.quote_data.columns.astype(str).str.strip()
                
                # 일자 컬럼 처리
                if '일자' in self.quote_data.columns:
                    logger.debug(f"DataLoader: 견적서 일자 컬럼 원본 데이터 샘플: {self.quote_data['일자'].head(3).tolist()}")
                    
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
                                    logger.debug(f"DataLoader: 날짜 파싱 실패 - {date_str}")
                                    return None
                    
                    # 일자 컬럼을 파싱
                    self.quote_data['일자_parsed'] = self.quote_data['일자'].apply(parse_korean_date)
                    
                    # 파싱 성공한 데이터 확인
                    valid_dates = self.quote_data[self.quote_data['일자_parsed'].notna()]
                    failed_dates = self.quote_data[self.quote_data['일자_parsed'].isna()]
                    
                    logger.debug(f"DataLoader: 날짜 파싱 결과 - 성공: {len(valid_dates)}행, 실패: {len(failed_dates)}행")
                    
                    if not failed_dates.empty:
                        logger.debug(f"DataLoader: 파싱 실패한 날짜 샘플: {failed_dates['일자'].head(3).tolist()}")
                    
                    # 파싱된 날짜를 문자열로 변환하여 일자 컬럼 업데이트
                    self.quote_data['일자'] = self.quote_data['일자_parsed'].dt.strftime('%Y-%m-%d %H:%M:%S')
                    self.quote_data = self.quote_data.drop('일자_parsed', axis=1)
                    
                    logger.debug(f"DataLoader: 날짜 변환 완료 - 샘플: {self.quote_data['일자'].head(3).tolist()}")

            self.loading_progress = 75

            # 3. log_data
            def load_log_data():
                sh = gc.open_by_key(SHEET_IDS['log_list'])
                ws = sh.worksheet('Sheet1')
                return ws.get_all_records()

            log_data = safe_api_call(load_log_data)
            if log_data:
                self.log_data = pd.DataFrame(log_data)
                self.log_data.columns = self.log_data.columns.astype(str).str.strip()
                if '일시' in self.log_data.columns:
                    self.log_data['일시'] = pd.to_datetime(
                        self.log_data['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
                    )
                    self.log_data.sort_values('일시', ascending=False, inplace=True)

            # 결과 캐시에 저장
            result_data = {
                'current_month': self.current_month_data,
                'quote_data': self.quote_data,
                'log_data': self.log_data
            }
            set_cached_data(cache_key, result_data)

            self.loading_progress = 100
            self.is_loading = False
            self.loading_status = "로딩 완료"
            
            logger.info(f"초기 데이터 로드 완료: {len(self.current_month_data)}행")
            return result_data

        except Exception as e:
            logger.error(f"초기 데이터 로드 실패: {e}")
            self.is_loading = False
            self.loading_status = f"로딩 실패: {str(e)}"
            
            # 에러 발생 시 빈 데이터 반환
            return {
                'current_month': pd.DataFrame(),
                'quote_data': pd.DataFrame(),
                'log_data': pd.DataFrame()
            }
    
    def load_full_data_async(self, callback=None):
        """전체 데이터를 백그라운드에서 로드 (최적화된 방식)"""
        def load_in_background():
            try:
                self.is_loading = True
                self.loading_progress = 0
                
                # 캐시 확인
                cache_key = "full_data"
                cached_data = get_cached_data(cache_key)
                if cached_data is not None:
                    logger.info("전체 데이터를 캐시에서 로드했습니다.")
                    self.renewal_data = cached_data.get('renewal_data')
                    self.customer_data = cached_data.get('customer_data')
                    self.price_data = cached_data.get('price_data')
                    self.quote_data = cached_data.get('quote_data')
                    self.log_data = cached_data.get('log_data')
                    if callback:
                        callback()
                    return

                # 전역 인증 사용
                gc = self._authorize()
                
                # 1. renewal_list 전체 데이터
                self.loading_status = "리뉴얼 데이터 로딩 중..."
                self.loading_progress = 20
                
                def load_renewal():
                    sh = gc.open_by_key(SHEET_IDS['renewal_list'])
                    try:
                        ws = sh.worksheet('renewal_list')
                    except:
                        ws = sh.worksheet('Sheet1')
                    return ws.get_all_records()

                all_data = safe_api_call(load_renewal)
                if all_data:
                    self.renewal_data = pd.DataFrame(all_data)
                    self.renewal_data.columns = self.renewal_data.columns.astype(str).str.strip()
                    
                    # 날짜 컬럼 변환
                    if '계산서 발행일' in self.renewal_data.columns:
                        self.renewal_data['계산서_날짜'] = pd.to_datetime(self.renewal_data['계산서 발행일'], errors='coerce')
                    if '만료일' in self.renewal_data.columns:
                        self.renewal_data['만료_날짜'] = pd.to_datetime(self.renewal_data['만료일'], errors='coerce')

                self.loading_progress = 40
                
                # 2. customer_list 데이터
                self.loading_status = "고객사 데이터 로딩 중..."
                
                def load_customer():
                    sh = gc.open_by_key(SHEET_IDS['customer_list'])
                    ws = sh.worksheet('Sheet1')
                    return ws.get_all_values()

                all_values = safe_api_call(load_customer)
                if all_values:
                    raw_header = [h.strip() for h in all_values[0]]
                    valid_idx = [i for i, h in enumerate(raw_header) if h]
                    headers = [raw_header[i] for i in valid_idx]
                    rows = [[row[i] for i in valid_idx] for row in all_values[1:]]
                    self.customer_data = pd.DataFrame(rows, columns=headers)
                    self.customer_data.columns = self.customer_data.columns.str.strip()
                else:
                    # fallback: renewal_list에서 고객사 정보 추출
                    if self.renewal_data is not None and not self.renewal_data.empty:
                        info_cols = [c for c in [
                            '고객사명','담당자명','담당자','직함','이메일','연락처','핸드폰','주소','사업자번호','영문 회사명','영문 담당자명','영문 주소','비고'
                        ] if c in self.renewal_data.columns]
                        if '고객사명' in self.renewal_data.columns and info_cols:
                            self.customer_data = self.renewal_data.drop_duplicates(subset='고객사명')[info_cols]
                        else:
                            self.customer_data = pd.DataFrame()
                    else:
                        self.customer_data = pd.DataFrame()
                
                self.loading_progress = 60
                
                # 3. price_list 데이터
                self.loading_status = "가격 데이터 로딩 중..."
                
                def load_price():
                    sh = gc.open_by_key(SHEET_IDS['price_list'])
                    ws = sh.worksheet('Sheet1')
                    return ws.get_all_records()

                price_data = safe_api_call(load_price)
                if price_data:
                    self.price_data = pd.DataFrame(price_data)
                    self.price_data.columns = self.price_data.columns.astype(str).str.strip()
                else:
                    self.price_data = pd.DataFrame()
                
                self.loading_progress = 80
                
                # 4. quote_list 데이터
                self.loading_status = "견적서 데이터 로딩 중..."
                
                def load_quote():
                    sh = gc.open_by_key(SHEET_IDS['quote_list'])
                    try:
                        ws = sh.worksheet('quote_list')
                    except:
                        ws = sh.worksheet('Sheet1')
                    return ws.get_all_records()

                quote_data = safe_api_call(load_quote)
                if quote_data:
                    self.quote_data = pd.DataFrame(quote_data)
                    self.quote_data.columns = self.quote_data.columns.astype(str).str.strip()
                    
                    # 일자 컬럼 처리
                    if '일자' in self.quote_data.columns:
                        logger.debug(f"DataLoader: 견적서 일자 컬럼 원본 데이터 샘플: {self.quote_data['일자'].head(3).tolist()}")
                        
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
                                        logger.debug(f"DataLoader: 날짜 파싱 실패 - {date_str}")
                                        return None
                        
                        # 일자 컬럼을 파싱
                        self.quote_data['일자_parsed'] = self.quote_data['일자'].apply(parse_korean_date)
                        
                        # 파싱 성공한 데이터 확인
                        valid_dates = self.quote_data[self.quote_data['일자_parsed'].notna()]
                        failed_dates = self.quote_data[self.quote_data['일자_parsed'].isna()]
                        
                        logger.debug(f"DataLoader: 날짜 파싱 결과 - 성공: {len(valid_dates)}행, 실패: {len(failed_dates)}행")
                        
                        if not failed_dates.empty:
                            logger.debug(f"DataLoader: 파싱 실패한 날짜 샘플: {failed_dates['일자'].head(3).tolist()}")
                        
                        # 파싱된 날짜를 문자열로 변환하여 일자 컬럼 업데이트
                        self.quote_data['일자'] = self.quote_data['일자_parsed'].dt.strftime('%Y-%m-%d %H:%M:%S')
                        self.quote_data = self.quote_data.drop('일자_parsed', axis=1)
                        
                        logger.debug(f"DataLoader: 날짜 변환 완료 - 샘플: {self.quote_data['일자'].head(3).tolist()}")
                else:
                    self.quote_data = pd.DataFrame()
                
                self.loading_progress = 90
                
                # 5. log_data
                self.loading_status = "로그 데이터 로딩 중..."
                
                def load_log():
                    sh = gc.open_by_key('1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk')
                    ws = sh.worksheet('Sheet1')
                    return ws.get_all_records()

                log_data = safe_api_call(load_log)
                if log_data:
                    self.log_data = pd.DataFrame(log_data)
                    self.log_data.columns = self.log_data.columns.astype(str).str.strip()
                    if '일시' in self.log_data.columns:
                        self.log_data['일시'] = pd.to_datetime(
                            self.log_data['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
                        )
                        self.log_data.sort_values('일시', ascending=False, inplace=True)
                else:
                    self.log_data = pd.DataFrame()

                # 결과 캐시에 저장
                result_data = {
                    'renewal_data': self.renewal_data,
                    'customer_data': self.customer_data,
                    'price_data': self.price_data,
                    'quote_data': self.quote_data,
                    'log_data': self.log_data
                }
                set_cached_data(cache_key, result_data)

                self.loading_progress = 100
                self.is_loading = False
                self.loading_status = "로딩 완료"
                
                logger.info(f"전체 데이터 로딩 완료: renewal={len(self.renewal_data) if self.renewal_data is not None else 0}행, customer={len(self.customer_data) if self.customer_data is not None else 0}행, price={len(self.price_data) if self.price_data is not None else 0}행, quote={len(self.quote_data) if self.quote_data is not None else 0}행, log={len(self.log_data) if self.log_data is not None else 0}행")
                
                if callback:
                    callback()
                    
            except Exception as e:
                logger.error(f"전체 데이터 로딩 실패: {e}")
                self.is_loading = False
                self.loading_status = f"로딩 실패: {str(e)}"
        
        # 백그라운드 스레드에서 실행
        if self.background_thread and self.background_thread.is_alive():
            logger.debug("이전 백그라운드 스레드가 실행 중입니다. 새로운 로딩을 시작합니다.")
        
        self.background_thread = threading.Thread(target=load_in_background, daemon=True)
        self.background_thread.start()
    
    def get_renewal_data(self):
        """리뉴얼 데이터 반환 (캐시된 데이터 우선)"""
        # 전체 데이터가 로드되었으면 그것을 반환
        if self.renewal_data is not None:
            return self.renewal_data
        # 전체 데이터가 없고 이번달 데이터만 있으면 전체 데이터 로딩 시작
        elif self.current_month_data is not None and self.renewal_data is None:
            logger.debug("전체 데이터가 로드되지 않아 백그라운드에서 로딩을 시작합니다.")
            self.load_full_data_async()
            return self.current_month_data
        else:
            # 데이터가 없으면 즉시 로드 시도
            logger.debug("데이터가 없어 즉시 로드를 시도합니다.")
            try:
                initial_data = self.load_current_month_data()
                if initial_data and 'current_month' in initial_data:
                    self.current_month_data = initial_data['current_month']
                    return self.current_month_data
            except Exception as e:
                logger.error(f"즉시 로드 실패: {e}")
            return pd.DataFrame()
    
    def get_customer_data(self):
        """고객사 데이터 반환"""
        # 전체 데이터가 로드되었으면 그것을 반환
        if self.customer_data is not None:
            return self.customer_data
        # 전체 데이터가 없으면 빈 데이터프레임 반환
        else:
            return pd.DataFrame()
    
    def get_price_data(self):
        """가격 데이터 반환"""
        return self.price_data if self.price_data is not None else pd.DataFrame()
    
    def get_log_data(self):
        """로그 데이터 반환"""
        return self.log_data if self.log_data is not None else pd.DataFrame()
    
    def get_quote_data(self):
        """견적서 데이터 반환"""
        return self.quote_data if self.quote_data is not None else pd.DataFrame()
    
    def clear_all_cache(self):
        """모든 캐시 데이터 클리어"""
        logger.info("DataLoader: 모든 캐시 데이터 클리어")
        self.renewal_data = None
        self.customer_data = None
        self.price_data = None
        self.quote_data = None
        self.log_data = None
        self.current_month_data = None
        clear_cache()  # 전역 캐시도 클리어
    
    def force_reload_quote_data(self):
        """견적서 데이터 강제 재로드"""
        try:
            logger.info("DataLoader: 견적서 데이터 강제 재로드 시작")
            
            # 캐시 클리어
            self.quote_data = None
            clear_cache()
            
            # 구글 시트에서 직접 로드
            gc = get_global_auth()
            
            def load_quote_data():
                sh = gc.open_by_key(SHEET_IDS['quote_list'])
                try:
                    ws = sh.worksheet('quote_list')
                except:
                    ws = sh.worksheet('Sheet1')
                return ws.get_all_records()
            
            quote_data = safe_api_call(load_quote_data)
            if quote_data:
                self.quote_data = pd.DataFrame(quote_data)
                self.quote_data.columns = self.quote_data.columns.str.strip()
                
                # 일자 컬럼의 원본 데이터 확인
                if '일자' in self.quote_data.columns:
                    logger.debug(f"DataLoader: 일자 컬럼 원본 데이터 샘플: {self.quote_data['일자'].head(5).tolist()}")
                    
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
                                    logger.debug(f"DataLoader: 날짜 파싱 실패 - {date_str}")
                                    return None
                    
                    # 일자 컬럼을 파싱
                    self.quote_data['일자_parsed'] = self.quote_data['일자'].apply(parse_korean_date)
                    
                    # 파싱 성공한 데이터 확인
                    valid_dates = self.quote_data[self.quote_data['일자_parsed'].notna()]
                    failed_dates = self.quote_data[self.quote_data['일자_parsed'].isna()]
                    
                    logger.debug(f"DataLoader: 날짜 파싱 결과 - 성공: {len(valid_dates)}행, 실패: {len(failed_dates)}행")
                    
                    if not failed_dates.empty:
                        logger.debug(f"DataLoader: 파싱 실패한 날짜 샘플: {failed_dates['일자'].head(3).tolist()}")
                    
                    # 파싱된 날짜를 문자열로 변환하여 일자 컬럼 업데이트
                    self.quote_data['일자'] = self.quote_data['일자_parsed'].dt.strftime('%Y-%m-%d %H:%M:%S')
                    self.quote_data = self.quote_data.drop('일자_parsed', axis=1)
                    
                    logger.debug(f"DataLoader: 날짜 변환 완료 - 샘플: {self.quote_data['일자'].head(3).tolist()}")
                
                if '만료일' in self.quote_data.columns:
                    self.quote_data['만료일'] = pd.to_datetime(self.quote_data['만료일']).dt.date
                
                logger.info(f"DataLoader: 견적서 데이터 강제 재로드 완료 - {len(self.quote_data)}행")
                return self.quote_data
            else:
                logger.error("DataLoader: 견적서 데이터 로드 실패")
                return pd.DataFrame()
                
        except Exception as e:
            logger.error(f"DataLoader: 견적서 데이터 강제 재로드 오류 - {e}")
            return pd.DataFrame()
    
    def is_data_loaded(self):
        """전체 데이터 로딩 완료 여부"""
        return (self.renewal_data is not None and 
                self.customer_data is not None and 
                self.price_data is not None and 
                self.log_data is not None)
    
    def get_loading_status(self):
        """로딩 상태 반환"""
        return {
            'is_loading': self.is_loading,
            'progress': self.loading_progress,
            'status': self.loading_status
        }
    
    def stop_background_loading(self):
        """백그라운드 로딩 중단"""
        self.is_loading = False
        if self.background_thread and self.background_thread.is_alive():
            logger.debug("백그라운드 로딩 중단 요청")
            # 스레드가 자연스럽게 종료되도록 is_loading 플래그만 변경
            # 강제 종료는 하지 않음 (데이터 손실 방지) 