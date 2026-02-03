import json
import os
import threading
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES
from core.data_loader import get_global_auth, safe_api_call

class DataManager:
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(DataManager, cls).__new__(cls)
                    cls._instance.initialized = False
        return cls._instance
    
    def __init__(self):
        if self.initialized:
            return
            
        self.initialized = True
        self.cache_file = 'data_cache.json'
        
        # 데이터 저장소
        self.price_data = []
        self.customer_data = []
        
        # 데이터 로드 (캐시 우선)
        self.load_from_cache()
        
    def load_from_cache(self):
        """로컬 캐시 파일에서 데이터 로드"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.price_data = data.get('price_data', [])
                    self.customer_data = data.get('customer_data', [])
                print(f"DataManager: 캐시에서 데이터 로드 완료 - 가격 {len(self.price_data)}건, 고객 {len(self.customer_data)}건")
                return True
        except Exception as e:
            print(f"DataManager: 캐시 로드 실패: {e}")
        return False
        
    def save_to_cache(self):
        """메모리 데이터를 로컬 캐시 파일에 저장"""
        try:
            data = {
                'price_data': self.price_data,
                'customer_data': self.customer_data
            }
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print("DataManager: 데이터 캐시 저장 완료")
        except Exception as e:
            print(f"DataManager: 캐시 저장 실패: {e}")

    def fetch_price_data(self):
        """Google Sheets에서 가격 데이터 가져오기"""
        try:
            print("DataManager: 가격 데이터 동기화 시작...")
            gc = get_global_auth()
            sh = gc.open_by_key('1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk')
            ws = sh.worksheet('Sheet1')
            
            expected_headers = ['대분류', '소분류', '제품', '상세 제품명', '구분', 'USD', 'EUR', 'DC', '원가', '매입처']
            data = ws.get_all_records(expected_headers=expected_headers)
            print(f"DataManager: 가격 데이터 {len(data)}건 로드 완료")
            return data
        except Exception as e:
            print(f"DataManager: 가격 데이터 동기화 실패: {e}")
            return None

    def fetch_customer_data(self):
        """Google Sheets에서 고객 데이터 가져오기"""
        try:
            print("DataManager: 고객 데이터 동기화 시작...")
            gc = get_global_auth()
            sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
            ws = sh.worksheet('Sheet1')
            
            data = ws.get_all_records()
            print(f"DataManager: 고객 데이터 {len(data)}건 로드 완료")
            return data
        except Exception as e:
            print(f"DataManager: 고객 데이터 동기화 실패: {e}")
            return None

    def sync_data(self, callback=None):
        """백그라운드에서 실행될 동기화 로직"""
        print("DataManager: 데이터 동기화 스레드 시작")
        
        # 가격 데이터 동기화
        price_data = self.fetch_price_data()
        if price_data:
            self.price_data = price_data
            
        # 고객 데이터 동기화
        customer_data = self.fetch_customer_data()
        if customer_data:
            self.customer_data = customer_data
            
        # 캐시 저장
        if price_data or customer_data:
            self.save_to_cache()
            
        print("DataManager: 데이터 동기화 완료")
        
        if callback:
            callback()

    def start_background_sync(self, callback=None):
        """백그라운드 동기화 시작"""
        thread = threading.Thread(target=self.sync_data, args=(callback,))
        thread.daemon = True
        thread.start()
        
    def get_price_data(self):
        """가격 데이터 반환 (없으면 동기 로드 시도)"""
        if not self.price_data:
            print("DataManager: 가격 데이터 없음, 동기 로드 시도")
            data = self.fetch_price_data()
            if data:
                self.price_data = data
                self.save_to_cache()
        return self.price_data
        
    def get_customer_data(self):
        """고객 데이터 반환 (없으면 동기 로드 시도)"""
        if not self.customer_data:
            print("DataManager: 고객 데이터 없음, 동기 로드 시도")
            data = self.fetch_customer_data()
            if data:
                self.customer_data = data
                self.save_to_cache()
        return self.customer_data
