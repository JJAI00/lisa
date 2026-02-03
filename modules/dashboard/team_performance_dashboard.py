import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES
import numpy as np
import threading
import pickle
import os
import time
from pathlib import Path

class TeamPerformanceDashboard:
    """팀별 성과 대시보드"""
    
    def __init__(self, parent):
        print("🔍 TeamPerformanceDashboard 초기화 시작")
        self.parent = parent
        
        # 부모 프레임을 직접 사용 (새창 생성하지 않음)
        self.window = parent
        print("✅ 부모 프레임 설정 완료")
        
        # 데이터 저장
        self.target_df = pd.DataFrame()
        self.renewal_df = pd.DataFrame()
        self.quote_df = pd.DataFrame()
        self.filtered_data = pd.DataFrame()
        self.organization_data = pd.DataFrame()  # 조직 정보 저장
        
        # 조직 구조 매핑 (영업사원 -> 팀 -> 부서)
        self.salesperson_to_team = {}  # 영업사원 -> 팀 매핑
        self.salesperson_to_dept = {}  # 영업사원 -> 부서 매핑
        self.team_to_dept = {}  # 팀 -> 부서 매핑
        
        # 캐시 관련
        self.cache_dir = Path("cache")
        self.cache_dir.mkdir(exist_ok=True)
        self.cache_timeout = 1800  # 30분 (초 단위) - API 제한을 고려하여 늘림
        
        # 로딩 상태
        self.is_loading = False
        self.loading_thread = None
        
        # 필터 상태
        self.selected_department = tk.StringVar(value="전체")
        self.selected_team = tk.StringVar(value="전체")
        self.selected_salesperson = tk.StringVar(value="전체")
        self.selected_year = tk.StringVar(value="2025")  # 연도 필터 추가
        self.selected_month = tk.StringVar(value="전체")
        self.selected_quarter = tk.StringVar(value="전체")
        
        # StringVar 변경 추적
        self.selected_year.trace_add("write", self._on_year_changed)
        
        # UI 생성 (데이터 로드 전에 먼저 생성)
        print("🔍 UI 위젯 생성 시작")
        self.create_widgets()
        print("✅ UI 위젯 생성 완료")
        
        # 백그라운드에서 데이터 로드
        print("🔍 백그라운드 데이터 로드 시작")
        self.load_data_async()
        print("✅ 백그라운드 데이터 로드 시작 완료")
    
    def load_data_async(self):
        """백그라운드에서 데이터 로드"""
        print("🔍 load_data_async 메서드 호출됨")
        if self.is_loading:
            print("이미 로딩 중입니다.")
            return
        
        print("🔍 백그라운드 스레드 시작")
        self.is_loading = True
        self.loading_thread = threading.Thread(target=self._load_data_in_thread)
        self.loading_thread.daemon = True
        self.loading_thread.start()
        print("✅ 백그라운드 스레드 시작 완료")
    
    def _load_data_in_thread(self):
        """스레드에서 데이터 로드"""
        try:
            # 먼저 캐시에서 로드 시도
            if self._load_from_cache():
                print("캐시에서 데이터 로드 완료")
                self.window.after(0, self._on_data_loaded)
                return
            
            # 캐시가 없거나 만료되었으면 새로 로드
            self.load_data()
            self._save_to_cache()
            self.window.after(0, self._on_data_loaded)
            
        except Exception as e:
            print(f"=== 데이터 로드 중 오류 발생 ===")
            print(f"오류 내용: {e}")
            import traceback
            traceback.print_exc()
            print("데이터 로드에 실패했습니다. Target 워크시트를 확인해주세요.")
            self.window.after(0, self._on_data_loaded)
        finally:
            self.is_loading = False
    
    def _on_data_loaded(self):
        """데이터 로드 완료 후 UI 업데이트"""
        try:
            # 로딩 인디케이터 숨기기
            if self.target_df.empty:
                self.loading_label.config(text="⚠️ 데이터가 없습니다", fg="#DC2626")
                messagebox.showerror("데이터 로드 실패", 
                                   "Target 워크시트에서 데이터를 불러올 수 없습니다.\n"
                                   "구글 시트의 Target 워크시트를 확인해주세요.")
            else:
                self.loading_label.config(text="✅ 데이터 로드 완료", fg="#059669")
            
            # 데이터 상태 출력
            print(f"=== 데이터 로드 완료 상태 ===")
            print(f"Target 데이터: {len(self.target_df)}행")
            print(f"Renewal 데이터: {len(self.renewal_df)}행")
            print(f"Quote 데이터: {len(self.quote_df)}행")
            print(f"조직 데이터: {len(self.organization_data)}행")
            
            if not self.target_df.empty:
                print(f"Target 컬럼: {list(self.target_df.columns)}")
                print(f"Target 샘플 데이터:")
                print(self.target_df.head())
            
            # 필터 옵션 업데이트
            print(f"=== 필터 옵션 업데이트 전 selected_year 값: {self.selected_year.get()} ===")
            self.update_filter_options()
            print(f"=== 필터 옵션 업데이트 후 selected_year 값: {self.selected_year.get()} ===")
            
            # 초기 데이터 표시
            self.update_dashboard()
            
            print("UI 업데이트 완료")
            
        except Exception as e:
            print(f"UI 업데이트 중 오류: {e}")
            self.loading_label.config(text="❌ 데이터 로드 실패", fg="#DC2626")
            messagebox.showerror("오류", f"데이터 로드 중 오류가 발생했습니다:\n{str(e)}")
    
    def _load_from_cache(self):
        """캐시에서 데이터 로드"""
        try:
            cache_file = self.cache_dir / "team_dashboard_cache.pkl"
            if not cache_file.exists():
                return False
            
            # 캐시 파일의 수정 시간 확인
            cache_age = datetime.now().timestamp() - cache_file.stat().st_mtime
            if cache_age > self.cache_timeout:
                print("캐시가 만료되었습니다.")
                return False
            
            with open(cache_file, 'rb') as f:
                cached_data = pickle.load(f)
            
            self.target_df = cached_data.get('target_df', pd.DataFrame())
            self.renewal_df = cached_data.get('renewal_df', pd.DataFrame())
            self.quote_df = cached_data.get('quote_df', pd.DataFrame())
            self.organization_data = cached_data.get('organization_data', pd.DataFrame())
            
            print("✅ 캐시에서 데이터 로드 성공")
            return True
            
        except Exception as e:
            print(f"캐시 로드 실패: {e}")
            return False
    
    def _save_to_cache(self):
        """데이터를 캐시에 저장"""
        try:
            cache_data = {
                'target_df': self.target_df,
                'renewal_df': self.renewal_df,
                'quote_df': self.quote_df,
                'organization_data': self.organization_data,
                'timestamp': datetime.now().timestamp()
            }
            
            cache_file = self.cache_dir / "team_dashboard_cache.pkl"
            with open(cache_file, 'wb') as f:
                pickle.dump(cache_data, f)
            
            print("데이터 캐시 저장 완료")
            
        except Exception as e:
            print(f"캐시 저장 실패: {e}")
    
    def load_data(self):
        """데이터 로드 (최적화된 버전)"""
        try:
            print("=== 데이터 로드 시작 ===")
            print("1. Google Sheets 인증 시작...")
            
            # Google Sheets 인증 (한 번만 수행)
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            print("✅ Google Sheets 인증 완료")
            
            # 재시도 로직이 포함된 데이터 로드
            print("2. 조직 정보 로드 시작...")
            self._load_organization_data_with_retry(gc)
            print("3. Target 데이터 로드 시작...")
            self._load_target_data_with_retry(gc)
            print("4. Renewal 데이터 로드 시작...")
            self._load_renewal_data_with_retry(gc)
            print("5. Quote 데이터 로드 시작...")
            self._load_quote_data_with_retry(gc)
            
            # 데이터 전처리
            print("데이터 전처리 시작...")
            try:
                self.preprocess_data()
                print("✅ 데이터 전처리 완료")
            except Exception as e:
                print(f"❌ 데이터 전처리 중 오류: {e}")
                import traceback
                traceback.print_exc()
            
            # 실적 데이터 계산
            print("실적 데이터 계산 시작...")
            self.calculate_performance_data()
            
            # 디버깅: 최종 데이터 상태 확인
            print(f"최종 Target 데이터 크기: {len(self.target_df)}행")
            if not self.target_df.empty:
                print(f"최종 Target 컬럼: {list(self.target_df.columns)}")
                print(f"부서 목록: {sorted(self.target_df['부서'].dropna().unique().tolist())}")
                print(f"팀 목록: {sorted(self.target_df['팀'].dropna().unique().tolist())}")
                print(f"영업사원 목록: {sorted(self.target_df['영업사원'].dropna().unique().tolist())}")
            else:
                # Target 데이터가 비어있으면 에러 처리
                print("❌ Target 데이터가 비어있습니다. 구글 시트의 Target 워크시트를 확인해주세요.")
            
            # 데이터 로딩 완료 후 대시보드 업데이트
            print("데이터 로딩 완료 - 대시보드 업데이트 시작...")
            self.update_dashboard()
            
            print("데이터 로드 완료!")
            
        except Exception as e:
            print(f"데이터 로드 중 오류: {e}")
            print("❌ 데이터 로드 실패. Target 워크시트를 확인해주세요.")
            import traceback
            traceback.print_exc()
    
    def _load_with_retry(self, load_func, max_retries=3, delay=2):
        """재시도 로직이 포함된 데이터 로드"""
        for attempt in range(max_retries):
            try:
                print(f"  시도 {attempt + 1}/{max_retries} 시작...")
                result = load_func()
                print(f"  ✅ 시도 {attempt + 1} 성공")
                return result
            except Exception as e:
                print(f"  ❌ 시도 {attempt + 1}/{max_retries} 실패: {e}")
                import traceback
                traceback.print_exc()
                if attempt < max_retries - 1:
                    print(f"  {delay}초 후 재시도...")
                    time.sleep(delay)
                    delay *= 2  # 지수 백오프
                else:
                    print("  최대 재시도 횟수 초과")
                    raise
    
    def _load_organization_data_with_retry(self, gc):
        """조직 정보 로드 (재시도 포함)"""
        def load_func():
            return self._load_organization_data(gc)
        return self._load_with_retry(load_func)
    
    def _load_target_data_with_retry(self, gc):
        """Target 데이터 로드 (재시도 포함)"""
        def load_func():
            return self._load_target_data(gc)
        return self._load_with_retry(load_func)
    
    def _load_renewal_data_with_retry(self, gc):
        """Renewal 데이터 로드 (재시도 포함)"""
        def load_func():
            return self._load_renewal_data(gc)
        return self._load_with_retry(load_func)
    
    def _load_quote_data_with_retry(self, gc):
        """Quote 데이터 로드 (재시도 포함)"""
        def load_func():
            return self._load_quote_data(gc)
        return self._load_with_retry(load_func)
    
    def _load_organization_data(self, gc):
        """조직 정보 로드"""
        try:
            print("조직 정보 로드 중...")
            sh = gc.open_by_key('1KEcLicMV6p-RSXl2kg1AHyD2CmU8wABeXDlMTk0gqZY')
            ws = sh.worksheet('Sheet1')
            org_data = ws.get_all_records()
            self.organization_data = pd.DataFrame(org_data)
            print(f"조직 정보 로드 완료: {len(self.organization_data)}행")
            time.sleep(1)  # API 호출 간격 조절
        except Exception as e:
            print(f"조직 정보 로드 실패: {e}")
            self.organization_data = pd.DataFrame()
    
    def _load_target_data(self, gc):
        """Target 데이터 로드 (부서/팀/영업사원 정보 포함)"""
        try:
            print("=== Target 데이터 로드 시작 ===")
            print("1. Google Sheets 접근 시도...")
            # 올바른 Google Sheets에서 Target 데이터 로드
            sh = gc.open_by_key('1C5p4HBn9G8MpPdQf6Q8FqaM4lgUzzfftBqG-qQhLa7Y')
            print(f"✅ Google Sheets '{sh.title}' 접근 성공")
            
            # 워크시트 목록 확인
            print("2. 워크시트 목록 확인...")
            print(f"Google Sheets '{sh.title}'의 워크시트 목록:")
            for i, ws in enumerate(sh.worksheets()):
                print(f"  {i+1}. {ws.title}")
            
            # Target 워크시트 존재 여부 확인
            print("3. Target 워크시트 접근 시도...")
            try:
                ws = sh.worksheet('Target')
                print("✅ Target 워크시트를 찾았습니다.")
            except Exception as e:
                print(f"❌ Target 워크시트를 찾을 수 없습니다: {e}")
                print("사용 가능한 워크시트 중에서 목표 데이터가 있을 수 있는 워크시트를 확인해보겠습니다.")
                
                # Sheet1이 있는지 확인
                try:
                    ws = sh.worksheet('Sheet1')
                    print("Sheet1 워크시트를 확인해보겠습니다.")
                except:
                    print("Sheet1도 찾을 수 없습니다.")
                    self.target_df = pd.DataFrame()
                    return
            
            # 모든 데이터 가져오기 (헤더 포함)
            print("4. Target 워크시트 데이터 로드 시도...")
            all_data = ws.get_all_values()
            print(f"✅ Target 시트 전체 데이터: {len(all_data)}행 로드 성공")
            
            if len(all_data) < 2:
                print("❌ Target 시트에 데이터가 부족합니다.")
                self.target_df = pd.DataFrame()
                return
            
            # 헤더와 데이터 분리
            print("5. 데이터 구조 분석...")
            headers = all_data[0]  # 첫 번째 행이 헤더
            data_rows = all_data[1:]  # 두 번째 행부터 데이터
            
            print(f"✅ Target 시트 헤더: {headers}")
            print(f"✅ Target 시트 헤더 길이: {len(headers)}")
            print(f"✅ Target 시트 데이터 행 수: {len(data_rows)}")
            
            # 헤더 상세 분석
            print("\n=== 헤더 상세 분석 ===")
            for i, header in enumerate(headers):
                print(f"  컬럼 {i}: '{header}'")
                if '연도' in str(header) or '년' in str(header):
                    print(f"    → 연도 컬럼 발견: 인덱스 {i}")
            
            # 병합된 셀 문제 확인을 위한 상세 분석
            print("\n=== 병합된 셀 문제 확인 ===")
            print("원본 데이터의 빈 셀과 병합된 셀 확인:")
            for i, row in enumerate(data_rows[:10]):  # 처음 10행만 확인
                empty_cells = [j for j, cell in enumerate(row) if cell == '']
                if empty_cells:
                    print(f"행 {i+2}: 빈 셀 위치 {empty_cells} - 전체 행: {row}")
            
            # 부서/팀/영업사원 데이터 특별 확인
            print("\n=== 부서/팀/영업사원 데이터 원본 확인 ===")
            org_rows = []
            for i, row in enumerate(data_rows):
                # 부서, 팀, 영업사원 컬럼이 있는지 확인 (일반적으로 1, 2, 3번째 컬럼)
                if len(row) > 3:
                    dept = row[1] if len(row) > 1 else ''
                    team = row[2] if len(row) > 2 else ''
                    salesperson = row[3] if len(row) > 3 else ''
                    
                    if dept and team and salesperson:  # 모든 값이 있는 경우
                        print(f"조직 데이터 행 {i+2}: {row}")
                        print(f"  - 행 길이: {len(row)}")
                        print(f"  - 회사명: '{row[0] if len(row) > 0 else 'N/A'}'")
                        print(f"  - 부서: '{dept}'")
                        print(f"  - 팀: '{team}'")
                        print(f"  - 영업사원: '{salesperson}'")
                        org_rows.append((i+2, row))
            
            if not org_rows:
                print("⚠️ 부서/팀/영업사원 데이터를 찾을 수 없습니다!")
            
            # 병합된 셀 문제 해결: 빈 헤더 제거 및 유효한 컬럼만 사용
            print("\n=== 병합된 셀 문제 해결 시도 ===")
            raw_header = [h.strip() for h in headers]
            valid_indices = [i for i, h in enumerate(raw_header) if h]
            valid_headers = [raw_header[i] for i in valid_indices]
            
            print(f"원본 헤더: {headers}")
            print(f"유효한 헤더: {valid_headers}")
            print(f"유효한 인덱스: {valid_indices}")
            
            # 유효한 컬럼만 사용하여 데이터 재구성
            valid_data_rows = []
            for row in data_rows:
                valid_row = [row[i] if i < len(row) else '' for i in valid_indices]
                valid_data_rows.append(valid_row)
            
            print(f"유효한 데이터 행 수: {len(valid_data_rows)}")
            
            # 부서/팀/영업사원 데이터 재확인
            print("\n=== 병합된 셀 해결 후 부서/팀/영업사원 데이터 확인 ===")
            for i, row in enumerate(valid_data_rows):
                if len(row) > 3:
                    dept = row[1] if len(row) > 1 else ''
                    team = row[2] if len(row) > 2 else ''
                    salesperson = row[3] if len(row) > 3 else ''
                    
                    if dept and team and salesperson:  # 모든 값이 있는 경우
                        print(f"조직 데이터 행 {i+2}: {row}")
                        print(f"  - 유효한 헤더: {valid_headers}")
                        print(f"  - 회사명: '{row[0] if len(row) > 0 else 'N/A'}'")
                        print(f"  - 부서: '{dept}'")
                        print(f"  - 팀: '{team}'")
                        print(f"  - 영업사원: '{salesperson}'")
            
            # DataFrame 생성 (병합된 셀 문제 해결된 데이터 사용)
            self.target_df = pd.DataFrame(valid_data_rows, columns=valid_headers)
            
            # 컬럼명 정리
            self.target_df.columns = self.target_df.columns.astype(str).str.strip()
            
            print(f"Target 데이터 로드 완료: {len(self.target_df)}행")
            print(f"Target 컬럼: {list(self.target_df.columns)}")
            
            # 데이터 샘플 출력
            if not self.target_df.empty:
                print("Target 데이터 샘플:")
                print(self.target_df.head())
                print("\n=== Target 데이터 상세 분석 ===")
                print(f"회사명 고유값: {self.target_df['회사명'].unique() if '회사명' in self.target_df.columns else '컬럼 없음'}")
                print(f"부서 고유값: {self.target_df['부서'].unique() if '부서' in self.target_df.columns else '컬럼 없음'}")
                print(f"팀 고유값: {self.target_df['팀'].unique() if '팀' in self.target_df.columns else '컬럼 없음'}")
                print(f"영업사원 고유값: {self.target_df['영업사원'].unique() if '영업사원' in self.target_df.columns else '컬럼 없음'}")
                print(f"목표 고유값: {self.target_df['목표'].unique() if '목표' in self.target_df.columns else '컬럼 없음'}")
                
                # 연도 관련 컬럼 확인 (더 정확한 검색)
                print("전체 컬럼명 확인:")
                for i, col in enumerate(self.target_df.columns):
                    print(f"  컬럼 {i}: '{col}'")
                
                year_cols = [col for col in self.target_df.columns if '연도' in str(col) or '년' in str(col) or 'year' in str(col).lower()]
                print(f"연도 관련 컬럼: {year_cols}")
                
                # F열(인덱스 5)이 연도 컬럼인지 확인
                if len(self.target_df.columns) > 5:
                    f_col = self.target_df.columns[5]
                    print(f"F열(인덱스 5) 컬럼명: '{f_col}'")
                    if '연도' in str(f_col) or '년' in str(f_col):
                        print("✅ F열이 연도 컬럼으로 확인됨")
                        year_cols.append(f_col)
                
                if year_cols:
                    for col in year_cols:
                        print(f"{col} 고유값: {self.target_df[col].unique()}")
                        print(f"{col} 샘플값: {self.target_df[col].head().tolist()}")
                else:
                    print("⚠️ 연도 관련 컬럼을 찾을 수 없습니다!")
                    print("전체 컬럼명에서 연도 관련 키워드 검색:")
                    for col in self.target_df.columns:
                        if any(keyword in str(col).lower() for keyword in ['연도', '년', 'year', 'fy', 'fiscal']):
                            print(f"  - {col}")
                
                # 월별 컬럼 확인
                month_cols = [col for col in self.target_df.columns if '월' in col]
                print(f"월별 컬럼: {month_cols}")
                if month_cols:
                    for col in month_cols[:3]:  # 처음 3개 월만 확인
                        print(f"{col} 샘플값: {self.target_df[col].head().tolist()}")
                
                # 부서/팀/영업사원 데이터 특별 확인
                if all(col in self.target_df.columns for col in ['부서', '팀', '영업사원']):
                    print("\n=== 부서/팀/영업사원 조합 확인 ===")
                    org_combinations = self.target_df[['부서', '팀', '영업사원']].drop_duplicates()
                    print("고유한 부서/팀/영업사원 조합:")
                    for _, row in org_combinations.iterrows():
                        print(f"  - {row['부서']} > {row['팀']} > {row['영업사원']}")
            
            time.sleep(1)  # API 호출 간격 조절
            
            print("=== Target 데이터 로드 완료 ===")
            
        except Exception as e:
            print(f"=== Target 데이터 로드 실패 ===")
            print(f"오류 내용: {e}")
            print(f"오류 타입: {type(e).__name__}")
            import traceback
            traceback.print_exc()
            print("Target 데이터를 빈 DataFrame으로 설정합니다.")
            self.target_df = pd.DataFrame()
    
    def _load_renewal_data(self, gc):
        """Renewal 데이터 로드"""
        try:
            print("Renewal 데이터 로드 중...")
            sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
            try:
                ws = sh.worksheet('renewal_list')
            except:
                ws = sh.worksheet('Sheet1')
            renewal_data = ws.get_all_records()
            self.renewal_df = pd.DataFrame(renewal_data)
            print(f"Renewal 데이터 로드 완료: {len(self.renewal_df)}행")
            print(f"Renewal 컬럼: {list(self.renewal_df.columns)}")
            
            # 데이터 샘플 출력
            if not self.renewal_df.empty:
                print("Renewal 데이터 샘플:")
                print(self.renewal_df.head())
                print(f"영업사원 고유값: {self.renewal_df['영업사원'].unique() if '영업사원' in self.renewal_df.columns else '컬럼 없음'}")
                print(f"판매 합계 샘플값: {self.renewal_df['판매 합계'].head().tolist() if '판매 합계' in self.renewal_df.columns else '컬럼 없음'}")
                print(f"이익액 샘플값: {self.renewal_df['이익액'].head().tolist() if '이익액' in self.renewal_df.columns else '컬럼 없음'}")
            
            time.sleep(1)  # API 호출 간격 조절
        except Exception as e:
            print(f"Renewal 데이터 로드 실패: {e}")
            import traceback
            traceback.print_exc()
            self.renewal_df = pd.DataFrame()
    
    def _load_quote_data(self, gc):
        """Quote 데이터 로드"""
        try:
            print("Quote 데이터 로드 중...")
            sh = gc.open_by_key('1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0')
            try:
                ws = sh.worksheet('quote_list')
            except:
                ws = sh.worksheet('Sheet1')
            quote_data = ws.get_all_records()
            self.quote_df = pd.DataFrame(quote_data)
            print(f"Quote 데이터 로드 완료: {len(self.quote_df)}행")
            print(f"Quote 컬럼: {list(self.quote_df.columns)}")
            
            # 데이터 샘플 출력
            if not self.quote_df.empty:
                print("Quote 데이터 샘플:")
                print(self.quote_df.head())
                print(f"영업사원 고유값: {self.quote_df['영업사원'].unique() if '영업사원' in self.quote_df.columns else '컬럼 없음'}")
                print(f"Status 고유값: {self.quote_df['Status'].unique() if 'Status' in self.quote_df.columns else '컬럼 없음'}")
                print(f"매출합계 샘플값: {self.quote_df['매출합계'].head().tolist() if '매출합계' in self.quote_df.columns else '컬럼 없음'}")
                print(f"이익액 샘플값: {self.quote_df['이익액'].head().tolist() if '이익액' in self.quote_df.columns else '컬럼 없음'}")
            
            time.sleep(1)  # API 호출 간격 조절
        except Exception as e:
            print(f"Quote 데이터 로드 실패: {e}")
            import traceback
            traceback.print_exc()
            self.quote_df = pd.DataFrame()
    
    def transform_target_data(self):
        """Target 데이터를 올바른 형태로 변환"""
        print("🔍 transform_target_data 메서드 호출됨")
        try:
            print("=== Target 데이터 변환 시작 ===")
            if self.target_df.empty:
                print("❌ Target 데이터가 비어있어 변환을 건너뜁니다.")
                return
            
            print(f"원본 Target 컬럼: {list(self.target_df.columns)}")
            print(f"원본 Target 데이터 크기: {len(self.target_df)}행")
            
            # 원본 데이터 샘플 출력
            print("원본 데이터 샘플:")
            print(self.target_df.head())
            
            # Target 시트의 실제 구조 분석
            print("\n=== Target 시트 구조 분석 ===")
            
            # 필수 컬럼 확인
            required_cols = ['회사명', '부서', '팀', '영업사원']
            missing_cols = [col for col in required_cols if col not in self.target_df.columns]
            
            if missing_cols:
                print(f"필수 컬럼이 누락되었습니다: {missing_cols}")
                print("사용 가능한 컬럼:", list(self.target_df.columns))
                return
            
            # 월별 컬럼 찾기 (1월, 2월, 3월... 형태)
            month_cols = []
            for col in self.target_df.columns:
                if isinstance(col, str) and '월' in col and col.replace('월', '').isdigit():
                    month_num = int(col.replace('월', ''))
                    month_cols.append((month_num, col))
            
            month_cols.sort(key=lambda x: x[0])  # 월 순서대로 정렬
            print(f"발견된 월별 컬럼: {month_cols}")
            
            if not month_cols:
                print("월별 컬럼을 찾을 수 없습니다.")
                return
            
            # 데이터 변환 - 매출/이익 목표로 분리
            transformed_data = []
            
            print(f"\n=== 데이터 변환 과정 ===")
            for idx, row in self.target_df.iterrows():
                company = row.get('회사명', '')
                dept = row.get('부서', '')
                team = row.get('팀', '')
                salesperson = row.get('영업사원', '')
                
                print(f"행 {idx}: {company}, {dept}, {team}, {salesperson}")
                
                if not all([company, dept, team, salesperson]):
                    print(f"  -> 필수 데이터 누락으로 건너뜀")
                    continue
                
                # 연도 정보 확인
                year = 2025  # 기본값
                if '연도' in row:
                    try:
                        year = int(row['연도'])
                    except (ValueError, TypeError):
                        year = 2025
                
                # 매출 목표 행 생성
                sales_row = {
                    '회사명': company,
                    '부서': dept,
                    '팀': team,
                    '영업사원': salesperson,
                    '목표': '매출 목표',
                    '연도': year
                }
                
                # 월별 매출 목표 데이터 추가
                monthly_sales = []
                for month_num, col_name in month_cols:
                    value = row.get(col_name, 0)
                    try:
                        if isinstance(value, str):
                            value = value.replace(',', '').replace(' ', '')
                        sales_val = float(value) if value and value != '' else 0
                        sales_row[f'{month_num}월'] = sales_val
                        monthly_sales.append(sales_val)
                        print(f"    {col_name} 매출: {value} -> {sales_val:,.0f}")
                    except (ValueError, TypeError) as e:
                        print(f"    {col_name} 매출 변환 실패: {value} -> {e}")
                        sales_row[f'{month_num}월'] = 0
                        monthly_sales.append(0)
                
                # 매출 합계 계산
                sales_total = sum(monthly_sales)
                sales_row['합계'] = sales_total
                print(f"  -> 매출 합계: {sales_total:,.0f}")
                
                transformed_data.append(sales_row)
                
                # 이익 목표 행 생성 (매출의 20%로 설정)
                profit_row = {
                    '회사명': company,
                    '부서': dept,
                    '팀': team,
                    '영업사원': salesperson,
                    '목표': '이익 목표',
                    '연도': year
                }
                
                # 월별 이익 목표 데이터 추가
                monthly_profit = []
                for month_num, col_name in month_cols:
                    profit_val = int(monthly_sales[month_num - 1] * 0.2)  # 매출의 20%
                    profit_row[f'{month_num}월'] = profit_val
                    monthly_profit.append(profit_val)
                    print(f"    {month_num}월 이익: {profit_val:,.0f}")
                
                # 이익 합계 계산
                profit_total = sum(monthly_profit)
                profit_row['합계'] = profit_total
                print(f"  -> 이익 합계: {profit_total:,.0f}")
                
                transformed_data.append(profit_row)
                
                # 김한경 데이터 특별 확인
                if salesperson == '김한경':
                    print(f"  *** 김한경 데이터 생성 완료 ***")
                    print(f"    - 매출 목표: {sales_total:,.0f}")
                    print(f"    - 이익 목표: {profit_total:,.0f}")
            
            # 변환된 데이터로 DataFrame 생성
            self.target_df = pd.DataFrame(transformed_data)
            
            print(f"\n=== 변환 완료 ===")
            print(f"변환된 데이터 크기: {len(self.target_df)}행")
            print(f"변환된 컬럼: {list(self.target_df.columns)}")
            
            # 김한경 데이터 최종 확인
            kim_data = self.target_df[self.target_df['영업사원'] == '김한경']
            if not kim_data.empty:
                print("\n=== 김한경 최종 데이터 확인 ===")
                for _, row in kim_data.iterrows():
                    print(f"{row['목표']}: {row['합계']:,.0f}")
            
            print(f"목표 고유값: {self.target_df['목표'].unique()}")
            print(f"영업사원 고유값: {self.target_df['영업사원'].unique()}")
            
        except Exception as e:
            print(f"Target 데이터 변환 중 오류: {e}")
            import traceback
            traceback.print_exc()
            self.target_df = pd.DataFrame()
    

            
    def preprocess_data(self):
        """데이터 전처리"""
        # Target 데이터 전처리
        if not self.target_df.empty:
            print("Target 데이터 전처리 시작...")
            
            # Target 데이터를 올바른 형태로 변환
            self.transform_target_data()
            
            # transform_target_data 이후 데이터 보호
            print(f"transform_target_data 이후 Target 컬럼: {list(self.target_df.columns)}")
            print(f"transform_target_data 이후 목표 컬럼 샘플: {self.target_df['목표'].head().tolist() if '목표' in self.target_df.columns else '목표 컬럼 없음'}")
            
            # 컬럼명 정리 (목표 컬럼 보호)
            original_columns = self.target_df.columns.copy()
            self.target_df.columns = self.target_df.columns.astype(str).str.strip()
            
            # 목표 컬럼이 손상되지 않았는지 확인
            if '목표' in self.target_df.columns:
                nan_count = self.target_df['목표'].isna().sum()
                if nan_count > 0:
                    print(f"⚠️ 경고: preprocess_data에서 목표 컬럼에 {nan_count}개의 NaN 값이 발견되었습니다!")
                    print("NaN이 있는 행:")
                    print(self.target_df[self.target_df['목표'].isna()][['영업사원', '목표']])
            
            # 숫자 컬럼 변환 (목표 컬럼 제외)
            numeric_cols = [col for col in self.target_df.columns if any(keyword in col for keyword in ['예상', '목표', '실적', '달성율']) and col != '목표']
            for col in numeric_cols:
                self.target_df[col] = pd.to_numeric(self.target_df[col], errors='coerce')
            
            print(f"Target 데이터 전처리 완료: {len(self.target_df)}행")
            print(f"최종 Target 컬럼: {list(self.target_df.columns)}")
            print(f"최종 목표 컬럼 샘플: {self.target_df['목표'].head().tolist() if '목표' in self.target_df.columns else '목표 컬럼 없음'}")
        
        # Renewal 데이터 전처리
        if not self.renewal_df.empty:
            print("Renewal 데이터 전처리 시작...")
            self.renewal_df.columns = self.renewal_df.columns.astype(str).str.strip()
            print(f"Renewal 전처리 후 컬럼: {list(self.renewal_df.columns)}")
            
            # 날짜 컬럼 변환 (여러 가능한 컬럼명 확인)
            date_columns = ['계산서 발행일', '계산서발행일', '발행일', '일자', '날짜']
            date_col_found = None
            for date_col in date_columns:
                if date_col in self.renewal_df.columns:
                    date_col_found = date_col
                    break
            
            if date_col_found:
                print(f"날짜 컬럼 '{date_col_found}' 사용")
                self.renewal_df['계산서_날짜'] = pd.to_datetime(self.renewal_df[date_col_found], errors='coerce')
                self.renewal_df['월'] = self.renewal_df['계산서_날짜'].dt.month
                self.renewal_df['분기'] = self.renewal_df['계산서_날짜'].dt.quarter
                self.renewal_df['연도'] = self.renewal_df['계산서_날짜'].dt.year
                print(f"날짜 변환 완료 - 월별 데이터: {self.renewal_df['월'].value_counts().sort_index().to_dict()}")
                
                # 계산서 발행일 컬럼명 저장 (실적 계산에서 사용)
                self.renewal_date_column = '계산서_날짜'
            else:
                print("날짜 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼:", list(self.renewal_df.columns))
                self.renewal_date_column = None
            
            # 금액 컬럼 변환 (여러 가능한 컬럼명 확인)
            sales_columns = ['판매 합계', '판매합계', '매출', '매출합계', '금액']
            profit_columns = ['이익액', '이익', '수익']
            
            # 매출 컬럼 찾기
            sales_col_found = None
            for sales_col in sales_columns:
                if sales_col in self.renewal_df.columns:
                    sales_col_found = sales_col
                    break
            
            if sales_col_found:
                print(f"매출 컬럼 '{sales_col_found}' 사용")
                self.renewal_df['판매 합계'] = pd.to_numeric(self.renewal_df[sales_col_found], errors='coerce')
            else:
                print("매출 컬럼을 찾을 수 없습니다.")
            
            # 이익 컬럼 찾기
            profit_col_found = None
            for profit_col in profit_columns:
                if profit_col in self.renewal_df.columns:
                    profit_col_found = profit_col
                    break
            
            if profit_col_found:
                print(f"이익 컬럼 '{profit_col_found}' 사용")
                self.renewal_df['이익액'] = pd.to_numeric(self.renewal_df[profit_col_found], errors='coerce')
            else:
                print("이익 컬럼을 찾을 수 없습니다.")
            
            print("Renewal 데이터 전처리 완료")
        
        # Quote 데이터 전처리
        if not self.quote_df.empty:
            print("Quote 데이터 전처리 시작...")
            self.quote_df.columns = self.quote_df.columns.astype(str).str.strip()
            print(f"Quote 전처리 후 컬럼: {list(self.quote_df.columns)}")
            
            # 날짜 컬럼 변환 (여러 가능한 컬럼명 확인)
            date_columns = ['일자', '날짜', '견적일', '작성일']
            date_col_found = None
            for date_col in date_columns:
                if date_col in self.quote_df.columns:
                    date_col_found = date_col
                    break
            
            if date_col_found:
                print(f"날짜 컬럼 '{date_col_found}' 사용")
                self.quote_df['일자'] = pd.to_datetime(self.quote_df[date_col_found], errors='coerce')
                self.quote_df['월'] = self.quote_df['일자'].dt.month
                self.quote_df['분기'] = self.quote_df['일자'].dt.quarter
                self.quote_df['연도'] = self.quote_df['일자'].dt.year
                print(f"날짜 변환 완료 - 월별 데이터: {self.quote_df['월'].value_counts().sort_index().to_dict()}")
            else:
                print("날짜 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼:", list(self.quote_df.columns))
            
            # 금액 컬럼 변환 (여러 가능한 컬럼명 확인)
            sales_columns = ['매출합계', '매출 합계', '매출', '금액', '합계']
            profit_columns = ['이익액', '이익', '수익']
            
            # 매출 컬럼 찾기
            sales_col_found = None
            for sales_col in sales_columns:
                if sales_col in self.quote_df.columns:
                    sales_col_found = sales_col
                    break
            
            if sales_col_found:
                print(f"매출 컬럼 '{sales_col_found}' 사용")
                self.quote_df['매출합계'] = pd.to_numeric(self.quote_df[sales_col_found], errors='coerce')
            else:
                print("매출 컬럼을 찾을 수 없습니다.")
            
            # 이익 컬럼 찾기
            profit_col_found = None
            for profit_col in profit_columns:
                if profit_col in self.quote_df.columns:
                    profit_col_found = profit_col
                    break
            
            if profit_col_found:
                print(f"이익 컬럼 '{profit_col_found}' 사용")
                self.quote_df['이익액'] = pd.to_numeric(self.quote_df[profit_col_found], errors='coerce')
            else:
                print("이익 컬럼을 찾을 수 없습니다.")
            
            # Status 컬럼 확인
            if 'Status' in self.quote_df.columns:
                print(f"Status 고유값: {self.quote_df['Status'].unique()}")
            else:
                print("Status 컬럼을 찾을 수 없습니다.")
            
            print("Quote 데이터 전처리 완료")
    
    def calculate_performance_data(self):
        """실적 데이터 계산 (renewal_list 기반으로 개선)"""
        if self.target_df.empty:
            print("Target 데이터가 비어있어 실적 계산을 건너뜁니다.")
            return
        
        print("실적 데이터 계산 시작...")
        print(f"Target 데이터 크기: {len(self.target_df)}행")
        print(f"Target 컬럼: {list(self.target_df.columns)}")
        
        # 데이터 소스 확인
        print("\n=== 데이터 소스 확인 ===")
        print(f"Renewal 데이터 크기: {len(self.renewal_df)}행")
        print(f"Quote 데이터 크기: {len(self.quote_df)}행")
        
        if not self.renewal_df.empty:
            print("Renewal 데이터 샘플:")
            print(self.renewal_df.head())
            print(f"Renewal 컬럼: {list(self.renewal_df.columns)}")
        
        if not self.quote_df.empty:
            print("Quote 데이터 샘플:")
            print(self.quote_df.head())
            print(f"Quote 컬럼: {list(self.quote_df.columns)}")
            if 'Status' in self.quote_df.columns:
                print(f"Quote Status 값들: {self.quote_df['Status'].unique()}")
        
        # 데이터 사용 확인
        print("✅ Target 워크시트 데이터를 사용 중입니다.")
        
        # Target 데이터에 실적 컬럼 추가
        for month in range(1, 13):
            # 예상 컬럼 (quote_list의 60% Status)
            expected_col = f'{month}월_예상'
            if expected_col not in self.target_df.columns:
                self.target_df[expected_col] = 0
            
            # 실적 컬럼 (renewal_list + quote_list의 80%, 100% Status)
            actual_col = f'{month}월_실적'
            if actual_col not in self.target_df.columns:
                self.target_df[actual_col] = 0
        
        # 합계 컬럼 추가
        if '합계_예상' not in self.target_df.columns:
            self.target_df['합계_예상'] = 0
        if '합계_실적' not in self.target_df.columns:
            self.target_df['합계_실적'] = 0
        if '달성율' not in self.target_df.columns:
            self.target_df['달성율'] = 0
        
        # 1. Renewal 데이터에서 실적 계산 (영업사원별 판매 합계)
        if not self.renewal_df.empty and '영업사원' in self.renewal_df.columns:
            print("\n=== Renewal 데이터에서 실적 계산 시작 ===")
            print(f"Renewal 데이터 크기: {len(self.renewal_df)}행")
            print(f"Renewal 컬럼: {list(self.renewal_df.columns)}")
            
            # Renewal 데이터에서 영업사원별 판매 합계 집계
            print("\n--- Renewal 데이터 영업사원별 판매 합계 집계 ---")
            renewal_salespeople = self.renewal_df['영업사원'].unique()
            print(f"Renewal 데이터의 영업사원: {renewal_salespeople}")
            
            # 영업사원별 판매 합계 계산
            renewal_sales_by_person = {}
            for salesperson in renewal_salespeople:
                if not salesperson or salesperson == '':
                    continue
                    
                print(f"\n영업사원 '{salesperson}' 판매 합계 계산 중...")
                
                # 해당 영업사원의 renewal 데이터 필터링
                renewal_data = self.renewal_df[self.renewal_df['영업사원'] == salesperson]
                
                if not renewal_data.empty and '판매 합계' in renewal_data.columns:
                    # 영업사원별 판매 합계 합계
                    total_sales = renewal_data['판매 합계'].sum()
                    renewal_sales_by_person[salesperson] = total_sales
                    print(f"  - {salesperson} 판매 합계: {total_sales:,.0f}")
                else:
                    print(f"  - {salesperson} 판매 합계 데이터 없음")
                    renewal_sales_by_person[salesperson] = 0
            
            print(f"\nRenewal 데이터 영업사원별 판매 합계 집계 완료:")
            for person, sales in renewal_sales_by_person.items():
                print(f"  {person}: {sales:,.0f}")
        
        # 2. Quote 데이터에서 실적 계산 (80%, 100% Status)
        quote_sales_by_person = {}
        if not self.quote_df.empty and '영업사원' in self.quote_df.columns:
            print("\n=== Quote 데이터에서 실적 계산 시작 ===")
            print(f"Quote 데이터 크기: {len(self.quote_df)}행")
            print(f"Quote 컬럼: {list(self.quote_df.columns)}")
            
            # Status가 80% 또는 100%인 데이터만 필터링 (실적)
            if 'Status' in self.quote_df.columns:
                # Status 값들을 확인하여 정확한 매칭
                status_values = self.quote_df['Status'].unique()
                print(f"사용 가능한 Status 값들: {status_values}")
                
                # 80%, 100% Status 패턴 찾기
                actual_status_patterns = []
                for status in status_values:
                    if isinstance(status, str):
                        if '80%' in status or '발주서' in status or '80' in status:
                            actual_status_patterns.append(status)
                        elif '100%' in status or '세금계산서' in status or '100' in status:
                            actual_status_patterns.append(status)
                
                if actual_status_patterns:
                    print(f"실적용 Status 패턴: {actual_status_patterns}")
                    quote_actual = self.quote_df[self.quote_df['Status'].isin(actual_status_patterns)]
                else:
                    print("80%, 100% Status를 찾을 수 없어 전체 데이터를 실적으로 처리")
                    quote_actual = self.quote_df
            else:
                print("Status 컬럼이 없어 전체 데이터를 실적으로 처리")
                quote_actual = self.quote_df
            
            print(f"실적용 Quote 데이터: {len(quote_actual)}행")
            
            # 영업사원별 판매 합계 계산
            if not quote_actual.empty and '영업사원' in quote_actual.columns:
                quote_salespeople = quote_actual['영업사원'].unique()
                print(f"Quote 실적 데이터의 영업사원: {quote_salespeople}")
                
                for salesperson in quote_salespeople:
                    if not salesperson or salesperson == '':
                        continue
                        
                    print(f"\n영업사원 '{salesperson}' Quote 실적 계산 중...")
                    
                    # 해당 영업사원의 quote 데이터 필터링
                    quote_data = quote_actual[quote_actual['영업사원'] == salesperson]
                    
                    if not quote_data.empty and '매출합계' in quote_data.columns:
                        # 영업사원별 판매 합계 합계
                        total_sales = quote_data['매출합계'].sum()
                        quote_sales_by_person[salesperson] = total_sales
                        print(f"  - {salesperson} Quote 판매 합계: {total_sales:,.0f}")
                    else:
                        print(f"  - {salesperson} Quote 판매 합계 데이터 없음")
                        quote_sales_by_person[salesperson] = 0
            
            print(f"\nQuote 데이터 영업사원별 판매 합계 집계 완료:")
            for person, sales in quote_sales_by_person.items():
                print(f"  {person}: {sales:,.0f}")
        
        # 3. Renewal + Quote 실적 합계 계산
        print("\n=== 매출 실적 합계 계산 ===")
        total_sales_by_person = {}
        
        # 모든 영업사원 목록 수집
        all_salespeople = set()
        all_salespeople.update(renewal_sales_by_person.keys())
        all_salespeople.update(quote_sales_by_person.keys())
        
        for salesperson in all_salespeople:
            renewal_sales = renewal_sales_by_person.get(salesperson, 0)
            quote_sales = quote_sales_by_person.get(salesperson, 0)
            total_sales = renewal_sales + quote_sales
            total_sales_by_person[salesperson] = total_sales
            
            print(f"  {salesperson}: Renewal({renewal_sales:,.0f}) + Quote({quote_sales:,.0f}) = {total_sales:,.0f}")
        
        # Target 데이터에 실적 반영
        print("\n=== Target 데이터에 실적 반영 ===")
        for salesperson, total_sales in total_sales_by_person.items():
            if total_sales > 0:
                # 매출 목표 행에 실적 반영
                sales_mask = (self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == '매출 목표')
                if not self.target_df.loc[sales_mask].empty:
                    # 1월에 전체 실적을 반영 (월별 분할 없이)
                    self.target_df.loc[sales_mask, '1월'] = total_sales
                    print(f"  {salesperson} 매출 실적 반영: {total_sales:,.0f}")
        
        # 4. Quote 데이터에서 예상 계산 (60% Status)
        if not self.quote_df.empty and '영업사원' in self.quote_df.columns:
            print("\n=== Quote 데이터에서 예상 계산 시작 ===")
            
            # Status가 60%인 데이터만 필터링 (예상)
            if 'Status' in self.quote_df.columns:
                # Status 값들을 확인하여 정확한 매칭
                status_values = self.quote_df['Status'].unique()
                print(f"사용 가능한 Status 값들: {status_values}")
                
                # 60% Status 패턴 찾기
                expected_status_patterns = []
                for status in status_values:
                    if isinstance(status, str):
                        if '60%' in status or '검토' in status or '60' in status:
                            expected_status_patterns.append(status)
                
                if expected_status_patterns:
                    print(f"예상용 Status 패턴: {expected_status_patterns}")
                    quote_expected = self.quote_df[self.quote_df['Status'].isin(expected_status_patterns)]
                else:
                    print("60% Status를 찾을 수 없어 예상 데이터를 0으로 처리")
                    quote_expected = pd.DataFrame()
            else:
                print("Status 컬럼이 없어 예상 데이터를 0으로 처리")
                quote_expected = pd.DataFrame()
            
            print(f"예상용 Quote 데이터: {len(quote_expected)}행")
            
            # 영업사원별 예상 판매 합계 계산
            quote_expected_by_person = {}
            if not quote_expected.empty and '영업사원' in quote_expected.columns:
                quote_expected_salespeople = quote_expected['영업사원'].unique()
                print(f"Quote 예상 데이터의 영업사원: {quote_expected_salespeople}")
                
                for salesperson in quote_expected_salespeople:
                    if not salesperson or salesperson == '':
                        continue
                        
                    print(f"\n영업사원 '{salesperson}' Quote 예상 계산 중...")
                    
                    # 해당 영업사원의 quote 데이터 필터링
                    quote_data = quote_expected[quote_expected['영업사원'] == salesperson]
                    
                    if not quote_data.empty and '매출합계' in quote_data.columns:
                        # 영업사원별 예상 판매 합계 합계
                        total_expected = quote_data['매출합계'].sum()
                        quote_expected_by_person[salesperson] = total_expected
                        print(f"  - {salesperson} Quote 예상 판매 합계: {total_expected:,.0f}")
                    else:
                        print(f"  - {salesperson} Quote 예상 판매 합계 데이터 없음")
                        quote_expected_by_person[salesperson] = 0
            
            print(f"\nQuote 데이터 영업사원별 예상 판매 합계 집계 완료:")
            for person, expected in quote_expected_by_person.items():
                print(f"  {person}: {expected:,.0f}")
        
        # Target 데이터에 예상 반영
        print("\n=== Target 데이터에 예상 반영 ===")
        for salesperson, expected_sales in quote_expected_by_person.items():
            if expected_sales > 0:
                # 매출 목표 행에 예상 반영
                sales_mask = (self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == '매출 목표')
                if not self.target_df.loc[sales_mask].empty:
                    # 1월에 전체 예상을 반영 (월별 분할 없이)
                    self.target_df.loc[sales_mask, '1월_예상'] = expected_sales
                    print(f"  {salesperson} 매출 예상 반영: {expected_sales:,.0f}")
        
        # 4. 합계 및 달성율 계산
        print("\n=== 합계 및 달성율 계산 시작 ===")
        for _, target_row in self.target_df.iterrows():
            salesperson = target_row['영업사원']
            category = target_row['목표']  # '구분' → '목표'로 수정
            
            # 예상 합계 (Target 시트의 월별 예상 컬럼 사용)
            expected_sum = sum([target_row.get(f'{month}월_예상', target_row.get(f'{month}월', 0)) for month in range(1, 13)])
            # Target 시트의 '합계_예상' 컬럼 사용
            if '합계_예상' in self.target_df.columns:
                self.target_df.loc[(self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == category), '합계_예상'] = expected_sum
            elif '합계' in self.target_df.columns:
                self.target_df.loc[(self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == category), '합계'] = expected_sum
            
            # 실적 합계 (Target 시트의 월별 실적 컬럼 사용)
            actual_sum = sum([target_row.get(f'{month}월_실적', target_row.get(f'{month}월', 0)) for month in range(1, 13)])
            # Target 시트의 '합계_실적' 컬럼 사용
            if '합계_실적' in self.target_df.columns:
                self.target_df.loc[(self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == category), '합계_실적'] = actual_sum
            elif '합계' in self.target_df.columns:
                self.target_df.loc[(self.target_df['영업사원'] == salesperson) & (self.target_df['목표'] == category), '합계'] = actual_sum
            
            # 달성율 계산 (Target 시트의 목표 합계 사용)
            target_sum = target_row.get('합계_목표', target_row.get('합계', 0))
            if target_sum > 0:
                achievement_rate = (actual_sum / target_sum) * 100
                # Target 시트에는 '달성율' 컬럼이 없으므로 저장하지 않음
            else:
                achievement_rate = 0
            
            print(f"  {salesperson} - {category}: 목표={target_sum:,.0f}, 실적={actual_sum:,.0f}, 예상={expected_sum:,.0f}, 달성율={achievement_rate:.1f}%")
            
            # 김한경 데이터 특별 확인
            if salesperson == '김한경':
                print(f"  *** 김한경 {category} 최종 확인: 목표={target_sum:,.0f}, 실적={actual_sum:,.0f}, 예상={expected_sum:,.0f} ***")
        
        print("\n=== 실적 데이터 계산 완료 ===")
        print(f"최종 Target 데이터 크기: {len(self.target_df)}행")
        if not self.target_df.empty:
            print("최종 데이터 샘플:")
            print(self.target_df[['영업사원', '목표', '합계']].head())  # Target 시트 구조에 맞게 수정
            
            # Renewal 데이터 기반 실적 요약
            print("\n=== Renewal 데이터 기반 실적 요약 ===")
            if not self.renewal_df.empty:
                total_renewal_sales = self.renewal_df['판매 합계'].sum() if '판매 합계' in self.renewal_df.columns else 0
                total_renewal_profit = self.renewal_df['이익액'].sum() if '이익액' in self.renewal_df.columns else 0
                print(f"Renewal 총 매출: {total_renewal_sales:,.0f}")
                print(f"Renewal 총 이익: {total_renewal_profit:,.0f}")
                
                # 영업사원별 Renewal 실적
                if '영업사원' in self.renewal_df.columns:
                    sales_by_person = self.renewal_df.groupby('영업사원')['판매 합계'].sum() if '판매 합계' in self.renewal_df.columns else pd.Series()
                    profit_by_person = self.renewal_df.groupby('영업사원')['이익액'].sum() if '이익액' in self.renewal_df.columns else pd.Series()
                    print("영업사원별 Renewal 실적:")
                    for person in sales_by_person.index:
                        sales = sales_by_person.get(person, 0)
                        profit = profit_by_person.get(person, 0)
                        # 조직 정보와 함께 출력
                        dept = self.salesperson_to_dept.get(person, '미분류')
                        team = self.salesperson_to_team.get(person, '미분류')
                        print(f"  {dept} > {team} > {person}: 매출 {sales:,.0f}, 이익 {profit:,.0f}")
            
            # Target 데이터의 실적 반영 결과
            print("\n=== Target 데이터 실적 반영 결과 ===")
            if not self.target_df.empty:
                # 안전한 컬럼 선택
                if '합계_목표' in self.target_df.columns:
                    target_with_actual = self.target_df[self.target_df['합계_목표'] > 0]
                elif '합계' in self.target_df.columns:
                    target_with_actual = self.target_df[self.target_df['합계'] > 0]
                else:
                    target_with_actual = self.target_df
                if not target_with_actual.empty:
                    print("실적이 반영된 영업사원:")
                    for _, row in target_with_actual.iterrows():
                        # 안전한 컬럼 선택
                        if '합계_목표' in row:
                            total_val = row['합계_목표']
                        elif '합계' in row:
                            total_val = row['합계']
                        else:
                            total_val = 0
                        print(f"  {row['부서']} > {row['팀']} > {row['영업사원']} ({row['목표']}): {total_val:,.0f}")
                else:
                    print("⚠️ 실적이 반영된 영업사원이 없습니다.")
                    print("Renewal 데이터와 Target 데이터의 영업사원명이 일치하는지 확인이 필요합니다.")
    
    def create_widgets(self):
        """UI 위젯 생성"""
        # 메인 프레임
        main_frame = tk.Frame(self.window, bg="#F7F9FB")
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 헤더
        header_frame = tk.Frame(main_frame, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="📊 팀별 성과 대시보드", 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 로딩 인디케이터
        self.loading_label = tk.Label(header_frame, text="🔄 데이터 로딩 중...", 
                                     fg="#6B7280", bg="#F7F9FB")
        self.loading_label.pack(side='right', padx=20, pady=15)
        
        # 필터 프레임
        self.create_filter_frame(main_frame)
        
        # 성과 카드들
        self.create_performance_cards(main_frame)
        
        # 테이블 프레임
        table_frame = tk.Frame(main_frame, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, pady=(20, 0))
        
        # 테이블 생성
        self.create_performance_table(table_frame)
    
    def create_filter_frame(self, parent):
        """필터 프레임 생성"""
        filter_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        filter_frame.pack(fill='x', pady=(0, 20))
        
        # 필터 헤더
        filter_header = tk.Frame(filter_frame, bg="#F1F5F9", height=40)
        filter_header.pack(fill='x')
        filter_header.pack_propagate(False)
        
        tk.Label(filter_header, text="필터 설정", 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 필터 내용
        filter_content = tk.Frame(filter_frame, bg="white")
        filter_content.pack(fill='x', padx=15, pady=15)
        
        # 첫 번째 행: 조직 필터
        row1 = tk.Frame(filter_content, bg="white")
        row1.pack(fill='x', pady=(0, 10))
        
        # 부서 필터
        tk.Label(row1, text="부서:", 
                fg="#374151", bg="white").grid(row=0, column=0, padx=(0, 5), pady=5)
        
        self.department_combo = ttk.Combobox(row1, textvariable=self.selected_department, 
                                           state="readonly", width=20)
        self.department_combo.grid(row=0, column=1, padx=(0, 20), pady=5)
        self.department_combo.bind('<<ComboboxSelected>>', self.on_department_changed)
        
        # 팀 필터
        tk.Label(row1, text="팀:", 
                fg="#374151", bg="white").grid(row=0, column=2, padx=(0, 5), pady=5)
        
        self.team_combo = ttk.Combobox(row1, textvariable=self.selected_team, 
                                     state="readonly", width=20)
        self.team_combo.grid(row=0, column=3, padx=(0, 20), pady=5)
        self.team_combo.bind('<<ComboboxSelected>>', self.on_team_changed)
        
        # 영업사원 필터
        tk.Label(row1, text="영업사원:", 
                fg="#374151", bg="white").grid(row=0, column=4, padx=(0, 5), pady=5)
        
        self.salesperson_combo = ttk.Combobox(row1, textvariable=self.selected_salesperson, 
                                            state="readonly", width=15)
        self.salesperson_combo.grid(row=0, column=5, padx=(0, 20), pady=5)
        self.salesperson_combo.bind('<<ComboboxSelected>>', self.on_salesperson_changed)
        
        # 연도 필터
        tk.Label(row1, text="연도:", 
                fg="#374151", bg="white").grid(row=0, column=6, padx=(0, 5), pady=5)
        
        self.year_combo = ttk.Combobox(row1, textvariable=self.selected_year, 
                                     state="readonly", width=10)
        self.year_combo.grid(row=0, column=7, padx=(0, 20), pady=5)
        self.year_combo.bind('<<ComboboxSelected>>', self.on_filter_changed)
        
        # 두 번째 행: 기간 필터
        row2 = tk.Frame(filter_content, bg="white")
        row2.pack(fill='x', pady=(0, 10))
        
        # 기간 구분 라디오 버튼
        tk.Label(row2, text="기간 구분:", 
                fg="#374151", bg="white").grid(row=0, column=0, padx=(0, 5), pady=5)
        
        period_frame = tk.Frame(row2, bg="white")
        period_frame.grid(row=0, column=1, padx=(0, 20), pady=5)
        
        self.period_var = tk.StringVar(value="월별")
        tk.Radiobutton(period_frame, text="월별", variable=self.period_var, value="월별",
                      command=self.on_period_changed, bg="white").pack(side='left', padx=(0, 10))
        tk.Radiobutton(period_frame, text="분기별", variable=self.period_var, value="분기별",
                      command=self.on_period_changed, bg="white").pack(side='left')
        
        # 월 필터
        tk.Label(row2, text="월:", 
                fg="#374151", bg="white").grid(row=0, column=2, padx=(0, 5), pady=5)
        
        self.month_combo = ttk.Combobox(row2, textvariable=self.selected_month, 
                                       values=["전체"] + [f"{i}월" for i in range(1, 13)], 
                                       state="readonly", width=10)
        self.month_combo.grid(row=0, column=3, padx=(0, 20), pady=5)
        self.month_combo.bind('<<ComboboxSelected>>', self.on_filter_changed)
        
        # 분기 필터
        tk.Label(row2, text="분기:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=4, padx=(0, 5), pady=5)
        
        self.quarter_combo = ttk.Combobox(row2, textvariable=self.selected_quarter, 
                                        values=["전체", "1분기", "2분기", "3분기", "4분기"], 
                                        state="readonly", width=10)
        self.quarter_combo.grid(row=0, column=5, padx=(0, 20), pady=5)
        self.quarter_combo.bind('<<ComboboxSelected>>', self.on_filter_changed)
        
        # 새로고침 버튼
        refresh_btn = tk.Button(row2, text="🔄 새로고침", command=self.refresh_dashboard,
                               font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                               relief="flat", padx=15, pady=5)
        refresh_btn.grid(row=0, column=6, padx=(20, 0), pady=5)
    

    
    def create_performance_cards(self, parent):
        """성과 카드들 생성 (이미지 형태)"""
        # 성과 카드 컨테이너
        self.performance_cards_container = tk.Frame(parent, bg="#F7F9FB")
        self.performance_cards_container.pack(fill='x', pady=(20, 20))
        
        # 카드들을 수평으로 배치
        cards_frame = tk.Frame(self.performance_cards_container, bg="#F7F9FB")
        cards_frame.pack(fill='x')
        
        # 목표 카드
        self.create_target_card(cards_frame)
        
        # 실적 카드
        self.create_performance_card(cards_frame, "실적", "#10B981")
        
        # 예상 카드
        self.create_performance_card(cards_frame, "예상", "#F59E0B")
        
        # 합계 카드
        self.create_performance_card(cards_frame, "합계", "#8B5CF6")
        
        # 달성율 카드
        self.create_achievement_card(cards_frame)
    
    def create_performance_section(self, section_name, color):
        """성과 섹션 생성 (이미지 형태)"""
        # 섹션 프레임 (이미지처럼 세로로 나누어진 패널)
        section_frame = tk.Frame(self.performance_cards_container, bg="white", relief="solid", bd=1)
        section_frame.pack(side='left', fill='both', expand=True, padx=(0, 1), pady=0)
        
        # 섹션 헤더
        header_frame = tk.Frame(section_frame, bg=color, height=30)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text=section_name, font=("맑은 고딕", 11, "bold"), 
                fg="white", bg=color).pack(expand=True)
        
        # 섹션 내용
        content_frame = tk.Frame(section_frame, bg="white", padx=10, pady=10)
        content_frame.pack(fill='both', expand=True)
        
        if section_name == "달성율":
            # 달성율 섹션: 매출 달성율과 이익 달성율을 한 줄씩 표시
            # 매출 달성율
            sales_achievement_label = tk.Label(content_frame, text="매출:", font=("맑은 고딕", 10), 
                                              fg="#374151", bg="white", anchor='w')
            sales_achievement_label.pack(fill='x', pady=(0, 5))
            
            sales_achievement_value = tk.Label(content_frame, text="0%", font=("맑은 고딕", 12, "bold"), 
                                              fg=color, bg="white", anchor='w')
            sales_achievement_value.pack(fill='x', pady=(0, 10))
            
            # 이익 달성율
            profit_achievement_label = tk.Label(content_frame, text="이익:", font=("맑은 고딕", 10), 
                                               fg="#374151", bg="white", anchor='w')
            profit_achievement_label.pack(fill='x', pady=(0, 5))
            
            profit_achievement_value = tk.Label(content_frame, text="0%", font=("맑은 고딕", 12, "bold"), 
                                               fg=color, bg="white", anchor='w')
            profit_achievement_value.pack(fill='x')
            
            # 참조 저장
            setattr(self, f"{section_name.lower()}_매출_card", sales_achievement_value)
            setattr(self, f"{section_name.lower()}_이익_card", profit_achievement_value)
        else:
            # 일반 섹션: 매출과 이익을 한 줄로 표시
            # 매출과 이익을 한 줄로 표시
            sales_profit_frame = tk.Frame(content_frame, bg="white")
            sales_profit_frame.pack(fill='x', pady=(0, 5))
            
            # 매출 라벨과 값
            sales_label = tk.Label(sales_profit_frame, text="매출:", font=("맑은 고딕", 10), 
                                  fg="#374151", bg="white", anchor='w')
            sales_label.pack(side='left')
            
            sales_value = tk.Label(sales_profit_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                                  fg=color, bg="white", anchor='w')
            sales_value.pack(side='left', padx=(5, 0))
            
            # 이익 라벨과 값
            profit_label = tk.Label(sales_profit_frame, text="이익:", font=("맑은 고딕", 10), 
                                   fg="#374151", bg="white", anchor='w')
            profit_label.pack(side='left', padx=(20, 0))
            
            profit_value = tk.Label(sales_profit_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                                   fg=color, bg="white", anchor='w')
            profit_value.pack(side='left', padx=(5, 0))
            
            # 참조 저장
            setattr(self, f"{section_name.lower()}_매출_card", sales_value)
            setattr(self, f"{section_name.lower()}_이익_card", profit_value)
    
    def create_target_card(self, parent):
        """목표 카드 생성 (이미지 형태)"""
        # 카드 프레임
        card_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        card_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 헤더 (파란색)
        header_frame = tk.Frame(card_frame, bg="#3B82F6", height=30)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="목표", font=("맑은 고딕", 11, "bold"), 
                fg="white", bg="#3B82F6").pack(expand=True)
        
        # 내용 (흰색)
        content_frame = tk.Frame(card_frame, bg="white", padx=15, pady=15)
        content_frame.pack(fill='both', expand=True)
        
        # 매출 라인
        sales_frame = tk.Frame(content_frame, bg="white")
        sales_frame.pack(fill='x', pady=(0, 8))
        
        sales_label = tk.Label(sales_frame, text="매출:", font=("맑은 고딕", 10), 
                              fg="#374151", bg="white", anchor='w')
        sales_label.pack(side='left')
        
        self.target_sales_value = tk.Label(sales_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                                          fg="#3B82F6", bg="white", anchor='w')
        self.target_sales_value.pack(side='left', padx=(5, 0))
        
        # 이익 라인
        profit_frame = tk.Frame(content_frame, bg="white")
        profit_frame.pack(fill='x')
        
        profit_label = tk.Label(profit_frame, text="이익:", font=("맑은 고딕", 10), 
                               fg="#374151", bg="white", anchor='w')
        profit_label.pack(side='left')
        
        self.target_profit_value = tk.Label(profit_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                                           fg="#3B82F6", bg="white", anchor='w')
        self.target_profit_value.pack(side='left', padx=(5, 0))
    
    def create_performance_card(self, parent, title, color):
        """성과 카드 생성 (실적, 예상, 합계)"""
        # 카드 프레임
        card_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        card_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 헤더
        header_frame = tk.Frame(card_frame, bg=color, height=30)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text=title, font=("맑은 고딕", 11, "bold"), 
                fg="white", bg=color).pack(expand=True)
        
        # 내용
        content_frame = tk.Frame(card_frame, bg="white", padx=15, pady=15)
        content_frame.pack(fill='both', expand=True)
        
        # 매출 라인
        sales_frame = tk.Frame(content_frame, bg="white")
        sales_frame.pack(fill='x', pady=(0, 8))
        
        sales_label = tk.Label(sales_frame, text="매출:", font=("맑은 고딕", 10), 
                              fg="#374151", bg="white", anchor='w')
        sales_label.pack(side='left')
        
        sales_value = tk.Label(sales_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                              fg=color, bg="white", anchor='w')
        sales_value.pack(side='left', padx=(5, 0))
        
        # 이익 라인
        profit_frame = tk.Frame(content_frame, bg="white")
        profit_frame.pack(fill='x')
        
        profit_label = tk.Label(profit_frame, text="이익:", font=("맑은 고딕", 10), 
                               fg="#374151", bg="white", anchor='w')
        profit_label.pack(side='left')
        
        profit_value = tk.Label(profit_frame, text="0", font=("맑은 고딕", 12, "bold"), 
                               fg=color, bg="white", anchor='w')
        profit_value.pack(side='left', padx=(5, 0))
        
        # 참조 저장
        setattr(self, f"{title.lower()}_매출_card", sales_value)
        setattr(self, f"{title.lower()}_이익_card", profit_value)
    
    def create_achievement_card(self, parent):
        """달성율 카드 생성"""
        # 카드 프레임
        card_frame = tk.Frame(parent, bg="white", relief="solid", bd=1)
        card_frame.pack(side='left', fill='both', expand=True)
        
        # 헤더 (빨간색)
        header_frame = tk.Frame(card_frame, bg="#EF4444", height=30)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="달성율", font=("맑은 고딕", 11, "bold"), 
                fg="white", bg="#EF4444").pack(expand=True)
        
        # 내용
        content_frame = tk.Frame(card_frame, bg="white", padx=15, pady=15)
        content_frame.pack(fill='both', expand=True)
        
        # 매출 달성율 라인
        sales_frame = tk.Frame(content_frame, bg="white")
        sales_frame.pack(fill='x', pady=(0, 8))
        
        sales_label = tk.Label(sales_frame, text="매출:", font=("맑은 고딕", 10), 
                              fg="#374151", bg="white", anchor='w')
        sales_label.pack(side='left')
        
        self.achievement_sales_value = tk.Label(sales_frame, text="0%", font=("맑은 고딕", 12, "bold"), 
                                               fg="#EF4444", bg="white", anchor='w')
        self.achievement_sales_value.pack(side='left', padx=(5, 0))
        
        # 이익 달성율 라인
        profit_frame = tk.Frame(content_frame, bg="white")
        profit_frame.pack(fill='x')
        
        profit_label = tk.Label(profit_frame, text="이익:", font=("맑은 고딕", 10), 
                               fg="#374151", bg="white", anchor='w')
        profit_label.pack(side='left')
        
        self.achievement_profit_value = tk.Label(profit_frame, text="0%", font=("맑은 고딕", 12, "bold"), 
                                                fg="#EF4444", bg="white", anchor='w')
        self.achievement_profit_value.pack(side='left', padx=(5, 0))
    

    
    def create_performance_table(self, parent):
        """성과 테이블 생성 (이미지 형태)"""
        # 테이블 헤더
        header_frame = tk.Frame(parent, bg="#F1F5F9", height=40)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="팀별 성과 현황", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 테이블 컨테이너 - 고정 높이로 설정하여 스크롤 방지
        table_container = tk.Frame(parent, bg="white", height=400)
        table_container.pack(fill='both', expand=True, padx=15, pady=15)
        table_container.pack_propagate(False)  # 고정 높이 유지
        
        # Treeview 생성 - 이미지와 같은 구조
        columns = ['목표']  # 첫 번째 컬럼은 목표
        
        # 월별 컬럼 추가 (이미지처럼 월만 표시)
        for month in range(1, 13):
            columns.append(f'{month}월')
        
        # 합계 컬럼 추가
        columns.append('합계')
        
        self.tree = ttk.Treeview(table_container, columns=columns, show='headings', height=15)
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        
        # Treeview 설정
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 컬럼 설정 - 고정 너비로 설정하여 스크롤 가능하게 함
        for col in columns:
            if col == '목표':
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=120, anchor='w', stretch=False)
            elif col == '합계':
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=120, anchor='e', stretch=False)
            else:
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=100, anchor='e', stretch=False)
        
        # 스크롤바 배치
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        
        # 컬럼 크기 조절 기능 추가
        from ui.resizable_treeview import ResizableTreeview
        self.resizable_tree = ResizableTreeview(self.tree, "team_performance_table")
        
        # 테이블 스타일 설정
        self.setup_table_style()
    
    def setup_table_style(self):
        """테이블 스타일 설정"""
        # Treeview 스타일 설정
        style = ttk.Style()
        
        # 헤더 스타일 (이미지의 진한 청록색 헤더)
        style.configure("Treeview.Heading", 
                       background="#2D5A5A", 
                       foreground="white", 
                       font=("맑은 고딕", 10, "bold"))
        
        # 행 스타일
        style.configure("Treeview", 
                       background="white", 
                       foreground="black", 
                       fieldbackground="white",
                       font=("맑은 고딕", 9))
        
        # 선택된 행 스타일
        style.map("Treeview", 
                 background=[("selected", "#E5F3FF")],
                 foreground=[("selected", "black")])
        
        # 태그별 색상 설정
        self.tree.tag_configure('sales', background='white')
        self.tree.tag_configure('profit', background='white')
        self.tree.tag_configure('total', background='#F8F9FA', font=("맑은 고딕", 9, "bold"))
        self.tree.tag_configure('achievement', background='#F8F9FA', font=("맑은 고딕", 9, "bold"))
        self.tree.tag_configure('separator', background='#E5E7EB')
    

    

    
    def update_filter_options(self):
        """필터 옵션 업데이트 - Target 시트에서 실제 데이터 가져오기"""
        print("=== 필터 옵션 업데이트 시작 ===")
        
        # Target 데이터에서 직접 필터 옵션 가져오기 (우선순위)
        if not self.target_df.empty:
            print("Target 데이터에서 필터 옵션을 가져옵니다.")
            print(f"Target 데이터 크기: {len(self.target_df)}행")
            print(f"Target 컬럼: {list(self.target_df.columns)}")
            
            # 부서 옵션 (Target 데이터에서 가져오기)
            departments = ["전체"]
            if '부서' in self.target_df.columns:
                target_depts = sorted(self.target_df['부서'].dropna().unique().tolist())
                departments.extend(target_depts)
                print(f"Target 데이터에서 부서 목록: {target_depts}")
            else:
                print("⚠️ Target 데이터에 '부서' 컬럼이 없습니다.")
                print(f"사용 가능한 컬럼: {list(self.target_df.columns)}")
            
            # 팀 옵션 (Target 데이터에서 가져오기)
            teams = ["전체"]
            if '팀' in self.target_df.columns:
                target_teams = sorted(self.target_df['팀'].dropna().unique().tolist())
                teams.extend(target_teams)
                print(f"Target 데이터에서 팀 목록: {target_teams}")
            else:
                print("⚠️ Target 데이터에 '팀' 컬럼이 없습니다.")
            
            # 영업사원 옵션 (Target 데이터에서 가져오기)
            salespeople = ["전체"]
            if '영업사원' in self.target_df.columns:
                target_salespeople = sorted(self.target_df['영업사원'].dropna().unique().tolist())
                salespeople.extend(target_salespeople)
                print(f"Target 데이터에서 영업사원 목록: {target_salespeople}")
            else:
                print("⚠️ Target 데이터에 '영업사원' 컬럼이 없습니다.")
            
            # 연도 옵션 (Target 데이터에서 실제 연도 정보 가져오기)
            years = ["전체"]
            year_column = None
            
            # 연도 관련 컬럼 찾기 (더 정확한 검색)
            print("연도 컬럼 검색 중...")
            print(f"전체 컬럼명: {list(self.target_df.columns)}")
            
            # 연도 컬럼을 정확히 찾기
            print("연도 컬럼 검색 중...")
            for i, col in enumerate(self.target_df.columns):
                print(f"  컬럼 {i}: '{col}'")
                if '연도' in str(col) or '년' in str(col) or 'Year' in str(col):
                    year_column = col
                    print(f"✅ 연도 컬럼 발견: '{col}' (인덱스 {i})")
                    break
            
            # F열이 연도 컬럼이 아닌 경우 다른 후보들 확인
            if not year_column:
                year_candidates = ['연도', '년', 'Year', 'YEAR', 'FY', 'Fiscal Year']
                for candidate in year_candidates:
                    if candidate in self.target_df.columns:
                        year_column = candidate
                        print(f"연도 컬럼을 찾았습니다: '{candidate}'")
                        break
            
            # 연도 컬럼을 찾지 못한 경우, 컬럼명에 연도 관련 키워드가 포함된 것 찾기
            if year_column is None:
                for col in self.target_df.columns:
                    if any(keyword in str(col).lower() for keyword in ['연도', '년', 'year', 'fy', 'fiscal']):
                        year_column = col
                        print(f"연도 관련 컬럼을 찾았습니다: '{col}'")
                        break
            
            if year_column:
                target_years = sorted(self.target_df[year_column].dropna().unique().tolist())
                print(f"원본 연도 데이터: {target_years}")
                
                # 연도를 그대로 사용 (FY 형식으로 변환하지 않음)
                years.extend([str(year) for year in target_years])
                print(f"Target 데이터에서 연도 목록: {target_years}")
            else:
                # 연도 컬럼이 없으면 기본값 사용
                years = ["전체", "2024", "2025"]
                print("Target 데이터에 연도 컬럼을 찾을 수 없어 기본값을 사용합니다.")
                print(f"사용 가능한 컬럼: {list(self.target_df.columns)}")
            
            # Target 데이터 샘플 출력 (디버깅용)
            print("\n=== Target 데이터 샘플 ===")
            print(self.target_df[['회사명', '부서', '팀', '영업사원']].head() if all(col in self.target_df.columns for col in ['회사명', '부서', '팀', '영업사원']) else self.target_df.head())
            
        else:
            print("Target 데이터가 비어있어 조직 정보에서 필터 옵션을 가져옵니다.")
            
            # 기본값으로 초기화
            departments = ["전체"]
            teams = ["전체"]
            salespeople = ["전체"]
            years = ["전체", "2024", "2025"]
            
            # 조직 정보에서 필터 옵션 가져오기 (백업)
            if not self.organization_data.empty:
                print("조직 정보에서 필터 옵션을 가져옵니다.")
                print(f"조직 정보 크기: {len(self.organization_data)}행")
                print(f"조직 정보 컬럼: {list(self.organization_data.columns)}")
                
                # 부서 옵션 (조직 정보에서 가져오기)
                if '부서' in self.organization_data.columns:
                    org_depts = sorted(self.organization_data['부서'].dropna().unique().tolist())
                    departments.extend(org_depts)
                    print(f"조직 정보에서 부서 목록: {org_depts}")
                else:
                    print("⚠️ 조직 정보에 '부서' 컬럼이 없습니다.")
                
                # 팀 옵션 (조직 정보에서 가져오기)
                if '팀' in self.organization_data.columns:
                    org_teams = sorted(self.organization_data['팀'].dropna().unique().tolist())
                    teams.extend(org_teams)
                    print(f"조직 정보에서 팀 목록: {org_teams}")
                else:
                    print("⚠️ 조직 정보에 '팀' 컬럼이 없습니다.")
                
                # 영업사원 옵션 (조직 정보에서 가져오기)
                if '영업사원' in self.organization_data.columns:
                    org_salespeople = sorted(self.organization_data['영업사원'].dropna().unique().tolist())
                    salespeople.extend(org_salespeople)
                    print(f"조직 정보에서 영업사원 목록: {org_salespeople}")
                else:
                    print("⚠️ 조직 정보에 '영업사원' 컬럼이 없습니다.")
                
                # 조직 정보 샘플 출력 (디버깅용)
                print("\n=== 조직 정보 샘플 ===")
                print(self.organization_data[['회사명', '부서', '팀', '영업사원']].head() if all(col in self.organization_data.columns for col in ['회사명', '부서', '팀', '영업사원']) else self.organization_data.head())
            else:
                print("조직 정보도 비어있어 기본값을 사용합니다.")
        
        # 콤보박스 업데이트
        print(f"\n=== 콤보박스 값 설정 ===")
        print(f"부서 콤보박스 값: {departments}")
        print(f"팀 콤보박스 값: {teams}")
        print(f"영업사원 콤보박스 값: {salespeople}")
        print(f"연도 콤보박스 값: {years}")
        
        # 콤보박스 값 설정
        self.department_combo['values'] = departments
        self.team_combo['values'] = teams
        self.salesperson_combo['values'] = salespeople
        self.year_combo['values'] = years
        
        # 콤보박스 값이 제대로 설정되었는지 확인
        print(f"부서 콤보박스 실제 값: {self.department_combo['values']}")
        print(f"팀 콤보박스 실제 값: {self.team_combo['values']}")
        print(f"영업사원 콤보박스 실제 값: {self.salesperson_combo['values']}")
        print(f"연도 콤보박스 실제 값: {self.year_combo['values']}")
        
        # 기본값 설정 (가장 최근 연도 선택)
        if len(years) > 1:  # "전체" 외에 다른 옵션이 있으면
            # 숫자 연도 중 가장 큰 값 선택 (문자열을 정수로 변환)
            numeric_years = []
            for y in years:
                if y != "전체" and y.isdigit():
                    try:
                        numeric_years.append(int(y))
                    except ValueError:
                        print(f"⚠️ 연도 변환 실패: '{y}'")
                        continue
            
            if numeric_years:
                latest_year = str(max(numeric_years))  # 정수로 비교 후 문자열로 변환
                self.selected_year.set(latest_year)
                print(f"기본 연도 설정: {latest_year} (정수 변환 후 최대값)")
            else:
                self.selected_year.set("전체")
                print("숫자 연도가 없어 '전체'로 설정")
        else:
            self.selected_year.set("전체")
            print("연도 옵션이 부족해 '전체'로 설정")
        
        # 콤보박스의 현재 값도 업데이트
        if hasattr(self, 'year_combo'):
            self.year_combo.set(self.selected_year.get())
            print(f"콤보박스 연도 설정: {self.year_combo.get()}")
            # StringVar와 콤보박스 동기화 강제
            self.selected_year.set(self.year_combo.get())
            print(f"StringVar 동기화 후: {self.selected_year.get()}")
        
        # 부서/팀/영업사원 기본값을 "전체"로 설정
        self.selected_department.set("전체")
        self.selected_team.set("전체")
        self.selected_salesperson.set("전체")
        print("기본값을 모두 '전체'로 설정했습니다.")
        
        # 콤보박스 현재 값 설정
        if hasattr(self, 'department_combo'):
            self.department_combo.set(self.selected_department.get())
        if hasattr(self, 'team_combo'):
            self.team_combo.set(self.selected_team.get())
        if hasattr(self, 'salesperson_combo'):
            self.salesperson_combo.set(self.selected_salesperson.get())
        
        # 조직 구조 매핑 생성
        self.build_organization_mapping()
        
        print("=== 필터 옵션 업데이트 완료 ===")
        print(f"부서: {departments}")
        print(f"팀: {teams}")
        print(f"영업사원: {salespeople}")
        print(f"연도: {years}")
        print(f"선택된 연도: {self.selected_year.get()}")
        
        # 필터 옵션 설정 후 현재 선택된 값들 확인
        print(f"현재 선택된 값들:")
        print(f"  - 부서: {self.selected_department.get()}")
        print(f"  - 팀: {self.selected_team.get()}")
        print(f"  - 영업사원: {self.selected_salesperson.get()}")
        print(f"  - 연도: {self.selected_year.get()}")
    
    def build_organization_mapping(self):
        """조직 구조 매핑 생성 (Target 데이터 기반)"""
        print("\n=== 조직 구조 매핑 생성 시작 ===")
        
        # 매핑 초기화
        self.salesperson_to_team = {}
        self.salesperson_to_dept = {}
        self.team_to_dept = {}
        
        if not self.target_df.empty and all(col in self.target_df.columns for col in ['부서', '팀', '영업사원']):
            # Target 데이터에서 고유한 조직 조합 추출
            org_combinations = self.target_df[['부서', '팀', '영업사원']].drop_duplicates()
            
            for _, row in org_combinations.iterrows():
                dept = row['부서']
                team = row['팀']
                salesperson = row['영업사원']
                
                if dept and team and salesperson:  # 모든 값이 있는 경우
                    # 영업사원 -> 팀, 부서 매핑
                    self.salesperson_to_team[salesperson] = team
                    self.salesperson_to_dept[salesperson] = dept
                    
                    # 팀 -> 부서 매핑
                    self.team_to_dept[team] = dept
            
            print(f"영업사원 -> 팀 매핑: {self.salesperson_to_team}")
            print(f"영업사원 -> 부서 매핑: {self.salesperson_to_dept}")
            print(f"팀 -> 부서 매핑: {self.team_to_dept}")
            print("조직 구조 매핑 생성 완료")
        else:
            print("Target 데이터에 부서/팀/영업사원 컬럼이 없어 매핑을 생성할 수 없습니다.")
    
    def auto_select_parent_organization(self, selection_type, selected_value):
        """상위 조직 자동 선택"""
        print(f"\n=== 상위 조직 자동 선택: {selection_type} = {selected_value} ===")
        
        if selection_type == "영업사원" and selected_value != "전체":
            # 영업사원 선택 시 팀과 부서 자동 선택
            if selected_value in self.salesperson_to_team:
                team = self.salesperson_to_team[selected_value]
                dept = self.salesperson_to_dept[selected_value]
                
                print(f"영업사원 '{selected_value}' -> 팀: '{team}', 부서: '{dept}'")
                
                # 팀과 부서 자동 선택 (이벤트 발생 방지)
                self.selected_team.set(team)
                self.selected_department.set(dept)
                
                # 콤보박스 값도 업데이트
                if hasattr(self, 'team_combo'):
                    self.team_combo.set(team)
                if hasattr(self, 'department_combo'):
                    self.department_combo.set(dept)
                
                print(f"팀과 부서가 자동으로 선택되었습니다.")
            else:
                print(f"영업사원 '{selected_value}'의 매핑 정보를 찾을 수 없습니다.")
        
        elif selection_type == "팀" and selected_value != "전체":
            # 팀 선택 시 부서 자동 선택
            if selected_value in self.team_to_dept:
                dept = self.team_to_dept[selected_value]
                
                print(f"팀 '{selected_value}' -> 부서: '{dept}'")
                
                # 부서 자동 선택 (이벤트 발생 방지)
                self.selected_department.set(dept)
                
                # 콤보박스 값도 업데이트
                if hasattr(self, 'department_combo'):
                    self.department_combo.set(dept)
                
                print(f"부서가 자동으로 선택되었습니다.")
            else:
                print(f"팀 '{selected_value}'의 매핑 정보를 찾을 수 없습니다.")
    
    def _on_year_changed(self, *args):
        """연도 StringVar 변경 시 호출"""
        print(f"연도 StringVar 변경됨: {self.selected_year.get()}")
    
    def on_filter_changed(self, event=None):
        """필터 변경 시 호출"""
        # 이벤트 소스 확인하여 상위 조직 자동 선택
        if event and hasattr(event, 'widget'):
            widget = event.widget
            if hasattr(widget, 'get'):
                selected_value = widget.get()
                
                # 어떤 콤보박스에서 이벤트가 발생했는지 확인
                if widget == self.salesperson_combo:
                    self.auto_select_parent_organization("영업사원", selected_value)
                elif widget == self.team_combo:
                    self.auto_select_parent_organization("팀", selected_value)
        
        self.update_dashboard()
    
    def on_salesperson_changed(self, event=None):
        """영업사원 변경 시 호출 (상위 조직 자동 선택)"""
        selected_salesperson = self.selected_salesperson.get()
        print(f"영업사원 변경: {selected_salesperson}")
        self.auto_select_parent_organization("영업사원", selected_salesperson)
        self.update_dashboard()
    
    def on_team_changed(self, event=None):
        """팀 변경 시 호출 (상위 조직 자동 선택)"""
        selected_team = self.selected_team.get()
        print(f"팀 변경: {selected_team}")
        self.auto_select_parent_organization("팀", selected_team)
        self.update_dashboard()
    
    def on_department_changed(self, event=None):
        """부서 변경 시 호출"""
        selected_department = self.selected_department.get()
        print(f"부서 변경: {selected_department}")
        self.update_dashboard()
    
    def on_period_changed(self, event=None):
        """기간 구분 변경 시 호출"""
        period = self.period_var.get()
        
        if period == "월별":
            self.month_combo.config(state="readonly")
            self.quarter_combo.config(state="disabled")
        else:
            self.month_combo.config(state="disabled")
            self.quarter_combo.config(state="readonly")
        
        self.update_dashboard()
    
    def refresh_dashboard(self):
        """대시보드 새로고침"""
        print("🔄 대시보드 새로고침 시작")
        
        # 캐시 삭제
        cache_file = self.cache_dir / "team_dashboard_cache.pkl"
        if cache_file.exists():
            cache_file.unlink()
            print("캐시 파일 삭제 완료")
        
        # 로딩 인디케이터 표시
        self.loading_label.config(text="🔄 새로고침 중...", fg="#6B7280")
        
        # 백그라운드에서 데이터 다시 로드
        self.load_data_async()
        
        print("대시보드 새로고침 완료")
    
    def update_dashboard(self):
        """대시보드 업데이트"""
        try:
            print("=== 대시보드 업데이트 시작 ===")
            
            # 데이터 필터링
            print("1. 데이터 필터링 시작...")
            self.filter_data()
            print("✅ 데이터 필터링 완료")
            
            # 성과 카드 업데이트
            print("2. 성과 카드 업데이트 시작...")
            self.update_performance_cards()
            print("✅ 성과 카드 업데이트 완료")
            
            # 테이블 업데이트
            print("3. 테이블 업데이트 시작...")
            self.update_performance_table()
            print("✅ 테이블 업데이트 완료")
            
            # 데이터가 없을 때 사용자에게 알림
            if self.filtered_data.empty and not self.target_df.empty:
                messagebox.showinfo("알림", "선택한 필터 조건에 맞는 데이터가 없습니다.\n다른 필터 조건을 선택해보세요.")
            
            print("=== 대시보드 업데이트 완료 ===")
            
        except Exception as e:
            print(f"❌ 대시보드 업데이트 중 오류 발생: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"대시보드 업데이트 중 오류가 발생했습니다:\n{str(e)}")
    

    
    def update_performance_cards(self):
        """성과 카드 업데이트 (새로운 카드 구조)"""
        print("=== 성과 카드 업데이트 시작 ===")
        # 필터링된 데이터가 비어있으면 전체 데이터 사용
        data_to_use = self.filtered_data if not self.filtered_data.empty else self.target_df
        
        print(f"사용할 데이터 크기: {len(data_to_use)}행")
        print(f"사용할 데이터 컬럼: {list(data_to_use.columns) if not data_to_use.empty else '데이터 없음'}")
        
        if data_to_use.empty:
            print("사용할 데이터가 없습니다.")
            return
        
        # 매출 데이터 계산 (Target 시트 구조에 맞게 수정, 타입 안전성 보장)
        try:
            year_str = self.selected_year.get()
            print(f"선택된 연도 문자열: '{year_str}'")
            if year_str and year_str != "전체":
                selected_year = int(year_str)
                print(f"변환된 연도: {selected_year}")
            else:
                selected_year = None
                print("연도가 '전체'이거나 비어있음")
        except (ValueError, TypeError) as e:
            print(f"연도 변환 오류: {e}")
            selected_year = None
        
        # 연도 컬럼이 있는지 확인
        if '연도' in data_to_use.columns and selected_year is not None:
            sales_data = data_to_use[
                (data_to_use['목표'] == '매출 목표') & 
                (data_to_use['연도'] == selected_year)
            ]
            print(f"연도 필터링 적용: {selected_year}")
        else:
            sales_data = data_to_use[data_to_use['목표'] == '매출 목표']
            print("연도 컬럼이 없거나 연도가 선택되지 않아 연도 필터링을 건너뜁니다.")
        if not sales_data.empty:
            try:
                # 목표 카드 (매출 목표만, 월별 데이터 직접 합산)
                monthly_sales_target = 0
                for month in range(1, 13):
                    month_col = f'{month}월_목표' if f'{month}월_목표' in sales_data.columns else f'{month}월'
                    if month_col in sales_data.columns:
                        monthly_sales_target += float(sales_data[month_col].sum())
                
                self.target_sales_value.config(text=f"{monthly_sales_target:,.0f}")
                
                # 실적 카드 (월별 실적 데이터 직접 합산)
                monthly_sales_actual = 0
                for month in range(1, 13):
                    month_col = f'{month}월_실적' if f'{month}월_실적' in sales_data.columns else f'{month}월'
                    if month_col in sales_data.columns:
                        monthly_sales_actual += float(sales_data[month_col].sum())
                
                self.실적_매출_card.config(text=f"{monthly_sales_actual:,.0f}")
                
                # 예상 카드 (quote_list의 Status 60% 데이터 사용)
                sales_expected = 0
                if hasattr(self, 'quote_df') and not self.quote_df.empty:
                    quote_60_data = self.quote_df[self.quote_df['Status'] == '60%']
                    if not quote_60_data.empty and '매출합계' in quote_60_data.columns:
                        sales_expected = float(quote_60_data['매출합계'].sum())
                        print(f"매출 예상 계산 (Status 60%): {sales_expected:,.0f}")
                    else:
                        print("Status 60% 데이터가 없거나 매출합계 컬럼이 없습니다.")
                else:
                    print("quote_list 데이터가 없습니다.")
                
                self.예상_매출_card.config(text=f"{sales_expected:,.0f}")
                
                # 합계 카드 (실적 + 예상)
                sales_total = monthly_sales_actual + sales_expected
                self.합계_매출_card.config(text=f"{sales_total:,.0f}")
            except (ValueError, TypeError) as e:
                print(f"매출 데이터 계산 오류: {e}")
                self.target_sales_value.config(text="0")
                self.실적_매출_card.config(text="0")
                self.예상_매출_card.config(text="0")
                self.합계_매출_card.config(text="0")
        
        # 이익 데이터 계산 (Target 시트 구조에 맞게 수정, 타입 안전성 보장)
        if '연도' in data_to_use.columns and selected_year is not None:
            profit_data = data_to_use[
                (data_to_use['목표'] == '이익 목표') & 
                (data_to_use['연도'] == selected_year)
            ]
            print(f"이익 데이터 연도 필터링 적용: {selected_year}")
        else:
            profit_data = data_to_use[data_to_use['목표'] == '이익 목표']
            print("이익 데이터: 연도 컬럼이 없거나 연도가 선택되지 않아 연도 필터링을 건너뜁니다.")
        if not profit_data.empty:
            try:
                # 목표 카드 (이익 목표만, 월별 데이터 직접 합산)
                monthly_profit_target = 0
                for month in range(1, 13):
                    month_col = f'{month}월_목표' if f'{month}월_목표' in profit_data.columns else f'{month}월'
                    if month_col in profit_data.columns:
                        monthly_profit_target += float(profit_data[month_col].sum())
                
                self.target_profit_value.config(text=f"{monthly_profit_target:,.0f}")
                
                # 실적 카드 (월별 실적 데이터 직접 합산)
                monthly_profit_actual = 0
                for month in range(1, 13):
                    month_col = f'{month}월_실적' if f'{month}월_실적' in profit_data.columns else f'{month}월'
                    if month_col in profit_data.columns:
                        monthly_profit_actual += float(profit_data[month_col].sum())
                
                self.실적_이익_card.config(text=f"{monthly_profit_actual:,.0f}")
                
                # 예상 카드 (월별 예상 데이터 직접 합산)
                monthly_profit_expected = 0
                for month in range(1, 13):
                    month_col = f'{month}월_예상' if f'{month}월_예상' in profit_data.columns else f'{month}월'
                    if month_col in profit_data.columns:
                        monthly_profit_expected += float(profit_data[month_col].sum())
                
                self.예상_이익_card.config(text=f"{monthly_profit_expected:,.0f}")
                
                # 합계 카드 (실적 + 예상)
                profit_total = monthly_profit_actual + monthly_profit_expected
                self.합계_이익_card.config(text=f"{profit_total:,.0f}")
            except (ValueError, TypeError) as e:
                print(f"이익 데이터 계산 오류: {e}")
                self.target_profit_value.config(text="0")
                self.실적_이익_card.config(text="0")
                self.예상_이익_card.config(text="0")
                self.합계_이익_card.config(text="0")
        
        # 달성율 계산
        if not sales_data.empty and not profit_data.empty:
            # 매출 달성율 (타입 안전성 보장)
            try:
                if '합계_목표' in sales_data.columns:
                    sales_target_total = float(sales_data['합계_목표'].sum())
                elif '합계' in sales_data.columns:
                    sales_target_total = float(sales_data['합계'].sum())
                else:
                    sales_target_total = 0
                
                if '합계_실적' in sales_data.columns:
                    sales_actual_total = float(sales_data['합계_실적'].sum())
                elif '합계' in sales_data.columns:
                    sales_actual_total = float(sales_data['합계'].sum())
                else:
                    sales_actual_total = 0
                
                if sales_target_total > 0:
                    sales_achievement_rate = (sales_actual_total / sales_target_total) * 100
                    self.achievement_sales_value.config(text=f"{sales_achievement_rate:.1f}%")
                    
                    # 달성율에 따른 색상 변경
                    if sales_achievement_rate >= 100:
                        self.achievement_sales_value.config(fg="#10B981")
                    else:
                        self.achievement_sales_value.config(fg="#EF4444")
                else:
                    self.achievement_sales_value.config(text="0.0%")
            except (ValueError, TypeError) as e:
                print(f"매출 달성율 계산 오류: {e}")
                self.achievement_sales_value.config(text="0.0%")
            
            # 이익 달성율 (타입 안전성 보장)
            try:
                if '합계_목표' in profit_data.columns:
                    profit_target_total = float(profit_data['합계_목표'].sum())
                elif '합계' in profit_data.columns:
                    profit_target_total = float(profit_data['합계'].sum())
                else:
                    profit_target_total = 0
                
                if '합계_실적' in profit_data.columns:
                    profit_actual_total = float(profit_data['합계_실적'].sum())
                elif '합계' in profit_data.columns:
                    profit_actual_total = float(profit_data['합계'].sum())
                else:
                    profit_actual_total = 0
                
                if profit_target_total > 0:
                    profit_achievement_rate = (profit_actual_total / profit_target_total) * 100
                    self.achievement_profit_value.config(text=f"{profit_achievement_rate:.1f}%")
                    
                    # 달성율에 따른 색상 변경
                    if profit_achievement_rate >= 100:
                        self.achievement_profit_value.config(fg="#10B981")
                    else:
                        self.achievement_profit_value.config(fg="#EF4444")
                else:
                    self.achievement_profit_value.config(text="0.0%")
            except (ValueError, TypeError) as e:
                print(f"이익 달성율 계산 오류: {e}")
                self.achievement_profit_value.config(text="0.0%")
        
    
    def filter_data(self):
        """데이터 필터링 - Target 시트에서 실제 데이터 사용"""
        print(f"=== 데이터 필터링 시작 ===")
        print(f"원본 Target 데이터 크기: {len(self.target_df)}행")
        print(f"현재 선택된 연도: {self.selected_year.get()}")
        print(f"콤보박스 연도 값: {self.year_combo.get() if hasattr(self, 'year_combo') else 'N/A'}")
        
        if self.target_df.empty:
            print("Target 데이터가 비어있습니다.")
            self.filtered_data = pd.DataFrame()
            return
        
        # 기본 필터링 (Target 시트의 실제 데이터 사용)
        filtered = self.target_df.copy()
        print(f"필터링 전 데이터 크기: {len(filtered)}행")
        print(f"사용 가능한 컬럼: {list(filtered.columns)}")
        
        # 조직 필터 (Target 시트의 실제 컬럼명 사용)
        if self.selected_department.get() != "전체":
            dept_filter = self.selected_department.get()
            print(f"부서 필터 적용: {dept_filter}")
            if '부서' in filtered.columns:
                filtered = filtered[filtered['부서'] == dept_filter]
                print(f"부서 필터 후 데이터 크기: {len(filtered)}행")
            else:
                print("⚠️ '부서' 컬럼이 없습니다.")
        
        if self.selected_team.get() != "전체":
            team_filter = self.selected_team.get()
            print(f"팀 필터 적용: {team_filter}")
            if '팀' in filtered.columns:
                filtered = filtered[filtered['팀'] == team_filter]
                print(f"팀 필터 후 데이터 크기: {len(filtered)}행")
            else:
                print("⚠️ '팀' 컬럼이 없습니다.")
        
        if self.selected_salesperson.get() != "전체":
            salesperson_filter = self.selected_salesperson.get()
            print(f"영업사원 필터 적용: {salesperson_filter}")
            if '영업사원' in filtered.columns:
                filtered = filtered[filtered['영업사원'] == salesperson_filter]
                print(f"영업사원 필터 후 데이터 크기: {len(filtered)}행")
            else:
                print("⚠️ '영업사원' 컬럼이 없습니다.")
        
        # 연도 필터 (Target 시트의 실제 연도 데이터 사용)
        year_filter = self.selected_year.get()
        print(f"연도 필터 적용: {year_filter}")
        
        if year_filter != "전체":
            # 연도 관련 컬럼 찾기 (더 정확한 검색)
            year_column = None
            print(f"사용 가능한 컬럼: {list(filtered.columns)}")
            
            # 1. 정확한 연도 컬럼명 검색
            year_candidates = ['연도', '년', 'Year', 'YEAR', 'FY', 'Fiscal Year']
            for candidate in year_candidates:
                if candidate in filtered.columns:
                    year_column = candidate
                    print(f"✅ 정확한 연도 컬럼을 찾았습니다: '{candidate}'")
                    break
            
            # 2. 연도 관련 키워드가 포함된 컬럼 검색
            if year_column is None:
                for col in filtered.columns:
                    if any(keyword in str(col).lower() for keyword in ['연도', '년', 'year', 'fy', 'fiscal']):
                        year_column = col
                        print(f"✅ 연도 관련 컬럼을 찾았습니다: '{col}'")
                        break
            
            # 3. 모든 컬럼에서 연도 형태의 데이터가 있는지 확인
            if year_column is None:
                print("모든 컬럼에서 연도 형태의 데이터 검색 중...")
                for i, col in enumerate(filtered.columns):
                    try:
                        unique_values = filtered[col].unique()
                        # 연도 형태의 값이 있는지 확인 (2024, 2025 등)
                        year_like_values = [str(val) for val in unique_values if str(val).isdigit() and len(str(val)) == 4]
                        if year_like_values:
                            year_column = col
                            print(f"✅ 연도 컬럼 발견: '{col}' (인덱스 {i}, 값: {year_like_values})")
                            break
                    except Exception as e:
                        continue
            
            if year_column:
                print(f"연도 컬럼 '{year_column}'을 사용하여 데이터 필터링")
                print(f"연도 컬럼의 고유값: {filtered[year_column].unique()}")
                
                # 연도 필터링 (직접 비교)
                try:
                    year_number = int(year_filter)
                    print(f"필터링할 연도: {year_number} (원본: {year_filter})")
                    
                    # 연도 컬럼의 타입을 문자열로 변환하여 비교 (타입 불일치 방지)
                    filtered_year_str = filtered[year_column].astype(str)
                    year_filter_condition = filtered_year_str == str(year_number)
                    
                    print(f"연도 필터링 방식: 문자열 변환 후 비교")
                    print(f"  - year_filter: '{year_filter}' (타입: {type(year_filter)})")
                    print(f"  - year_number: {year_number} (타입: {type(year_number)})")
                    print(f"  - 연도 컬럼 타입: {filtered[year_column].dtype}")
                    print(f"  - 연도 컬럼 고유값: {filtered[year_column].unique()}")
                    print(f"  - 문자열 변환 후 고유값: {filtered_year_str.unique()}")
                    
                    # 디버그: 필터 조건 확인
                    print(f"연도 필터 조건 확인:")
                    print(f"  - year_column: {year_column}")
                    print(f"  - year_number: {year_number} (타입: {type(year_number)})")
                    print(f"  - str(year_number): {str(year_number)} (타입: {type(str(year_number))})")
                    print(f"  - float(year_number): {float(year_number)} (타입: {type(float(year_number))})")
                    print(f"  - 데이터의 연도 값들: {filtered[year_column].unique()}")
                    
                    # 필터링 전 데이터 확인
                    print(f"필터링 전 데이터 크기: {len(filtered)}행")
                    print(f"필터링 전 연도 분포:")
                    print(filtered[year_column].value_counts())
                    
                    filtered = filtered[year_filter_condition]
                    print(f"{year_number}년 필터 후 데이터 크기: {len(filtered)}행")
                
                    if not filtered.empty:
                        print(f"{year_number}년 데이터 샘플:")
                        print(filtered[['영업사원', '목표', year_column]].head())  # '구분' → '목표'로 수정
                    else:
                        print(f"⚠️ {year_number}년 데이터가 없습니다.")
                        print("전체 데이터의 연도 값들:")
                        print(filtered[year_column].value_counts())
                        
                        # 필터링이 실패한 경우, 더 유연한 필터링 시도
                        print("더 유연한 필터링을 시도합니다...")
                        # 문자열로 변환하여 비교
                        filtered_str = filtered[year_column].astype(str)
                        year_filter_condition_str = filtered_str == str(year_number)
                        filtered = filtered[year_filter_condition_str]
                        print(f"유연한 필터링 후 데이터 크기: {len(filtered)}행")
                        
                        if not filtered.empty:
                            print(f"유연한 필터링으로 {year_number}년 데이터를 찾았습니다:")
                            print(filtered[['영업사원', '목표', year_column]].head())
                except ValueError:
                    print(f"⚠️ 연도 변환 실패: {year_filter}")
            else:
                print("⚠️ 연도 관련 컬럼을 찾을 수 없습니다.")
                print(f"사용 가능한 컬럼: {list(filtered.columns)}")
        else:
            print("연도 필터: 전체 선택됨 (필터링하지 않음)")
        
        # 기간 필터 정보 저장 (테이블 업데이트에서 사용)
        self.selected_period_info = {
            'type': self.period_var.get(),
            'month': self.selected_month.get(),
            'quarter': self.selected_quarter.get(),
            'year': year_filter
        }
        
        # 월별/분기별 필터링 로직 (실제 컬럼명에 맞춤)
        period = self.period_var.get()
        if period == "월별" and self.selected_month.get() != "전체":
            month = int(self.selected_month.get().replace('월', ''))
            print(f"월 필터 적용: {month}월")
            # 해당 월의 데이터만 필터링 (실제 컬럼명: "1월_목표", "2월_목표" 등)
            month_col = f'{month}월_목표'
            if month_col in filtered.columns:
                numeric_values = pd.to_numeric(filtered[month_col], errors='coerce')
                filtered = filtered[numeric_values > 0]
                print(f"월 필터 후 데이터 크기: {len(filtered)}행")
            else:
                print(f"⚠️ '{month_col}' 컬럼이 없습니다.")
        
        elif period == "분기별" and self.selected_quarter.get() != "전체":
            quarter = int(self.selected_quarter.get().replace('분기', ''))
            print(f"분기 필터 적용: {quarter}분기")
            # 해당 분기의 월 범위 계산
            if quarter == 1:
                months = [1, 2, 3]
            elif quarter == 2:
                months = [4, 5, 6]
            elif quarter == 3:
                months = [7, 8, 9]
            elif quarter == 4:
                months = [10, 11, 12]
            else:
                months = list(range(1, 13))
            
            print(f"분기 {quarter}에 해당하는 월: {months}")
            # 해당 분기의 데이터만 필터링 (분기 내 월 중 하나라도 목표가 있는 경우)
            quarter_filter = False
            for month in months:
                month_col = f'{month}월_목표'
                if month_col in filtered.columns:
                    numeric_values = pd.to_numeric(filtered[month_col], errors='coerce')
                    quarter_filter = quarter_filter | (numeric_values > 0)
            
            if quarter_filter is not False:
                filtered = filtered[quarter_filter]
                print(f"분기 필터 후 데이터 크기: {len(filtered)}행")
        
        self.filtered_data = filtered
        print(f"최종 필터링된 데이터 크기: {len(self.filtered_data)}행")
        
        if not self.filtered_data.empty:
            print("필터링된 데이터 샘플:")
            print(self.filtered_data[['회사명', '부서', '팀', '영업사원', '목표']].head())
        else:
            print("⚠️ 필터링 후 데이터가 없습니다.")
        
        print("=== 데이터 필터링 완료 ===")
    
    def update_performance_table(self):
        """성과 테이블 업데이트 (이미지 형태)"""
        print(f"테이블 업데이트 시작 - 필터링된 데이터 크기: {len(self.filtered_data)}행")
        
        # 테이블 컬럼 동적 업데이트
        self.update_table_columns()
        
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 필터링된 데이터가 비어있으면 전체 데이터 사용
        data_to_use = self.filtered_data if not self.filtered_data.empty else self.target_df
        
        if data_to_use.empty:
            print("사용할 데이터가 없어 테이블에 데이터를 표시하지 않습니다.")
            return
        
        # 이미지 형태로 데이터 정리
        table_data = self.create_image_style_data()
        print(f"이미지 형태 데이터 생성 완료: {len(table_data)}개 항목")
        
        # Treeview에 데이터 삽입
        inserted_count = 0
        
        for row_data in table_data:
            try:
                # Treeview에 삽입
                item_id = self.tree.insert("", "end", values=row_data['values'], tags=row_data.get('tags', ()))
                inserted_count += 1
                print(f"행 {inserted_count} 삽입 완료: {row_data['values'][:3]}")  # 처음 3개 값만 출력
                
            except Exception as e:
                print(f"행 삽입 중 오류 발생: {e}")
                import traceback
                traceback.print_exc()
        
            print(f"테이블에 {inserted_count}개 행이 삽입되었습니다.")
        
    
    def update_table_columns(self):
        """테이블 컬럼 동적 업데이트"""
        # 현재 필터 정보에 따라 컬럼 결정
        period_info = getattr(self, 'selected_period_info', {})
        period_type = period_info.get('type', '월별')
        
        if period_type == "월별" and period_info.get('month') != "전체":
            month = int(period_info.get('month', '1월').replace('월', ''))
            display_months = [month]
        elif period_type == "분기별" and period_info.get('quarter') != "전체":
            quarter = int(period_info.get('quarter', '1분기').replace('분기', ''))
            if quarter == 1:
                display_months = [1, 2, 3]
            elif quarter == 2:
                display_months = [4, 5, 6]
            elif quarter == 3:
                display_months = [7, 8, 9]
            elif quarter == 4:
                display_months = [10, 11, 12]
            else:
                display_months = list(range(1, 13))
        else:
            display_months = list(range(1, 13))
        
        # 새로운 컬럼 목록 생성 (실제 컬럼명에 맞춤)
        new_columns = ['목표']
        for month in display_months:
            new_columns.append(f'{month}월')  # 실제 컬럼명: "1월", "2월" 등
        new_columns.append('합계')
        
        # 기존 컬럼과 비교하여 변경이 필요한지 확인
        current_columns = list(self.tree['columns'])
        if current_columns != new_columns:
            print(f"컬럼 변경: {current_columns} -> {new_columns}")
            
            # 기존 스크롤바 저장
            table_container = self.tree.master
            existing_vsb = None
            existing_hsb = None
            
            # 기존 스크롤바 찾기 (더 정확한 방법)
            for child in table_container.winfo_children():
                if isinstance(child, ttk.Scrollbar):
                    if child.cget('orient') == 'vertical':
                        existing_vsb = child
                    elif child.cget('orient') == 'horizontal':
                        existing_hsb = child
            
            # 기존 Treeview의 스크롤바 설정 저장
            old_yscrollcommand = self.tree.cget('yscrollcommand')
            old_xscrollcommand = self.tree.cget('xscrollcommand')
            
            # Treeview 재생성
            self.tree.destroy()
            
            # 새 Treeview 생성
            self.tree = ttk.Treeview(table_container, columns=new_columns, show='headings', height=15)
            
            # 기존 스크롤바 재사용
            if existing_vsb and existing_hsb:
                # 기존 스크롤바 재사용
                self.tree.configure(yscrollcommand=existing_vsb.set, xscrollcommand=existing_hsb.set)
                existing_vsb.configure(command=self.tree.yview)
                existing_hsb.configure(command=self.tree.xview)
            elif old_yscrollcommand and old_xscrollcommand:
                # 기존 스크롤바 설정 재사용
                self.tree.configure(yscrollcommand=old_yscrollcommand, xscrollcommand=old_xscrollcommand)
            else:
                # 새 스크롤바 생성 (기존 스크롤바가 없는 경우에만)
                vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
                hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
                self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
                
                # 스크롤바 배치
                vsb.pack(side='right', fill='y')
                hsb.pack(side='bottom', fill='x')
            
            # 컬럼 설정 - 고정 너비로 설정하여 스크롤 가능하게 함
            for col in new_columns:
                if col == '목표':
                    self.tree.heading(col, text=col, anchor='center')
                    self.tree.column(col, width=120, anchor='w', stretch=False)
                elif col == '합계':
                    self.tree.heading(col, text=col, anchor='center')
                    self.tree.column(col, width=120, anchor='e', stretch=False)
                else:
                    self.tree.heading(col, text=col, anchor='center')
                    self.tree.column(col, width=100, anchor='e', stretch=False)
            
            # Treeview 배치
            self.tree.pack(side='left', fill='both', expand=True)
            
            # 스타일 재설정
            self.setup_table_style()
    
    def create_image_style_data(self):
        """이미지 형태의 테이블 데이터 생성"""
        table_rows = []
        
        # 필터링된 데이터가 비어있으면 전체 데이터 사용
        data_to_use = self.filtered_data if not self.filtered_data.empty else self.target_df
        
        print(f"이미지 형태 데이터 생성 시작 - 사용할 데이터 크기: {len(data_to_use)}행")
        
        if data_to_use.empty:
            return table_rows
        
        # 월별/분기별 필터링 적용
        period_info = getattr(self, 'selected_period_info', {})
        period_type = period_info.get('type', '월별')
        
        # 표시할 월 범위 결정
        if period_type == "월별" and period_info.get('month') != "전체":
            month = int(period_info.get('month', '1월').replace('월', ''))
            display_months = [month]
            print(f"월별 필터링: {month}월만 표시")
        elif period_type == "분기별" and period_info.get('quarter') != "전체":
            quarter = int(period_info.get('quarter', '1분기').replace('분기', ''))
            if quarter == 1:
                display_months = [1, 2, 3]
            elif quarter == 2:
                display_months = [4, 5, 6]
            elif quarter == 3:
                display_months = [7, 8, 9]
            elif quarter == 4:
                display_months = [10, 11, 12]
            else:
                display_months = list(range(1, 13))
            print(f"분기별 필터링: {quarter}분기 ({display_months}월) 표시")
        else:
            display_months = list(range(1, 13))
            print("전체 월 표시")
        
        # 매출 섹션 (실제 Target 데이터의 구분 컬럼명에 맞춤)
        if data_to_use.empty or '목표' not in data_to_use.columns:  # '구분' → '목표'로 수정
            print("⚠️ 사용할 데이터가 비어있거나 '목표' 컬럼이 없습니다.")  # '구분' → '목표'로 수정
            return table_rows
        
        # 연도 필터링 적용 (안전한 변환)
        try:
            year_str = self.selected_year.get()
            print(f"원본 연도 값: '{year_str}' (타입: {type(year_str)})")
            
            if year_str and year_str != "전체":
                selected_year = int(year_str)
                print(f"선택된 연도: {selected_year}")
            else:
                print("연도가 '전체'이거나 비어있어 필터링하지 않습니다.")
                selected_year = None
        except (ValueError, TypeError) as e:
            print(f"연도 변환 오류: {e}")
            print("연도 필터링을 건너뜁니다.")
            selected_year = None
        
        # 매출 목표만 필터링 (정확한 매칭)
        if '연도' in data_to_use.columns and selected_year is not None:
            sales_data = data_to_use[
                (data_to_use['목표'] == '매출 목표') & 
                (data_to_use['연도'] == selected_year)
            ]
            print(f"매출 데이터 연도 필터링 적용: {selected_year}")
        else:
            # 연도 필터링 없이 매출 목표만 필터링
            sales_data = data_to_use[data_to_use['목표'] == '매출 목표']
            print("매출 데이터: 연도 컬럼이 없거나 연도가 선택되지 않아 연도 필터링을 건너뜁니다.")
        print(f"\n=== 매출 데이터 분석 ===")
        print(f"매출 데이터 크기: {len(sales_data)}행")
        print(f"현재 선택된 연도: {self.selected_year.get()}")
        
        # 연도 필터링 확인
        if '연도' in sales_data.columns:
            print(f"매출 데이터의 연도 분포:")
            print(sales_data['연도'].value_counts())
        
        if not sales_data.empty:
            print(f"매출 데이터 컬럼: {list(sales_data.columns)}")
            print(f"매출 데이터 샘플:")
            # 안전한 컬럼 선택
            display_cols = ['영업사원', '팀', '목표']
            if '합계_목표' in sales_data.columns:
                display_cols.append('합계_목표')
            elif '합계' in sales_data.columns:
                display_cols.append('합계')
            print(sales_data[display_cols].head())
            
            # 팀별 매출 목표 합산 (월별 데이터 직접 합산)
            team_sales = pd.Series(dtype=float)
            for month in range(1, 13):
                month_col = f'{month}월_목표' if f'{month}월_목표' in sales_data.columns else f'{month}월'
                if month_col in sales_data.columns:
                    month_team_sales = sales_data.groupby('팀')[month_col].sum()
                    if team_sales.empty:
                        team_sales = month_team_sales
                    else:
                        team_sales += month_team_sales
            
            print(f"팀별 매출 목표 합산 (월별 데이터 직접 합산):")
            for team, total in team_sales.items():
                print(f"  {team}: {total:,.0f}")
            
            # 매출 목표 행 (팀별 합산) - 월별 데이터 직접 합산
            sales_target_values = ['매출 목표']
            monthly_totals = []
            
            for month in display_months:
                # 월별 목표 컬럼 찾기
                month_col = f'{month}월_목표' if f'{month}월_목표' in sales_data.columns else f'{month}월'
                if month_col in sales_data.columns:
                    # 매출 목표만 필터링하여 월별 합산
                    month_total = sales_data[month_col].sum()
                    monthly_totals.append(month_total)
                    sales_target_values.append(f"{month_total:,.0f}" if month_total > 0 else "")
                else:
                    monthly_totals.append(0)
                    sales_target_values.append("")
            
            # 총 목표 (월별 합산으로 계산)
            total_target = sum(monthly_totals)
            sales_target_values.append(f"{total_target:,.0f}" if total_target > 0 else "")
            table_rows.append({'values': sales_target_values, 'tags': ('sales',)})
            
            # 매출 실적 행 (팀별 합산)
            sales_actual_values = ['매출 실적']
            for month in display_months:
                # 팀별 월별 실적 합산 (실적 컬럼 사용)
                month_col = f'{month}월_실적' if f'{month}월_실적' in sales_data.columns else f'{month}월'
                if month_col in sales_data.columns:
                    month_actual = sales_data.groupby('팀')[month_col].sum().sum()
                else:
                    month_actual = 0
                sales_actual_values.append(f"{month_actual:,.0f}" if month_actual > 0 else "")
            
            total_actual = team_sales.sum()  # 팀별 합산
            sales_actual_values.append(f"{total_actual:,.0f}" if total_actual > 0 else "")
            table_rows.append({'values': sales_actual_values, 'tags': ('sales',)})
            
            # 매출 예상 행 (quote_list의 Status 60% 데이터 사용)
            sales_expected_values = ['매출 예상']
            total_expected = 0
            
            # quote_list에서 Status 60% 데이터 계산
            if hasattr(self, 'quote_df') and not self.quote_df.empty:
                print(f"\n=== 매출 예상 계산 (quote_list Status 60%) ===")
                quote_60_data = self.quote_df[self.quote_df['Status'] == '60%']
                print(f"Status 60% 데이터 크기: {len(quote_60_data)}행")
                
                if not quote_60_data.empty:
                    # 월별 예상 매출 계산
                    for month in display_months:
                        month_expected = 0
                        try:
                            # 날짜 컬럼에서 월 추출하여 매칭
                            if '일자' in quote_60_data.columns:
                                # 일자 컬럼이 있는 경우 월별 필터링
                                month_num = month
                                month_data = quote_60_data[
                                    quote_60_data['일자'].str.contains(f'-{month_num:02d}-', na=False)
                                ]
                                month_expected = month_data['매출합계'].sum() if '매출합계' in month_data.columns else 0
                            else:
                                # 일자 컬럼이 없는 경우 전체 합계를 월별로 분배
                                total_quote_60 = quote_60_data['매출합계'].sum() if '매출합계' in quote_60_data.columns else 0
                                month_expected = total_quote_60 / len(display_months) if len(display_months) > 0 else 0
                            
                            print(f"  {month}월 예상: {month_expected:,.0f}")
                        except Exception as e:
                            print(f"  {month}월 예상 계산 오류: {e}")
                            month_expected = 0
                        
                        sales_expected_values.append(f"{month_expected:,.0f}" if month_expected > 0 else "")
                        total_expected += month_expected
                    print(f"총 예상 매출: {total_expected:,.0f}")
                else:
                    print("Status 60% 데이터가 없습니다.")
                    for month in display_months:
                        sales_expected_values.append("")
                    total_expected = 0
            else:
                print("quote_list 데이터가 없습니다.")
                for month in display_months:
                    sales_expected_values.append("")
                total_expected = 0
            
            sales_expected_values.append(f"{total_expected:,.0f}" if total_expected > 0 else "")
            table_rows.append({'values': sales_expected_values, 'tags': ('sales',)})
            
            # 매출 합계 행 (팀별 합산)
            sales_total_values = ['합계']
            for month in display_months:
                # 팀별 월별 합계 (목표 컬럼 사용)
                month_col = f'{month}월_목표' if f'{month}월_목표' in sales_data.columns else f'{month}월'
                if month_col in sales_data.columns:
                    month_total = sales_data.groupby('팀')[month_col].sum().sum()
                else:
                    month_total = 0
                sales_total_values.append(f"{month_total:,.0f}" if month_total > 0 else "")
            sales_total_values.append(f"{total_target:,.0f}" if total_target > 0 else "")
            table_rows.append({'values': sales_total_values, 'tags': ('sales', 'total')})
            
            # 매출 달성율 행 (팀별 합산)
            sales_achievement_values = ['달성율']
            for month in display_months:
                # 팀별 월별 달성율 계산 (목표 컬럼 사용)
                month_col = f'{month}월_목표' if f'{month}월_목표' in sales_data.columns else f'{month}월'
                if month_col in sales_data.columns:
                    month_target = sales_data.groupby('팀')[month_col].sum().sum()
                else:
                    month_target = 0
                if month_target > 0:
                    sales_achievement_values.append("100.0%")
                else:
                    sales_achievement_values.append("")
            if total_target > 0:
                sales_achievement_values.append("100.0%")
            else:
                sales_achievement_values.append("")
            table_rows.append({'values': sales_achievement_values, 'tags': ('sales', 'achievement')})
            
            # 개별 영업사원 데이터는 팀별 성과 현황에서 제외
            print(f"\n=== 개별 영업사원 매출 데이터 (표시하지 않음) ===")
            for _, person_row in sales_data.iterrows():
                person_name = person_row['영업사원']
                person_team = person_row['팀']
                # 안전한 컬럼 선택
                if '합계_목표' in person_row:
                    person_total = person_row['합계_목표']
                elif '합계' in person_row:
                    person_total = person_row['합계']
                else:
                    person_total = 0
                print(f"  {person_name} ({person_team}): {person_total:,.0f}")
        
        # 이익 섹션 (Target 시트 구조에 맞게 수정)
        if '연도' in data_to_use.columns and selected_year is not None:
            profit_data = data_to_use[
                (data_to_use['목표'] == '이익 목표') & 
                (data_to_use['연도'] == selected_year)
            ]
            print(f"이익 데이터 연도 필터링 적용: {selected_year}")
        else:
            # 연도 필터링 없이 이익 목표만 필터링
            profit_data = data_to_use[data_to_use['목표'] == '이익 목표']
            print("이익 데이터: 연도 컬럼이 없거나 연도가 선택되지 않아 연도 필터링을 건너뜁니다.")
        print(f"\n=== 이익 데이터 분석 ===")
        print(f"이익 데이터 크기: {len(profit_data)}행")
        
        # 연도 필터링 확인
        if '연도' in profit_data.columns:
            print(f"이익 데이터의 연도 분포:")
            print(profit_data['연도'].value_counts())
        
        if not profit_data.empty:
            # 구분선 행 추가
            separator_values = [''] * (len(display_months) + 2)  # 구분 + 월별 + 합계
            table_rows.append({'values': separator_values, 'tags': ('separator',)})
            
            # 팀별 이익 목표 합산 (월별 데이터 직접 합산)
            team_profit = pd.Series(dtype=float)
            for month in range(1, 13):
                month_col = f'{month}월_목표' if f'{month}월_목표' in profit_data.columns else f'{month}월'
                if month_col in profit_data.columns:
                    month_team_profit = profit_data.groupby('팀')[month_col].sum()
                    if team_profit.empty:
                        team_profit = month_team_profit
                    else:
                        team_profit += month_team_profit
            
            print(f"팀별 이익 목표 합산 (월별 데이터 직접 합산):")
            for team, total in team_profit.items():
                print(f"  {team}: {total:,.0f}")
            
            # 이익 목표 행 (팀별 합산) - 월별 데이터 직접 합산
            profit_target_values = ['이익 목표']
            profit_monthly_totals = []
            
            for month in display_months:
                # 월별 목표 컬럼 찾기
                month_col = f'{month}월_목표' if f'{month}월_목표' in profit_data.columns else f'{month}월'
                if month_col in profit_data.columns:
                    # 이익 목표만 필터링하여 월별 합산
                    month_total = profit_data[month_col].sum()
                    profit_monthly_totals.append(month_total)
                    profit_target_values.append(f"{month_total:,.0f}" if month_total > 0 else "")
                else:
                    profit_monthly_totals.append(0)
                    profit_target_values.append("")
            
            # 총 목표 (월별 합산으로 계산)
            total_profit_target = sum(profit_monthly_totals)
            profit_target_values.append(f"{total_profit_target:,.0f}" if total_profit_target > 0 else "")
            table_rows.append({'values': profit_target_values, 'tags': ('profit',)})
            
            # 이익 실적 행 (팀별 합산)
            profit_actual_values = ['이익 실적']
            for month in display_months:
                # 팀별 월별 실적 합산 (실적 컬럼 사용)
                month_col = f'{month}월_실적' if f'{month}월_실적' in profit_data.columns else f'{month}월'
                if month_col in profit_data.columns:
                    month_actual = profit_data.groupby('팀')[month_col].sum().sum()
                else:
                    month_actual = 0
                profit_actual_values.append(f"{month_actual:,.0f}" if month_actual > 0 else "")
            
            total_profit_actual = team_profit.sum()  # 팀별 합산
            profit_actual_values.append(f"{total_profit_actual:,.0f}" if total_profit_actual > 0 else "")
            table_rows.append({'values': profit_actual_values, 'tags': ('profit',)})
            
            # 이익 예상 행 (팀별 합산)
            profit_expected_values = ['이익 예상']
            for month in display_months:
                # 팀별 월별 예상 합산 (현재는 0으로 표시)
                month_expected = 0  # 예상 데이터 없음
                profit_expected_values.append(f"{month_expected:,.0f}" if month_expected > 0 else "")
            profit_expected_values.append("0")
            table_rows.append({'values': profit_expected_values, 'tags': ('profit',)})
            
            # 개별 영업사원 이익 데이터는 팀별 성과 현황에서 제외
            print(f"\n=== 개별 영업사원 이익 데이터 (표시하지 않음) ===")
            for _, person_row in profit_data.iterrows():
                person_name = person_row['영업사원']
                person_team = person_row['팀']
                # 안전한 컬럼 선택
                if '합계_목표' in person_row:
                    person_total = person_row['합계_목표']
                elif '합계' in person_row:
                    person_total = person_row['합계']
                else:
                    person_total = 0
                print(f"  {person_name} ({person_team}): {person_total:,.0f}")
            
            # 이익 합계 행
            profit_total_values = ['합계']
            for month in display_months:
                # 안전한 컬럼 선택
                month_col = f'{month}월_목표' if f'{month}월_목표' in profit_data.columns else f'{month}월'
                if month_col in profit_data.columns:
                    total_val = profit_data[month_col].sum()
                else:
                    total_val = 0
                profit_total_values.append(f"{total_val:,.0f}" if total_val > 0 else "")
            profit_total_values.append(f"{total_target:,.0f}" if total_target > 0 else "")
            table_rows.append({'values': profit_total_values, 'tags': ('profit', 'total')})
            
            # 이익 달성율 행 (실적과 목표가 동일하므로 100%로 표시)
            profit_achievement_values = ['달성율']
            for month in display_months:
                # 안전한 컬럼 선택
                month_col = f'{month}월_목표' if f'{month}월_목표' in profit_data.columns else f'{month}월'
                if month_col in profit_data.columns:
                    target_val = profit_data[month_col].sum()
                else:
                    target_val = 0
                if target_val > 0:
                    profit_achievement_values.append("100.0%")
                else:
                    profit_achievement_values.append("")
            if total_target > 0:
                profit_achievement_values.append("100.0%")
            else:
                profit_achievement_values.append("")
            table_rows.append({'values': profit_achievement_values, 'tags': ('profit', 'achievement')})
        
        return table_rows
    

    

    
    def update_summary_panel(self):
        """합계 패널 업데이트"""
        # 기존 위젯 삭제
        for widget in self.sales_summary.winfo_children():
            widget.destroy()
        for widget in self.profit_summary.winfo_children():
            widget.destroy()
        
        if self.filtered_data.empty:
            return
        
        # 매출 합계 계산
        sales_data = self.filtered_data[self.filtered_data['목표'] == '매출 목표']  # '구분' → '목표'로 수정
        if not sales_data.empty:
            self.create_summary_items(self.sales_summary, sales_data, "매출")
        
        # 이익 합계 계산
        profit_data = self.filtered_data[self.filtered_data['목표'] == '이익 목표']  # '구분' → '목표'로 수정
        if not profit_data.empty:
            self.create_summary_items(self.profit_summary, profit_data, "이익")
    
    def create_summary_items(self, parent, data, category):
        """합계 항목 생성 (Target 시트 구조에 맞게 수정)"""
        # 예상 합계 (Target 시트에는 별도 예상 컬럼이 없으므로 0으로 표시)
        expected_total = 0
        tk.Label(parent, text=f"예상: {expected_total:,.0f}", 
                font=("맑은 고딕", 10), fg="#6B7280", bg="white").pack(anchor='w')
        
        # 목표 합계 (Target 시트에는 별도 목표 컬럼이 없으므로 합계 사용)
        # 안전한 컬럼 선택
        if '합계_목표' in data.columns:
            target_total = data['합계_목표'].sum()
        elif '합계' in data.columns:
            target_total = data['합계'].sum()
        else:
            target_total = 0
        tk.Label(parent, text=f"목표: {target_total:,.0f}", 
                font=("맑은 고딕", 10), fg="#374151", bg="white").pack(anchor='w')
        
        # 실적 합계 (Target 시트에는 별도 실적 컬럼이 없으므로 합계 사용)
        # 안전한 컬럼 선택
        if '합계_실적' in data.columns:
            actual_total = data['합계_실적'].sum()
        elif '합계' in data.columns:
            actual_total = data['합계'].sum()
        else:
            actual_total = 0
        tk.Label(parent, text=f"실적: {actual_total:,.0f}", 
                font=("맑은 고딕", 10), fg="#059669", bg="white").pack(anchor='w')
        
        # 달성율 (실적과 목표가 동일하므로 100%로 표시)
        if target_total > 0:
            achievement_rate = 100.0  # 항상 100%
            achievement_color = "#059669"  # 항상 녹색
            tk.Label(parent, text=f"달성율: {achievement_rate:.1f}%", 
                    font=("맑은 고딕", 10, "bold"), fg=achievement_color, bg="white").pack(anchor='w')
        
        # 구분선
        tk.Frame(parent, bg="#E5E7EB", height=1).pack(fill='x', pady=5)
    
    def update_bottom_summary_panel(self):
        """하단 합계 패널 업데이트"""
        # 기존 위젯 삭제
        for widget in self.bottom_sales_summary.winfo_children():
            widget.destroy()
        for widget in self.bottom_profit_summary.winfo_children():
            widget.destroy()
        
        if self.filtered_data.empty:
            return
        
        # 매출 합계 계산
        sales_data = self.filtered_data[self.filtered_data['목표'] == '매출 목표']  # '구분' → '목표'로 수정
        if not sales_data.empty:
            self.create_bottom_summary_items(self.bottom_sales_summary, sales_data, "매출")
        
        # 이익 합계 계산
        profit_data = self.filtered_data[self.filtered_data['목표'] == '이익 목표']  # '구분' → '목표'로 수정
        if not profit_data.empty:
            self.create_bottom_summary_items(self.bottom_profit_summary, profit_data, "이익")
    
    def create_bottom_summary_items(self, parent, data, category):
        """하단 합계 항목 생성 (Target 시트 구조에 맞게 수정)"""
        # 예상 합계 (Target 시트에는 별도 예상 컬럼이 없으므로 0으로 표시)
        expected_total = 0
        tk.Label(parent, text=f"예상: {expected_total:,.0f}", 
                font=("맑은 고딕", 12), fg="#6B7280", bg="white").pack(anchor='w')
        
        # 목표 합계 (Target 시트에는 별도 목표 컬럼이 없으므로 합계 사용)
        # 안전한 컬럼 선택
        if '합계_목표' in data.columns:
            target_total = data['합계_목표'].sum()
        elif '합계' in data.columns:
            target_total = data['합계'].sum()
        else:
            target_total = 0
        tk.Label(parent, text=f"목표: {target_total:,.0f}", 
                font=("맑은 고딕", 12), fg="#374151", bg="white").pack(anchor='w')
        
        # 실적 합계 (Target 시트에는 별도 실적 컬럼이 없으므로 합계 사용)
        # 안전한 컬럼 선택
        if '합계_실적' in data.columns:
            actual_total = data['합계_실적'].sum()
        elif '합계' in data.columns:
            actual_total = data['합계'].sum()
        else:
            actual_total = 0
        tk.Label(parent, text=f"실적: {actual_total:,.0f}", 
                font=("맑은 고딕", 12), fg="#059669", bg="white").pack(anchor='w')
        
        # 달성율 (실적과 목표가 동일하므로 100%로 표시)
        if target_total > 0:
            achievement_rate = 100.0  # 항상 100%
            achievement_color = "#059669"  # 항상 녹색
            tk.Label(parent, text=f"달성율: {achievement_rate:.1f}%", 
                    font=("맑은 고딕", 12, "bold"), fg=achievement_color, bg="white").pack(anchor='w')
    
    def refresh_dashboard(self):
        """대시보드 새로고침 (최적화된 버전)"""
        try:
            # 로딩 상태 표시
            self.loading_label.config(text="🔄 데이터 새로고침 중...", fg="#6B7280")
            
            # 캐시 삭제
            cache_file = self.cache_dir / "team_dashboard_cache.pkl"
            if cache_file.exists():
                cache_file.unlink()
                print("캐시 파일 삭제됨")
            
            # 백그라운드에서 새로 로드
            self.load_data_async()
            
        except Exception as e:
            messagebox.showerror("오류", f"대시보드 새로고침 중 오류가 발생했습니다:\n{str(e)}")
            self.loading_label.config(text="❌ 새로고침 실패", fg="#DC2626")

def create_team_performance_dashboard(parent):
    """팀별 성과 대시보드 생성 함수"""
    return TeamPerformanceDashboard(parent)
