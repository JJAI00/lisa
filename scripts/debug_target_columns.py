#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Target 시트 컬럼 구조 디버깅 스크립트
"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

def debug_target_sheet():
    print("=== Target 시트 컬럼 구조 디버깅 시작 ===")
    
    try:
        # Google Sheets 인증
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json', scopes=scope)
        gc = gspread.authorize(creds)
        
        # Target 시트 접근
        print("1. Target 시트 접근 시도...")
        sh = gc.open_by_key('1C5p4HBn9G8MpPdQf6Q8FqaM4lgUzzfftBqG-qQhLa7Y')
        print("✅ Target 시트 접근 성공")
        
        # 워크시트 목록 확인
        print("2. 워크시트 목록 확인...")
        worksheets = sh.worksheets()
        print(f"발견된 워크시트: {[ws.title for ws in worksheets]}")
        
        # Target 워크시트 접근
        print("3. Target 워크시트 접근 시도...")
        ws = sh.worksheet('Target')
        print("✅ Target 워크시트 'Target' 접근 성공")
        
        # 데이터 로드
        print("4. Target 데이터 로드 시도...")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        print(f"✅ Target 데이터 로드 성공: {len(df)}행")
        
        # 컬럼 구조 분석
        print("5. Target 데이터 컬럼 확인...")
        print(f"원본 컬럼: {list(df.columns)}")
        print(f"컬럼 수: {len(df.columns)}")
        
        # 데이터 샘플 확인
        print("6. 데이터 샘플 확인...")
        if not df.empty:
            print("첫 번째 행:")
            print(df.iloc[0].to_dict())
            
            print("\n두 번째 행:")
            print(df.iloc[1].to_dict())
        
        # '목표' 컬럼 존재 여부 확인
        print("\n7. '목표' 컬럼 존재 여부 확인...")
        if '목표' in df.columns:
            print("✅ '목표' 컬럼 발견")
            print(f"'목표' 컬럼 고유값: {df['목표'].unique()}")
        else:
            print("❌ '목표' 컬럼을 찾을 수 없음")
            print("사용 가능한 컬럼:")
            for i, col in enumerate(df.columns):
                print(f"  {i}: '{col}'")
        
        # 데이터 타입 확인
        print("\n8. 데이터 타입 확인...")
        print(df.dtypes)
        
        return df
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    df = debug_target_sheet()
    if df is not None:
        print(f"\n=== 디버깅 완료 ===")
        print(f"데이터 크기: {df.shape}")
        print(f"컬럼: {list(df.columns)}")

