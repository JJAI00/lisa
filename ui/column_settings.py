import json
import os
from typing import Dict, List, Optional

class ColumnSettings:
    """컬럼 크기 설정을 관리하는 클래스"""
    
    def __init__(self, settings_file="column_settings.json"):
        self.settings_file = settings_file
        self.settings = self._load_settings()
    
    def _load_settings(self) -> Dict:
        """설정 파일에서 데이터 로드"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"설정 파일 로드 오류: {e}")
        return {}
    
    def _save_settings(self):
        """설정을 파일에 저장"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"설정 파일 저장 오류: {e}")
    
    def get_column_widths(self, table_name: str) -> Dict[str, int]:
        """특정 테이블의 컬럼 너비 설정 가져오기"""
        return self.settings.get(table_name, {}).get('column_widths', {})
    
    def set_column_widths(self, table_name: str, column_widths: Dict[str, int]):
        """특정 테이블의 컬럼 너비 설정 저장"""
        if table_name not in self.settings:
            self.settings[table_name] = {}
        self.settings[table_name]['column_widths'] = column_widths
        self._save_settings()
    
    def get_column_order(self, table_name: str) -> List[str]:
        """특정 테이블의 컬럼 순서 가져오기"""
        return self.settings.get(table_name, {}).get('column_order', [])
    
    def set_column_order(self, table_name: str, column_order: List[str]):
        """특정 테이블의 컬럼 순서 저장"""
        if table_name not in self.settings:
            self.settings[table_name] = {}
        self.settings[table_name]['column_order'] = column_order
        self._save_settings()
    
    def clear_table_settings(self, table_name: str):
        """특정 테이블의 설정 삭제"""
        if table_name in self.settings:
            del self.settings[table_name]
            self._save_settings()
    
    def clear_all_settings(self):
        """모든 설정 삭제"""
        self.settings = {}
        self._save_settings()

# 전역 인스턴스
column_settings = ColumnSettings()
