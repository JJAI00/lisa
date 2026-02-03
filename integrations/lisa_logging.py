# lisa_logging.py
"""
LISA 애플리케이션 공용 로깅 설정 모듈
모든 모듈에서 이 모듈을 import하여 일관된 로깅을 사용합니다.
"""

import logging
import os
from datetime import datetime

# 로그 파일 디렉토리 설정
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 로그 파일명 (날짜별)
LOG_FILE = os.path.join(LOG_DIR, f'lisa_{datetime.now().strftime("%Y%m%d")}.log')

# 로거 설정
def setup_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """
    모듈별 로거를 설정하고 반환합니다.
    
    Args:
        name: 로거 이름 (보통 모듈 이름 사용)
        level: 로깅 레벨 (기본값: INFO)
    
    Returns:
        설정된 Logger 객체
    """
    logger = logging.getLogger(name)
    
    # 이미 핸들러가 설정되어 있으면 반환
    if logger.handlers:
        return logger
    
    logger.setLevel(level)
    
    # 포맷터 설정
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 콘솔 핸들러 (DEBUG 레벨 이상 모두 표시)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)
    
    # 파일 핸들러 (INFO 레벨 이상만 파일에 기록)
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    
    return logger


# 전역 로거 (빠른 접근용)
_main_logger = None

def get_logger() -> logging.Logger:
    """메인 LISA 로거를 반환합니다."""
    global _main_logger
    if _main_logger is None:
        _main_logger = setup_logger('LISA')
    return _main_logger


# 개발 모드 플래그 (True면 DEBUG 레벨 로그도 파일에 기록)
DEBUG_MODE = False

def set_debug_mode(enabled: bool):
    """디버그 모드를 설정합니다."""
    global DEBUG_MODE
    DEBUG_MODE = enabled
    logger = get_logger()
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            handler.setLevel(logging.DEBUG if enabled else logging.INFO)
    logger.info(f"디버그 모드 {'활성화' if enabled else '비활성화'}")
