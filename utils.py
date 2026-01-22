"""
유틸리티 함수들
"""
import os
from typing import List


def find_hwp_files(input_folder: str, include_subdirs: bool = False) -> List[str]:
    """
    지정된 폴더에서 HWP 파일 찾기
    
    Args:
        input_folder: 검색할 폴더 경로
        include_subdirs: 하위 폴더 포함 여부
        
    Returns:
        HWP 파일 경로 리스트
    """
    hwp_files = []
    
    if include_subdirs:
        # 하위 폴더 포함
        for root, dirs, files in os.walk(input_folder):
            for file in files:
                if file.lower().endswith('.hwp'):
                    hwp_files.append(os.path.join(root, file))
    else:
        # 현재 폴더만
        try:
            files = os.listdir(input_folder)
            for file in files:
                if file.lower().endswith('.hwp'):
                    full_path = os.path.join(input_folder, file)
                    if os.path.isfile(full_path):
                        hwp_files.append(full_path)
        except Exception as e:
            print(f"폴더 읽기 오류: {e}")
    
    return hwp_files


def get_output_path(input_path: str, input_folder: str, output_folder: str) -> str:
    """
    입력 파일 경로를 기반으로 출력 파일 경로 생성
    
    Args:
        input_path: 입력 HWP 파일 전체 경로
        input_folder: 입력 폴더 경로
        output_folder: 출력 폴더 경로
        
    Returns:
        출력 HWPX 파일 경로
    """
    # 입력 폴더 기준 상대 경로 계산
    rel_path = os.path.relpath(input_path, input_folder)
    
    # 확장자를 .hwpx로 변경
    base_name = os.path.splitext(rel_path)[0]
    output_rel_path = base_name + '.hwpx'
    
    # 출력 폴더와 결합
    output_path = os.path.join(output_folder, output_rel_path)
    
    # 출력 디렉토리가 없으면 생성
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    return output_path


def validate_folders(input_folder: str, output_folder: str) -> tuple:
    """
    입력/출력 폴더 유효성 검사
    
    Returns:
        (유효여부, 메시지)
    """
    if not input_folder or not os.path.exists(input_folder):
        return False, "입력 폴더를 선택해주세요."
    
    if not output_folder:
        return False, "출력 폴더를 선택해주세요."
    
    if not os.path.isdir(input_folder):
        return False, "입력 폴더 경로가 올바르지 않습니다."
    
    return True, "OK"
