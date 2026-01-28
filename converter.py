"""
HWP to HWPX Converter
한글 2024 HwpConverter.exe를 사용하여 HWP 파일을 HWPX로 변환
"""
import os
import subprocess
from typing import Tuple, Optional


class HwpConverter:
    """HWP 파일을 HWPX로 변환하는 클래스"""
    
    def __init__(self):
        self.converter_path = r"C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\HwpConverter.exe"
        self.is_initialized = False
        
    def initialize(self) -> Tuple[bool, str]:
        """한글 변환기 확인"""
        try:
            if not os.path.exists(self.converter_path):
                return False, "HwpConverter.exe를 찾을 수 없습니다.\n한글 2024가 설치되어 있는지 확인하세요."
            self.is_initialized = True
            return True, "한글 변환기가 준비되었습니다."
        except Exception as e:
            return False, f"변환기 초기화 실패: {str(e)}"
    
    def convert_file(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        HWP 파일을 HWPX로 변환
        
        Args:
            input_path: 입력 HWP 파일 경로
            output_path: 출력 HWPX 파일 경로
            
        Returns:
            (성공여부, 메시지)
        """
        if not self.is_initialized:
            return False, "한글 변환기가 초기화되지 않았습니다."
        
        try:
            # 출력 디렉토리 확인
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # HwpConverter.exe 호출 형식: HwpConverter.exe <input> <output_format> <output_path>
            # 또는: HwpConverter.exe -i <input> -o <output> -f HWPX
            cmd = [
                self.converter_path,
                input_path,
                "HWPX",
                output_path
            ]
            
            # 프로세스 실행
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            # 결과 확인 - 출력 파일이 생성되었는지 확인
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                return True, f"변환 성공: {os.path.basename(input_path)}"
            else:
                # 명령 실패 시 다른 형식으로 시도
                return self._convert_with_alternate_method(input_path, output_path)
            
        except subprocess.TimeoutExpired:
            return False, f"변환 시간 초과: {os.path.basename(input_path)}"
        except Exception as e:
            return False, f"변환 중 오류: {os.path.basename(input_path)} - {str(e)}"
    
    def _convert_with_alternate_method(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """대체 변환 방법 (다른 커맨드라인 형식 시도)"""
        try:
            # -i, -o, -f 옵션 형식으로 시도
            cmd = [
                self.converter_path,
                "-i", input_path,
                "-o", output_path,
                "-f", "HWPX"
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                return True, f"변환 성공: {os.path.basename(input_path)}"
            else:
                return False, f"변환 실패: {os.path.basename(input_path)}"
                
        except Exception as e:
            return False, f"변환 실패: {os.path.basename(input_path)} - {str(e)}"
    
    def close(self):
        """변환기 정리 (필요시)"""
        self.is_initialized = False
