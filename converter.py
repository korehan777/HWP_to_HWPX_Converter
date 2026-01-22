"""
HWP to HWPX Converter
한글 COM 인터페이스를 사용하여 HWP 파일을 HWPX로 변환
"""
import os
import win32com.client as win32
from typing import Tuple, Optional


class HwpConverter:
    """HWP 파일을 HWPX로 변환하는 클래스"""
    
    def __init__(self):
        self.hwp = None
        self.is_initialized = False
        
    def initialize(self) -> Tuple[bool, str]:
        """한글 프로그램 초기화"""
        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.XHwpWindows.Item(0).Visible = False  # 창 숨기기
            self.is_initialized = True
            return True, "한글 프로그램이 초기화되었습니다."
        except Exception as e:
            return False, f"한글 프로그램 초기화 실패: {str(e)}\n한글이 설치되어 있는지 확인하세요."
    
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
            return False, "한글 프로그램이 초기화되지 않았습니다."
        
        try:
            # HWP 파일 열기
            if not self.hwp.Open(input_path, "HWP", "forceopen:true"):
                return False, f"파일 열기 실패: {os.path.basename(input_path)}"
            
            # HWPX로 저장
            save_format = "HWPX"
            if not self.hwp.SaveAs(output_path, save_format):
                self.hwp.Clear()
                return False, f"파일 저장 실패: {os.path.basename(output_path)}"
            
            # 문서 닫기
            self.hwp.Clear()
            
            return True, f"변환 성공: {os.path.basename(input_path)}"
            
        except Exception as e:
            try:
                self.hwp.Clear()
            except:
                pass
            return False, f"변환 중 오류: {os.path.basename(input_path)} - {str(e)}"
    
    def close(self):
        """한글 프로그램 종료"""
        if self.hwp:
            try:
                self.hwp.Quit()
            except:
                pass
            self.hwp = None
            self.is_initialized = False
