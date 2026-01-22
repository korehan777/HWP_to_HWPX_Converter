"""
HWP to HWPX 변환 프로그램 - GUI 메인
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import threading
from datetime import datetime

# 같은 폴더의 모듈 import
from converter import HwpConverter
from utils import find_hwp_files, get_output_path, validate_folders


class HwpConverterGUI:
    """HWP to HWPX 변환기 GUI 클래스"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("HWP → HWPX 일괄 변환기")
        self.root.geometry("700x650")
        self.root.resizable(False, False)
        
        # 변수 초기화
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.include_subdirs = tk.BooleanVar(value=False)
        self.is_converting = False
        
        # GUI 구성
        self.create_widgets()
        
    def create_widgets(self):
        """GUI 위젯 생성"""
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="HWP → HWPX 일괄 변환기", 
                                font=("맑은 고딕", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 입력 폴더 선택
        ttk.Label(main_frame, text="입력 폴더:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_folder, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_input).grid(
            row=1, column=2, pady=5)
        
        # 출력 폴더 선택
        ttk.Label(main_frame, text="출력 폴더:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_output).grid(
            row=2, column=2, pady=5)
        
        # 옵션
        option_frame = ttk.LabelFrame(main_frame, text="옵션", padding="10")
        option_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Checkbutton(option_frame, text="하위 폴더 포함", 
                       variable=self.include_subdirs).pack(anchor=tk.W)
        
        # 변환 버튼
        self.convert_btn = ttk.Button(main_frame, text="변환 시작", 
                                     command=self.start_conversion,
                                     style="Accent.TButton")
        self.convert_btn.grid(row=4, column=0, columnspan=3, pady=10)
        
        # 진행률 바
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=660)
        self.progress.grid(row=5, column=0, columnspan=3, pady=5)
        
        # 상태 레이블
        self.status_label = ttk.Label(main_frame, text="대기 중...", 
                                     foreground="gray")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="변환 로그", padding="5")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80,
                                                  font=("맑은 고딕", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 버전 정보
        version_label = ttk.Label(main_frame, text="v1.0 | 2026", 
                                 foreground="gray", font=("맑은 고딕", 8))
        version_label.grid(row=8, column=0, columnspan=3, pady=(10, 0))
        
    def browse_input(self):
        """입력 폴더 선택"""
        folder = filedialog.askdirectory(title="입력 폴더 선택")
        if folder:
            self.input_folder.set(folder)
            self.log(f"입력 폴더 선택: {folder}")
    
    def browse_output(self):
        """출력 폴더 선택"""
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_folder.set(folder)
            self.log(f"출력 폴더 선택: {folder}")
    
    def log(self, message: str, level: str = "INFO"):
        """로그 메시지 출력"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message: str, color: str = "black"):
        """상태 레이블 업데이트"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()
    
    def start_conversion(self):
        """변환 시작"""
        if self.is_converting:
            messagebox.showwarning("경고", "이미 변환 작업이 진행 중입니다.")
            return
        
        # 폴더 유효성 검사
        is_valid, message = validate_folders(
            self.input_folder.get(), 
            self.output_folder.get()
        )
        if not is_valid:
            messagebox.showerror("오류", message)
            return
        
        # 출력 폴더 생성
        try:
            os.makedirs(self.output_folder.get(), exist_ok=True)
        except Exception as e:
            messagebox.showerror("오류", f"출력 폴더 생성 실패: {e}")
            return
        
        # 변환 스레드 시작
        self.is_converting = True
        self.convert_btn.config(state=tk.DISABLED)
        thread = threading.Thread(target=self.convert_files, daemon=True)
        thread.start()
    
    def convert_files(self):
        """파일 변환 실행"""
        try:
            # HWP 파일 찾기
            self.update_status("HWP 파일 검색 중...", "blue")
            hwp_files = find_hwp_files(
                self.input_folder.get(), 
                self.include_subdirs.get()
            )
            
            if not hwp_files:
                self.log("변환할 HWP 파일이 없습니다.", "WARNING")
                self.update_status("변환할 파일이 없습니다.", "orange")
                messagebox.showinfo("알림", "선택한 폴더에 HWP 파일이 없습니다.")
                return
            
            self.log(f"총 {len(hwp_files)}개의 HWP 파일을 찾았습니다.")
            
            # 변환기 초기화
            self.update_status("한글 프로그램 초기화 중...", "blue")
            converter = HwpConverter()
            success, message = converter.initialize()
            
            if not success:
                self.log(f"오류: {message}", "ERROR")
                self.update_status("초기화 실패", "red")
                messagebox.showerror("오류", message)
                return
            
            self.log(message)
            
            # 진행률 설정
            total = len(hwp_files)
            self.progress['maximum'] = total
            self.progress['value'] = 0
            
            success_count = 0
            fail_count = 0
            
            # 파일 변환
            for idx, hwp_file in enumerate(hwp_files, 1):
                self.update_status(f"변환 중... ({idx}/{total})", "blue")
                
                output_path = get_output_path(
                    hwp_file, 
                    self.input_folder.get(), 
                    self.output_folder.get()
                )
                
                success, msg = converter.convert_file(hwp_file, output_path)
                
                if success:
                    self.log(f"✓ {msg}")
                    success_count += 1
                else:
                    self.log(f"✗ {msg}", "ERROR")
                    fail_count += 1
                
                self.progress['value'] = idx
                self.root.update_idletasks()
            
            # 변환기 종료
            converter.close()
            
            # 완료 메시지
            self.log("=" * 50)
            self.log(f"변환 완료! 성공: {success_count}, 실패: {fail_count}")
            self.update_status("변환 완료", "green")
            
            messagebox.showinfo("완료", 
                              f"변환이 완료되었습니다.\n\n"
                              f"성공: {success_count}개\n"
                              f"실패: {fail_count}개")
            
        except Exception as e:
            self.log(f"치명적 오류: {str(e)}", "ERROR")
            self.update_status("오류 발생", "red")
            messagebox.showerror("오류", f"예상치 못한 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            self.is_converting = False
            self.convert_btn.config(state=tk.NORMAL)


def main():
    """메인 함수"""
    root = tk.Tk()
    app = HwpConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
