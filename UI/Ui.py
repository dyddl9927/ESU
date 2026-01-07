import tkinter as tk
from Util.Util import Util as U 

class Ui :

    def __init__(self, root) :  
        self.root = root

    # UI 설정
    def setup_ui(self, service):
        # 메인 프레임
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill='both')
        
        # 타이틀
        title_label = tk.Label(main_frame, text="Windows 10 ESU 인증 관리", 
                              font=("Pretendard SemiBold", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # IP 표시
        self.current_ip = U.get_local_ip()
        ip_label = tk.Label(main_frame, text=f"현재 PC IP: {self.current_ip}", 
                           font=("Pretendard SemiBold", 10), fg="gray")
        ip_label.pack(pady=(0, 10))
        
        # Windows 버전 표시
        self.current_version = U.get_windows_build()
        version_label = tk.Label(main_frame, text=f"Windows 버전: {self.current_version}", 
                                font=("Pretendard SemiBold", 9), fg="blue")
        version_label.pack(pady=(0, 20))
        
        # 원본의 cdkey 입력란 주석 처리
        # CD 키 입력 프레임
        # cdkey_frame = tk.Frame(main_frame)
        # cdkey_frame.pack(pady=10, fill='x')
        
        # tk.Label(cdkey_frame, text="CD 키:", font=("맑은 고딕", 10)).pack(side='left', padx=(0, 10))
        # self.cdkey_entry = tk.Entry(cdkey_frame, width=40, font=("맑은 고딕", 10))
        # self.cdkey_entry.pack(side='left', fill='x', expand=True)
        
        # 버튼 프레임
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=30)
        
        # 원본의 A작업 작동 버튼 주석 처리
        # A작업 버튼 - "설치"로 변경
        # self.btn_check = tk.Button(button_frame, text="설치", 
        #                            font=("맑은 고딕", 12, "bold"),
        #                            bg="#4CAF50", fg="white",
        #                            width=15, height=2,
        #                            command=service.check_installation_id)
        # self.btn_check.pack(side='left', padx=10)
        
        # B작업 버튼 - "인증"
        self.btn_auth = tk.Button(button_frame, text="인증", 
                                  font=("Pretendard SemiBold", 12, "bold"),
                                  bg="#2196F3", fg="white",
                                  width=15, height=2,
                                  command=service.activate_esu)
        self.btn_auth.pack(side='left', padx=0)
        
        # 상태 표시
        self.status_label = tk.Label(main_frame, text="준비 완료", 
                                    font=("Pretendard SemiBold", 9),
                                    fg="black")
        self.status_label.pack(pady=(20, 0))

    # 창을 화면 중앙에 배치
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
