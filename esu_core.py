import tkinter as tk
from UI.Ui import Ui
from Util.Util import Util as U
from Service.Service import Service as S


class ESUAuthTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.iconbitmap(default=U._default_icon_path())
        self.root.title("Windows 10 ESU 인증 툴 - Soft11")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        
        self.ui = Ui(self.root)
        self.util = U(self.root)
        if not self.util.is_admin():
            self.util.request_admin()
            return

        self.service = S(self.root,self.ui,self.util)
        self.ui.center_window()
        
        self.ui.setup_ui(self.service)
        self.root.after(0, lambda: self.service.start_task(self.service.initialize))
        
        
    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            print(f"Error running application: {e}")
