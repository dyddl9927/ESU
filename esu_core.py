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
        self.service = S(self.root,self.ui,self.util)
        
        #self.excel_path = r"c:\1312\esu.xlsx"
        self.service.ensure_excel_exists()
        self.ui.center_window()
        if not self.util.is_admin():
            self.util.request_admin()
            return
        
        self.ui.setup_ui(self.service)
        self.service.check_installation_id()
        
        
    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            print(f"Error running application: {e}")
