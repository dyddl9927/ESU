import esu_core
from tkinter import messagebox

def main():
    try:
        app = esu_core.ESUAuthTool()
        app.run()
    except Exception as e:
        messagebox.showerror("오류", f"프로그램 실행 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
    