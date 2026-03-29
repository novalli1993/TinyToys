import tkinter as tk


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("基本图形界面")
        self.geometry("300x100")

    def gen_frame(self, root):
        frame = tk.Frame(master=root)
        return frame


if __name__ == '__main__':
    app = Application()
    app.mainloop()
