import tkinter as tk
from tkinter import filedialog


class Aplication(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        self.old_file_path = tk.StringVar()
        self.old_file_path.set("请选择旧Excel")
        self.old_file_frame = self.gen_select_file(self.master, "旧Excel", self.old_file_path)
        self.old_file_frame.grid(row=0, column=0, sticky='W')

        self.new_file_path = tk.StringVar()
        self.new_file_path.set("请选择新Excel")
        self.new_file_frame = self.gen_select_file(self.master, "新Excel", self.new_file_path)
        self.new_file_frame.grid(row=1, column=0, sticky='W')

        self.button_line = self.gen_button_line()
        self.button_line.grid(row=2, column=0, sticky="")

    def get_file_path(self, file_path_var):
        file_path_var.set(filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")]))

    def gen_frams(self, father):
        return tk.Frame(father, padx=10, pady=10)

    def gen_select_file(self, father, button_text, file_path):
        frame = self.gen_frams(father)
        button = tk.Button(frame, text=button_text, command=lambda: self.get_file_path(file_path))
        label = tk.Label(frame, textvariable=file_path)
        button.grid(row=0, column=0)
        label.grid(row=0, column=1)
        return frame

    def gen_button(self, father, button_text, command):
        return tk.Button(father, text=button_text, command=command)

    def gen_button_line(self):
        frame = self.gen_frams(self.master)
        button = self.gen_button(frame, "确认", None)
        button.grid(row=0, column=0, padx=10)
        button = self.gen_button(frame, "取消", None)
        button.grid(row=0, column=1, padx=10)
        return frame


if __name__ == '__main__':
    root = tk.Tk()
    root.minsize(300, 100)
    root.title("Excel文件处理")
    app = Aplication(root)
    app.mainloop()
