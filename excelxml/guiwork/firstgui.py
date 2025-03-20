import tkinter as tk
from tkinter import filedialog
def select_directory():
    # 弹出选择目录的对话框
    directory_path = filedialog.askdirectory()
    if directory_path:
        label.config(text=f"选择的目录路径：{directory_path}")

def test():
    # 创建主窗口
    window = tk.Tk()
    window.title("我的第一个Tkinter程序")
    window.geometry("400x300")

    window.option_add("*Font", "微软雅黑 12") 

    # 添加标签（显示文本）
    label = tk.Label(window, text="欢迎学习Tkinter!", font=("微软雅黑", 16))
    label.pack(pady=20)

    # 添加按钮（绑定点击事件）
    def on_click():
        label.config(text="按钮被点击了!")

    def show_input():
        input_text = entry.get()
        result_label.config(text=f"你输入了：{input_text}")

    button = tk.Button(window, text="点击我", command=on_click)
    button.pack()

    entry = tk.Entry(window, width=30)
    entry.pack(pady=10)

    result_label = tk.Label(window, text="")
    result_label.pack()

    submit_btn = tk.Button(window, text="提交", command=show_input)
    submit_btn.pack()

    button = tk.Button(window, text="选择目录", command=select_directory)
    button.pack(pady=20)

    label = tk.Label(window, text="未选择目录")
    label.pack(pady=10)

    # 运行主循环
    window.mainloop()

window = tk.Tk();



def draw():
    print("draw");

    # 添加标签（显示文本）
    mainlabel = tk.Label(window, text="征途2礼包道具生成器!", font=("微软雅黑", 16))
    mainlabel.pack(pady=10)

    gift_from_lbl = tk.Label(window, text="礼包来源:");
    gift_from_lbl.pack(pady=20);
    gift_from_lbl.place(x=30, y=85, anchor=tk.SW)

    entry = tk.Entry(window, width=30)
    entry.place(x=40, y=70, anchor=tk.SW)
    entry.pack(pady=10)


def main():    
    window.title("礼包道具生成器");
    window.geometry("400x300");
    draw();
    window.mainloop();

if __name__ == "__main__":
    #test();
    main();