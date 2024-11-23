import os
from tkinter import Tk, Label
from tkinterdnd2 import TkinterDnD, DND_FILES
from pptx import Presentation
from pptx.util import Inches


# Markdown 转 PPTX 的函数
def markdown_to_pptx(md_text, output_file):
    # 创建一个新的 PPTX 文件
    presentation = Presentation()
    slide = None

    # 解析 Markdown 文本
    lines = md_text.splitlines()
    for line in lines:
        line = line.rstrip()  # 去掉右侧多余空格
        if line.startswith("# "):
            # 添加标题幻灯片
            slide = presentation.slides.add_slide(presentation.slide_layouts[0])
            slide.shapes.title.text = line[2:]
        elif line.startswith("## "):
            # 添加内容幻灯片，并将内容作为标题
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes.title.text = line[3:]
        elif line.startswith("### "):
            # 添加带子标题的内容幻灯片
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes.title.text = line[4:]
        elif line.strip().startswith(("-", "*")):
            # 项目符号列表项
            if not slide:
                slide = presentation.slides.add_slide(presentation.slide_layouts[1])
                slide.shapes.title.text = "Bullet Points"

            # 确保有文本框用于放置项目符号
            if not slide.shapes.placeholders[1].has_text_frame:
                text_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5))
                text_frame = text_box.text_frame
            else:
                text_frame = slide.shapes.placeholders[1].text_frame

            # 计算当前行的缩进级别
            current_level = (len(line) - len(line.lstrip())) // 2
            content = line.strip("-* ").strip()

            # 添加项目符号，根据级别设置层次
            p = text_frame.add_paragraph()
            p.text = content
            p.level = current_level

    # 保存 PPTX 文件
    presentation.save(output_file)

# 文件处理函数
def process_file(file_path):
    try:
        # 获取文件名和目标输出路径
        base_name = os.path.basename(file_path)
        output_file = os.path.splitext(base_name)[0] + ".pptx"

        # 读取 Markdown 文件
        with open(file_path, 'r', encoding='utf-8') as file:
            markdown_text = file.read()

        # 执行转换
        markdown_to_pptx(markdown_text, output_file)
        status_label.config(text=f"Successfully converted to {output_file}")
    except Exception as e:
        status_label.config(text=f"Failed to convert: {e}")

# 初始化 GUI
class MarkdownToPptxApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Markdown to PPTX Converter")
        self.geometry("600x200")

        # 设置背景为黑色，文字为白色
        self.configure(bg="black")

        # 设置标题标签背景为黑色，文字为白色
        label = Label(self, text="Drop Your Markdown File Here", font=("Arial", 24), fg="white", bg="black")
        label.pack(pady=10)

        label = Label(self, text="Editing by Andy Chang, Network Computing Lab., NCKU, Taiwan", font=("Arial", 10), fg="white", bg="black")
        label.pack(pady=20)
        # 设置状态标签背景为黑色，文字为白色
        global status_label
        status_label = Label(self, text="", font=("Arial", 20), fg="white", bg="black")
        status_label.pack(pady=10)

        # 启用拖放功能
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop)

    def drop(self, event):
        file_path = event.data.strip("{").strip("}")  # 去掉多余的括号
        process_file(file_path)

# 运行应用程序
if __name__ == "__main__":
    app = MarkdownToPptxApp()
    app.mainloop()
