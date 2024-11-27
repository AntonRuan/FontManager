# 字体管理器
# 界面上面是搜索栏，搜索栏下面是字体列表，右边是字体预览，预览下面是字体信息，方法用tkinter
# 打包指令
# pyinstaller --noconfirm --onefile --windowed --add-data "font.csv;." Font.py


import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont
from tkinter import PhotoImage
import os
from photoshop import Session
import sys
import csv

def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class FontManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("字体管理器 V1.2")
        self.root.geometry("1000x700")
        self.root.configure(bg='#1e1e2e')  # 深色背景 - Catppuccin Mocha 主题色

        # 创建搜索框
        self.search_frame = tk.Frame(self.root, bg='#2c3e50')
        self.search_frame.pack(fill=tk.X, padx=20, pady=15)

        # 搜索图标标签
        self.search_label = tk.Label(self.search_frame, text="🔍", bg='#2c3e50', fg='white')
        self.search_label.pack(side=tk.LEFT, padx=(0,5))

        # 现代风格搜索框
        self.search_entry = tk.Entry(self.search_frame,
                                   bg='#34495e', fg='white', insertbackground='white',
                                   relief=tk.FLAT)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.search_entry.bind('<KeyRelease>', self.search_fonts)

        # 创建按钮框架
        self.button_frame = tk.Frame(self.root, bg='#2c3e50')
        self.button_frame.pack(fill=tk.X, padx=20, pady=(0,15))

        # 添加按钮
        self.add_btn = tk.Button(self.button_frame, text="添加字体",
                               bg='#34495e', fg='white',
                               relief=tk.FLAT, padx=10)
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.remove_btn = tk.Button(self.button_frame, text="移除字体",
                                  bg='#34495e', fg='white',
                                  relief=tk.FLAT, padx=10)
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        self.full_font_btn = tk.Button(self.button_frame, text="全部字体",
                                     bg='#34495e', fg='white',
                                     relief=tk.FLAT, padx=10,
                                     command=self.show_all_fonts)
        self.full_font_btn.pack(side=tk.LEFT, padx=5)

        # 创建主框架
        self.main_frame = tk.Frame(self.root, bg='#2c3e50')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 创建左侧字体列表框架
        self.font_frame = tk.Frame(self.main_frame, bg='#2c3e50')
        self.font_frame.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=(0, 10))

        # 字体列表标题
        self.font_title = tk.Label(self.font_frame, text="字体名称",
                                 bg='#2c3e50', fg='white', font=('Arial', 10))
        self.font_title.pack(pady=(0, 5))

        # 创建字体列表和滚动条的容器框架
        font_container = tk.Frame(self.font_frame, bg='#2c3e50')
        font_container.pack(fill=tk.Y, expand=True)

        # 字体列表
        self.font_list = tk.Listbox(font_container, bg='#34495e', fg='white', relief=tk.FLAT,
                                  selectmode=tk.SINGLE, width=30)
        self.font_list.pack(side=tk.LEFT, fill=tk.Y)
        self.font_list.bind('<<ListboxSelect>>', self.on_select_font)

        # 字体列表滚动条
        font_scrollbar = ttk.Scrollbar(font_container, orient="vertical", command=self.font_list.yview)
        font_scrollbar.pack(side=tk.LEFT, fill=tk.Y, expand=True)
        self.font_list.configure(yscrollcommand=font_scrollbar.set)

        # 创建样式列表框架
        self.style_frame = tk.Frame(self.main_frame, bg='#2c3e50')
        self.style_frame.pack(side=tk.LEFT, fill=tk.Y)

        # 样式列表标题
        self.style_title = tk.Label(self.style_frame, text="字体样式",
                                  bg='#2c3e50', fg='white', font=('Arial', 10))
        self.style_title.pack(pady=(0, 5))

        # 创建样式列表和滚动条的容器框架
        style_container = tk.Frame(self.style_frame, bg='#2c3e50')
        style_container.pack(fill=tk.Y, expand=True)

        # 样式列表
        self.style_list = tk.Listbox(style_container, bg='#34495e', fg='white', relief=tk.FLAT,
                                   selectmode=tk.SINGLE, width=20)
        self.style_list.pack(side=tk.LEFT, fill=tk.Y)
        self.style_list.bind('<<ListboxSelect>>', self.on_select_font)

        # 样式列表滚动条
        style_scrollbar = ttk.Scrollbar(style_container, orient="vertical", command=self.style_list.yview)
        style_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.style_list.configure(yscrollcommand=style_scrollbar.set)

        # 右侧预览区域
        self.preview_frame = tk.Frame(self.main_frame, bg='#34495e', relief=tk.RAISED)
        self.preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20)
        self.preview_frame.pack_propagate(False)  # 防止子组件影响frame大小
        self.preview_frame.configure(width=400, height=500)

        # 预览标题
        self.preview_title = tk.Label(self.preview_frame, text="字体预览",
                                    font=('Arial', 14, 'bold'), bg='#34495e', fg='white')
        self.preview_title.pack(pady=10)

        # 预览文本
        self.preview_text = tk.Text(self.preview_frame, height=8, width=30,
                                  bg='#ecf0f1', fg='#2c3e50', relief=tk.FLAT,
                                  font=('Microsoft YaHei', 12), padx=10, pady=10)
        self.preview_text.pack(fill=tk.BOTH, padx=15, pady=10)
        self.preview_text.insert('1.0', '123456789\nabcdefg\nABCDEFG\n预览文本')

        # 字体信息面板
        self.info_frame = tk.Frame(self.preview_frame, bg='#34495e')
        self.info_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        self.info_label = tk.Label(self.info_frame, text="字体信息",
                                 font=('Microsoft YaHei UI', 12), bg='#34495e', fg='white',
                                 justify=tk.LEFT, wraplength=400)
        self.info_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 存储字体信息的字典
        self.font_info = {}

        self.load_csv_fonts()

        # 填充字体列表
        self.load_ps_fonts()

    def load_fonts(self):
        """加载系统字体"""
        fonts = list(tkFont.families())
        fonts.sort()
        for font in fonts:
            self.font_list.insert(tk.END, font)

    def load_csv_fonts(self):
        """加载CSV中的字体"""
        # 读取csv文件，并通过该文件对字体进行分类
        # 创建字体品牌映射字典和品牌集合
        self.font_brand_map = {}  # {font_name: brand}
        self.brands = set()  # 存储所有品牌

        try:
            # 使用resource_path获取CSV文件路径
            csv_path = resource_path('font.csv')

            with open(csv_path, 'r', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                # 跳过表头
                next(csv_reader, None)

                # 读取字体和对应的商标
                for row in csv_reader:
                    if len(row) >= 4:  # 确保行至少有4列
                        font_name = row[2].strip()
                        brand = row[3].strip()
                        if font_name and brand:  # 确保字体名和商标都不为空
                            self.font_brand_map[font_name] = brand
                            self.brands.add(brand)

                # 为每个品牌创建过滤按钮
                for brand in self.brands:
                    brand_btn = tk.Button(
                        self.button_frame,
                        text=brand,
                        bg='#34495e', fg='white',
                        relief=tk.FLAT, padx=10,
                        command=lambda b=brand: self.filter_brand_fonts(b)
                    )
                    brand_btn.pack(side=tk.LEFT, padx=5)
        except Exception as e:
            print(f"读取CSV文件时出错: {e}")

    def load_ps_fonts(self):
        """加载Photoshop中的字体"""
        try:
            with Session() as ps:
                fonts = ps.app.fonts
                print(f"Photoshop支持 {len(fonts)} 个字体。")

                # 清空列表框
                self.font_list.delete(0, tk.END)

                for font in fonts:
                    try:
                        font_name = font.name
                        font_family = font.family
                        font_style = font.style
                        postscript_name = font.postScriptName
                        brand_name = font.family + " " + font.style
                        brand = self.font_brand_map.get(brand_name.replace(" ", ""), "未知")

                        # 检查字体家族是否已存在
                        if font_family not in self.font_info:
                            # 如果是新的字体家族,创建一个新条目
                            self.font_info[font_family] = {
                                "family": font_family,
                                "styles": {
                                    font_style: {
                                        "name": font_name,
                                        "postscript_name": postscript_name
                                    }
                                },
                                "type": "PS字体",
                                "brand": brand
                            }
                        else:
                            # 如果字体家族已存在,添加新的样式
                            self.font_info[font_family]["styles"][font_style] = {
                                "name": font_name,
                                "postscript_name": postscript_name
                            }
                    except Exception as e:
                        print(f"获取字体名称时出错: {e}")
                        continue

                # 对字体名称进行排序后再添加到列表中
                sorted_fonts = sorted(self.font_info.keys())
                for font_family in sorted_fonts:
                    self.font_list.insert(tk.END, font_family)
        except Exception as e:
            print(f"连接Photoshop时出错: {e}")

    def search_fonts(self, event):
        """搜索字体"""
        search_term = self.search_entry.get().lower()
        self.font_list.delete(0, tk.END)
        self.style_list.delete(0, tk.END)  # 清空样式列表

        # 先对字体名称进行排序
        sorted_fonts = sorted(self.font_info.keys())
        for font_family in sorted_fonts:
            if search_term in font_family.lower():
                self.font_list.insert(tk.END, font_family)

    def show_all_fonts(self):
        self.font_list.delete(0, tk.END)
        sorted_fonts = sorted(self.font_info.keys())
        for font_family in sorted_fonts:
            self.font_list.insert(tk.END, font_family)

    def filter_brand_fonts(self, brand):
        """过滤指定品牌的字体"""
        self.font_list.delete(0, tk.END)
        for font_family in self.font_info:
            if self.font_info[font_family]["brand"] == brand:
                self.font_list.insert(tk.END, font_family)

    def on_select_font(self, event):
        """当选择字体时更新预览和样式列表"""
        font_data = {}
        # 获取选中的字体族
        selection = self.font_list.curselection()
        if selection:
            font_family = self.font_list.get(selection[0])
            print(f"字体选择: {font_family}")
            # 保存当前选中的字体族
            self.current_font_family = font_family
            self.current_font_name = font_family
            self.current_postscript_name = self.font_info[font_family]['styles'][list(self.font_info[font_family]['styles'].keys())[0]]['postscript_name']

            # 更新样式列表
            self.style_list.delete(0, tk.END)
            font_data = self.font_info.get(font_family, {})
            if font_data and "styles" in font_data:
                for style in sorted(font_data["styles"].keys()):
                    self.style_list.insert(tk.END, style)

        style_selection = self.style_list.curselection()
        if style_selection:
            print(f"样式选择: {style_selection}")
            style = self.style_list.get(style_selection[0])
            font_data = self.font_info.get(self.current_font_family, {})
            # 使用保存的字体族获取字体数据
            if font_data and "styles" in font_data and style in font_data["styles"]:
                self.current_font_name = font_data["styles"][style]["name"]
                self.current_postscript_name = font_data["styles"][style]["postscript_name"]
                print(f"字体名称: {self.current_font_name}")

        # 更新预览文本字体
        if self.current_font_name:
            self.preview_text.configure(font=(self.current_font_name, 12))

        # 更新字体信息
        if font_data:
            font_text = f"""字体族: {self.current_font_family}
类型: {font_data['type']}
样式: {font_data['styles'][list(font_data['styles'].keys())[0]]['name']}"""
            self.info_label.configure(text=font_text)

        self.update_ps_font(self.current_postscript_name)

    def update_ps_font(self, postscript_name):
        with Session() as ps:
            doc = ps.active_document
            active_layer = doc.activeLayer

            if active_layer.kind == ps.LayerKind.TextLayer:
                try:
                    text_item = active_layer.textItem
                    text_item.font = postscript_name
                    print(f"Updated font to {postscript_name} for active text layer.")
                except Exception as e:
                    print(f"Error updating font for active text layer: {e}")
                    # 弹窗提示用户
                    win32api.MessageBox(0, f"{e}", "Error", win32con.MB_ICONERROR)

    def run(self):
        """运行程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = FontManager()
    app.run()

