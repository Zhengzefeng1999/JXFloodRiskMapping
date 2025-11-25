import os
import geopandas as gpd
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def extract_shp_to_excel(main_folder, output_file):
    """改进后的功能函数：提取shp文件到Excel，支持多种编码格式"""
    # 创建Excel写入对象
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # 记录成功处理的文件夹数量
    processed_count = 0
    
    # 首先检查主文件夹中是否直接包含SHP文件
    shp_path = os.path.join(main_folder, "2D Zones.shp")
    if os.path.exists(shp_path):
        try:
            # 尝试不同的编码格式读取SHP文件
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
            gdf = None
            last_error = None
            
            for encoding in encodings:
                try:
                    gdf = gpd.read_file(shp_path, encoding=encoding)
                    print(f"使用编码 {encoding} 成功读取主文件夹中的SHP文件")
                    break
                except Exception as e:
                    last_error = e
                    continue
                    
            if gdf is None:
                raise last_error
                
            # 移除几何列并写入Excel
            df = pd.DataFrame(gdf.drop(columns='geometry'))
            df.to_excel(writer, sheet_name='Main_Folder', index=False)
            processed_count += 1
            print("成功处理主文件夹中的SHP文件")
        except Exception as e:
            print(f"处理主文件夹中的SHP文件时发生错误：{str(e)}")
    else:
        # 遍历主文件夹下的所有子文件夹
        for folder_name in os.listdir(main_folder):
            folder_path = os.path.join(main_folder, folder_name)
            
            # 跳过非文件夹的文件
            if not os.path.isdir(folder_path):
                continue
                
            # 构建shp文件路径
            shp_path = os.path.join(folder_path, "2D Zones.shp")
            
            # 检查shp文件是否存在
            if not os.path.exists(shp_path):
                print(f"警告：在文件夹 {folder_name} 中未找到2D Zones.shp")
                continue
                
            try:
                # 尝试不同的编码格式读取SHP文件
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                gdf = None
                last_error = None
                
                for encoding in encodings:
                    try:
                        gdf = gpd.read_file(shp_path, encoding=encoding)
                        print(f"使用编码 {encoding} 成功读取文件夹 {folder_name} 中的SHP文件")
                        break
                    except Exception as e:
                        last_error = e
                        continue
                        
                if gdf is None:
                    raise last_error
                
                # 移除几何列（如果需要保留可以注释这行）
                df = pd.DataFrame(gdf.drop(columns='geometry'))
                
                # 将数据写入Excel的sheet
                df.to_excel(writer, sheet_name=folder_name, index=False)
                processed_count += 1
                
                print(f"成功处理文件夹：{folder_name}")
                
            except Exception as e:
                print(f"处理 {folder_name} 时发生错误：{str(e)}")
                continue

    # 如果没有处理任何文件夹，则创建一个默认的工作表
    if processed_count == 0:
        # 创建一个空的DataFrame作为默认工作表
        default_df = pd.DataFrame({'提示': ['未找到任何有效的"2D Zones.shp"文件']})
        default_df.to_excel(writer, sheet_name='处理结果', index=False)
        print("未找到任何有效的SHP文件，已创建默认工作表")
    else:
        print(f"共处理了 {processed_count} 个文件夹")

    # 保存并关闭Excel文件
    writer.close()
    print(f"\n处理完成！输出文件已保存至：{output_file}")
    return f"处理完成！输出文件已保存至：{output_file}"

class SHPExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SHP文件提取工具")
        self.root.geometry("600x300")
        
        # 输入文件夹路径
        self.input_folder = tk.StringVar()
        # 输出文件路径
        self.output_file = tk.StringVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="SHP文件提取到Excel工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 输入文件夹选择
        input_label = ttk.Label(main_frame, text="输入文件夹:")
        input_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        input_entry = ttk.Entry(main_frame, textvariable=self.input_folder, width=50)
        input_entry.grid(row=1, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        input_button = ttk.Button(main_frame, text="浏览...", command=self.select_input_folder)
        input_button.grid(row=1, column=2, pady=5)
        
        # 输出文件选择
        output_label = ttk.Label(main_frame, text="输出文件:")
        output_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        output_entry = ttk.Entry(main_frame, textvariable=self.output_file, width=50)
        output_entry.grid(row=2, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        output_button = ttk.Button(main_frame, text="浏览...", command=self.select_output_file)
        output_button.grid(row=2, column=2, pady=5)
        
        # 执行按钮
        execute_button = ttk.Button(main_frame, text="执行提取", command=self.execute_extraction)
        execute_button.grid(row=3, column=1, pady=20)
        
        # 日志文本框
        log_label = ttk.Label(main_frame, text="处理日志:")
        log_label.grid(row=4, column=0, sticky=(tk.W, tk.S), pady=(10, 0))
        
        self.log_text = tk.Text(main_frame, height=8, width=70)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=(5, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def select_input_folder(self):
        """选择输入文件夹"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder.set(folder_selected)
            
    def select_output_file(self):
        """选择输出文件"""
        file_selected = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_selected:
            self.output_file.set(file_selected)
            
    def execute_extraction(self):
        """执行提取操作"""
        input_folder = self.input_folder.get()
        output_file = self.output_file.get()
        
        # 验证输入
        if not input_folder:
            messagebox.showerror("错误", "请选择输入文件夹")
            return
            
        if not output_file:
            messagebox.showerror("错误", "请选择输出文件路径")
            return
            
        if not os.path.exists(input_folder):
            messagebox.showerror("错误", "输入文件夹不存在")
            return
            
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        try:
            # 重定向print输出到文本框
            import sys
            import io
            
            # 保存原始stdout
            original_stdout = sys.stdout
            
            # 创建StringIO对象捕获输出
            captured_output = io.StringIO()
            sys.stdout = captured_output
            
            # 执行提取功能
            result = extract_shp_to_excel(input_folder, output_file)
            
            # 恢复原始stdout
            sys.stdout = original_stdout
            
            # 获取捕获的输出
            output_text = captured_output.getvalue()
            
            # 显示在日志区域
            self.log_text.insert(tk.END, output_text)
            self.log_text.insert(tk.END, "\n" + result)
            self.log_text.see(tk.END)  # 自动滚动到底部
            
            # 强制更新界面
            self.log_text.update_idletasks()
            
            # 显示完成消息
            messagebox.showinfo("完成", "文件提取完成！")
            
        except Exception as e:
            # 恢复原始stdout
            sys.stdout = original_stdout
            
            error_msg = f"处理过程中发生错误: {str(e)}"
            self.log_text.insert(tk.END, error_msg)
            self.log_text.see(tk.END)  # 自动滚动到底部
            # 强制更新界面
            self.log_text.update_idletasks()
            messagebox.showerror("错误", error_msg)

def main():
    root = tk.Tk()
    app = SHPExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()