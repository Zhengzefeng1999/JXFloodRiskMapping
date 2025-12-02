import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pyproj import Transformer, CRS
import os

class CoordinateConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CGCS2000坐标转换工具")
        self.root.geometry("600x500")
        
        # 存储文件路径
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        
        # 源和目标坐标系变量
        self.source_crs = tk.StringVar()
        self.target_crs = tk.StringVar()
        
        # CGCS2000经度带选项
        self.cgcs2000_zones = [
            "CGCS2000 108°E", "CGCS2000 111°E", "CGCS2000 114°E", 
            "CGCS2000 117°E", "CGCS2000 120°E", "CGCS2000 123°E"
        ]
        self.all_crs_options = self.cgcs2000_zones + ["WGS1984"]
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="CGCS2000坐标转换工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 第一步：选择Excel文件
        step1_frame = tk.LabelFrame(self.root, text="第一步：选择输入Excel文件", padx=10, pady=10)
        step1_frame.pack(fill="x", padx=20, pady=10)
        
        file_frame = tk.Frame(step1_frame)
        file_frame.pack(fill="x")
        
        tk.Entry(file_frame, textvariable=self.input_file_path, width=50).pack(side="left", fill="x", expand=True)
        tk.Button(file_frame, text="浏览...", command=self.select_input_file).pack(side="right", padx=(5, 0))
        
        # 第二步：选择源坐标系统
        step2_frame = tk.LabelFrame(self.root, text="第二步：选择源坐标系统", padx=10, pady=10)
        step2_frame.pack(fill="x", padx=20, pady=10)
        
        source_frame = tk.Frame(step2_frame)
        source_frame.pack(fill="x")
        
        tk.Label(source_frame, text="源坐标系统:").pack(side="left")
        self.source_combobox = ttk.Combobox(source_frame, textvariable=self.source_crs, 
                                           values=self.all_crs_options, state="readonly", width=30)
        self.source_combobox.pack(side="left", padx=(10, 0))
        self.source_combobox.set("请选择源坐标系统")
        
        # 第三步：选择目标坐标系统
        step3_frame = tk.LabelFrame(self.root, text="第三步：选择目标坐标系统", padx=10, pady=10)
        step3_frame.pack(fill="x", padx=20, pady=10)
        
        target_frame = tk.Frame(step3_frame)
        target_frame.pack(fill="x")
        
        tk.Label(target_frame, text="目标坐标系统:").pack(side="left")
        self.target_combobox = ttk.Combobox(target_frame, textvariable=self.target_crs, 
                                           values=self.all_crs_options, state="readonly", width=30)
        self.target_combobox.pack(side="left", padx=(10, 0))
        self.target_combobox.set("请选择目标坐标系统")
        
        # 第四步：选择输出文件路径
        step4_frame = tk.LabelFrame(self.root, text="第四步：设置输出文件", padx=10, pady=10)
        step4_frame.pack(fill="x", padx=20, pady=10)
        
        output_frame = tk.Frame(step4_frame)
        output_frame.pack(fill="x")
        
        tk.Entry(output_frame, textvariable=self.output_file_path, width=50).pack(side="left", fill="x", expand=True)
        tk.Button(output_frame, text="保存为...", command=self.select_output_file).pack(side="right", padx=(5, 0))
        
        # 转换按钮
        convert_button = tk.Button(self.root, text="开始转换", command=self.convert_coordinates, 
                                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2, width=20)
        convert_button.pack(pady=20)
        
        # 日志文本框
        log_frame = tk.LabelFrame(self.root, text="转换日志", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8)
        scrollbar = tk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.log_message(f"已选择输入文件: {file_path}")
            
    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="保存转换结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            self.output_file_path.set(file_path)
            self.log_message(f"已设置输出文件: {file_path}")
            
    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def get_cgcs2000_crs_string(self, zone_name):
        """根据经度带名称生成CGCS2000 CRS字符串"""
        # 提取经度值
        lon_value = int(zone_name.split(" ")[1].replace("°E", ""))
        
        # CGCS2000 3度带投影参数
        crs_string = (
            f"+proj=tmerc +lat_0=0 +lon_0={lon_value} "
            f"+k=1 +x_0=500000 +y_0=0 "
            f"+ellps=GRS80 +units=m +no_defs"
        )
        return crs_string
        
    def get_crs_definition(self, crs_name):
        """获取坐标系定义"""
        if crs_name == "WGS1984":
            return "EPSG:4326"
        elif crs_name.startswith("CGCS2000"):
            return self.get_cgcs2000_crs_string(crs_name)
        else:
            return None
            
    def read_excel_file(self, file_path):
        """增强的Excel文件读取函数，支持多种格式和错误处理"""
        try:
            # 首先尝试使用openpyxl引擎读取.xlsx文件
            if file_path.endswith('.xlsx'):
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    return df
                except Exception as e:
                    self.log_message(f"使用openpyxl读取失败: {str(e)}")
                    # 如果openpyxl失败，尝试使用xlrd引擎
                    try:
                        df = pd.read_excel(file_path, engine='xlrd')
                        return df
                    except Exception as e2:
                        self.log_message(f"使用xlrd读取失败: {str(e2)}")
                        raise e2
            # 对于.xls文件，使用xlrd引擎
            elif file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
                return df
            else:
                # 尝试自动检测
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    return df
                except:
                    df = pd.read_excel(file_path, engine='xlrd')
                    return df
        except Exception as e:
            self.log_message(f"读取Excel文件时出错: {str(e)}")
            raise e
            
    def convert_coordinates(self):
        # 检查输入
        if not self.input_file_path.get():
            messagebox.showerror("错误", "请先选择输入Excel文件")
            return
            
        if self.source_crs.get() not in self.all_crs_options:
            messagebox.showerror("错误", "请选择有效的源坐标系统")
            return
            
        if self.target_crs.get() not in self.all_crs_options:
            messagebox.showerror("错误", "请选择有效的目标坐标系统")
            return
            
        if not self.output_file_path.get():
            messagebox.showerror("错误", "请设置输出文件路径")
            return
            
        # 获取坐标系定义
        source_crs_def = self.get_crs_definition(self.source_crs.get())
        target_crs_def = self.get_crs_definition(self.target_crs.get())
        
        if not source_crs_def or not target_crs_def:
            messagebox.showerror("错误", "无法识别的坐标系统")
            return
            
        try:
            self.log_message("开始读取Excel文件...")
            
            # 读取Excel文件，使用增强的读取函数
            df = self.read_excel_file(self.input_file_path.get())
            
            # 检查必要的列是否存在
            required_columns = ['X', 'Y']
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                messagebox.showerror("错误", f"Excel文件缺少必要的列: {missing_cols}")
                return
                
            self.log_message(f"成功读取数据，共{len(df)}行")
            
            # 检查数据有效性
            if df.empty:
                messagebox.showerror("错误", "Excel文件为空或没有有效数据")
                return
                
            # 创建坐标转换器
            self.log_message(f"源坐标系: {self.source_crs.get()}")
            self.log_message(f"目标坐标系: {self.target_crs.get()}")
            
            transformer = Transformer.from_crs(source_crs_def, target_crs_def, always_xy=True)
            
            # 执行坐标转换
            self.log_message("正在进行坐标转换...")
            
            # 确保X和Y列是数值类型
            df['X'] = pd.to_numeric(df['X'], errors='coerce')
            df['Y'] = pd.to_numeric(df['Y'], errors='coerce')
            
            # 检查是否有无效数据
            invalid_rows = df[df['X'].isna() | df['Y'].isna()]
            if not invalid_rows.empty:
                self.log_message(f"警告: 发现{len(invalid_rows)}行无效数据，将被忽略")
                df = df.dropna(subset=['X', 'Y'])
                if df.empty:
                    messagebox.showerror("错误", "所有数据都是无效的，请检查输入文件")
                    return
            
            # 记录原始列名（除了X和Y）
            original_columns = [col for col in df.columns if col not in ['X', 'Y']]
            
            # 执行坐标转换
            x_new, y_new = transformer.transform(df['X'].values, df['Y'].values)
            
            # 创建新的DataFrame，只保留原始列（除了X和Y）以及新的X_new和Y_new列
            new_df = df[original_columns].copy()
            new_df['X_new'] = x_new
            new_df['Y_new'] = y_new
            
            # 保存结果
            output_path = self.output_file_path.get()
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                
            # 写入Excel文件
            self.log_message("正在保存结果...")
            new_df.to_excel(output_path, index=False, engine='openpyxl')
            
            self.log_message(f"转换完成！结果已保存到: {output_path}")
            self.log_message("注意：原X和Y列已被删除，只保留了X_new和Y_new列")
            messagebox.showinfo("成功", f"坐标转换完成！\n结果已保存到: {output_path}\n注意：原X和Y列已被删除，只保留了X_new和Y_new列")
            
        except Exception as e:
            error_msg = f"转换过程中发生错误: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("错误", error_msg)

def main():
    root = tk.Tk()
    app = CoordinateConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()