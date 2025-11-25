import os
import geopandas as gpd
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path


def extract_shp_to_excel(main_folder, output_file, clip_mesh=False, mesh_shp_file=None, clipped_shp_output_folder=None):
    """改进后的功能函数：提取shp文件到Excel，包含指定列和计算列，支持裁剪处理"""
    # 定义需要提取的原始列
    required_columns = ['element_no', 'AREA2D', 'DEPTH2D', 'T_FLOOD_DU', 'T_INUDATIO', 'T_PEAK_2D']
    
    # 创建Excel写入对象
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # 记录成功处理的文件夹数量
    processed_count = 0
    
    # 如果需要裁剪网格化区间，先读取网格化区间文件
    clip_gdf = None
    if clip_mesh and mesh_shp_file and os.path.exists(mesh_shp_file):
        try:
            clip_gdf = gpd.read_file(mesh_shp_file)
            print(f"成功加载网格化区间文件")
        except Exception as e:
            print(f"读取网格化区间文件时出错：{str(e)}")
            clip_gdf = None
    
    # 如果需要输出裁剪后的shp文件，确保输出文件夹存在
    output_shp_dir = None
    if clip_mesh and clipped_shp_output_folder:
        output_shp_dir = Path(clipped_shp_output_folder)
        output_shp_dir.mkdir(parents=True, exist_ok=True)
        print(f"裁剪后的shp文件将保存至：{output_shp_dir}")
    
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
                
            # 如果需要裁剪，则执行裁剪操作
            if clip_mesh and clip_gdf is not None:
                try:
                    # 确保坐标系一致
                    if gdf.crs != clip_gdf.crs:
                        gdf = gdf.to_crs(clip_gdf.crs)
                    
                    # 合并所有裁剪几何
                    total_clip = clip_gdf.unary_union
                    
                    # 执行擦除操作（从gdf中移除与clip_gdf重叠的部分）
                    gdf['geometry'] = gdf.geometry.difference(total_clip)
                    gdf = gdf[~gdf.is_empty]  # 过滤空几何
                    
                    # 如果需要输出裁剪后的shp文件
                    if output_shp_dir:
                        clipped_shp_path = output_shp_dir / "Main_Folder_Clipped.shp"
                        gdf.to_file(clipped_shp_path)
                        print(f"主文件夹裁剪后的shp文件已保存至：{clipped_shp_path}")
                        
                except Exception as e:
                    print(f"裁剪主文件夹中的SHP文件时出错：{str(e)}")
            
            # 移除几何列
            df = pd.DataFrame(gdf.drop(columns='geometry'))
            
            # 只保留指定的列，如果列不存在则跳过
            available_columns = [col for col in required_columns if col in df.columns]
            if available_columns:
                df_filtered = df[available_columns].copy()
                
                # 添加计算列
                if 'AREA2D' in df_filtered.columns:
                    df_filtered['淹没面积(km2)'] = df_filtered['AREA2D'] / 1000000
                if 'DEPTH2D' in df_filtered.columns:
                    df_filtered['淹没水深(m)'] = df_filtered['DEPTH2D']
                if 'T_FLOOD_DU' in df_filtered.columns:
                    df_filtered['淹没历时(h)'] = df_filtered['T_FLOOD_DU'] / 3600
                if 'T_INUDATIO' in df_filtered.columns:
                    df_filtered['洪水到达时间(h)'] = df_filtered['T_INUDATIO'] / 3600
                
                df_filtered.to_excel(writer, sheet_name='Main_Folder', index=False)
                processed_count += 1
                print(f"成功处理主文件夹中的SHP文件，提取列: {available_columns}")
            else:
                # 如果没有指定的列，创建一个包含提示信息的工作表
                warning_df = pd.DataFrame({'提示': ['未找到指定的列']})
                warning_df.to_excel(writer, sheet_name='Main_Folder', index=False)
                print("主文件夹中的SHP文件未包含指定的列")
                
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
                
                # 如果需要裁剪，则执行裁剪操作
                if clip_mesh and clip_gdf is not None:
                    try:
                        # 确保坐标系一致
                        if gdf.crs != clip_gdf.crs:
                            gdf = gdf.to_crs(clip_gdf.crs)
                        
                        # 合并所有裁剪几何
                        total_clip = clip_gdf.unary_union
                        
                        # 执行擦除操作（从gdf中移除与clip_gdf重叠的部分）
                        gdf['geometry'] = gdf.geometry.difference(total_clip)
                        gdf = gdf[~gdf.is_empty]  # 过滤空几何
                        
                        # 如果需要输出裁剪后的shp文件
                        if output_shp_dir:
                            # 处理文件名，确保符合Windows文件命名规范
                            safe_folder_name = "".join(c for c in folder_name if c.isalnum() or c in (' ','.','_','-')).rstrip()
                            clipped_shp_path = output_shp_dir / f"{safe_folder_name}_Clipped.shp"
                            gdf.to_file(clipped_shp_path)
                            print(f"{folder_name} 裁剪后的shp文件已保存至：{clipped_shp_path}")
                            
                    except Exception as e:
                        print(f"裁剪 {folder_name} 中的SHP文件时出错：{str(e)}")
                
                # 移除几何列
                df = pd.DataFrame(gdf.drop(columns='geometry'))
                
                # 只保留指定的列，如果列不存在则跳过
                available_columns = [col for col in required_columns if col in df.columns]
                if available_columns:
                    df_filtered = df[available_columns].copy()
                    
                    # 添加计算列
                    if 'AREA2D' in df_filtered.columns:
                        df_filtered['淹没面积(km2)'] = df_filtered['AREA2D'] / 1000000
                    if 'DEPTH2D' in df_filtered.columns:
                        df_filtered['淹没水深(m)'] = df_filtered['DEPTH2D']
                    if 'T_FLOOD_DU' in df_filtered.columns:
                        df_filtered['淹没历时(h)'] = df_filtered['T_FLOOD_DU'] / 3600
                    if 'T_INUDATIO' in df_filtered.columns:
                        df_filtered['洪水到达时间(h)'] = df_filtered['T_INUDATIO'] / 3600
                    
                    # 将数据写入Excel的sheet
                    df_filtered.to_excel(writer, sheet_name=folder_name, index=False)
                    processed_count += 1
                    print(f"成功处理文件夹：{folder_name}，提取列: {available_columns}")
                else:
                    # 如果没有指定的列，创建一个包含提示信息的工作表
                    warning_df = pd.DataFrame({'提示': ['未找到指定的列']})
                    warning_df.to_excel(writer, sheet_name=folder_name, index=False)
                    print(f"文件夹 {folder_name} 中的SHP文件未包含指定的列")
                
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


def generate_analysis_table(input_excel, output_excel, project_object):
    """生成分析统计表格，所有方案放在一个sheet中，并输出最大淹没水深和最大淹没历时对应的element_no"""
    try:
        # 读取Excel文件
        excel_file = pd.ExcelFile(input_excel)
        sheet_names = excel_file.sheet_names
        
        # 创建新的Excel写入对象
        writer = pd.ExcelWriter(output_excel, engine='openpyxl')
        
        # 用于存储所有方案的数据
        all_analysis_data = []
        max_depth_records = []
        max_duration_records = []
        
        # 遍历每个工作表
        for sheet_name in sheet_names:
            # 读取工作表数据
            df = pd.read_excel(input_excel, sheet_name=sheet_name)
            
            # 检查必需的列是否存在
            if '淹没水深(m)' not in df.columns or '淹没面积(km2)' not in df.columns:
                print(f"工作表 {sheet_name} 缺少必需的列")
                continue
                
            # 计算总淹没面积
            total_area = df['淹没面积(km2)'].sum()
            
            # 初始化分层统计
            depth_levels = [
                ("<0.5m", df[df['淹没水深(m)'] < 0.5]),
                ("0.5~1.0m", df[(df['淹没水深(m)'] >= 0.5) & (df['淹没水深(m)'] < 1.0)]),
                ("1.0~2.0m", df[(df['淹没水深(m)'] >= 1.0) & (df['淹没水深(m)'] < 2.0)]),
                ("2.0~3.0m", df[(df['淹没水深(m)'] >= 2.0) & (df['淹没水深(m)'] < 3.0)]),
                (">3.0m", df[df['淹没水深(m)'] >= 3.0])
            ]
            
            # 创建分析结果数据
            for level_name, level_data in depth_levels:
                area = level_data['淹没面积(km2)'].sum()
                ratio = area / total_area if total_area > 0 else 0
                all_analysis_data.append({
                    '编制对象': project_object,
                    '方案名称': sheet_name,
                    '淹没水深(m)': level_name,
                    '淹没面积(km2)': round(area, 4),
                    '占比': f"{ratio:.2%}"
                })
            
            # 计算最大淹没水深及其对应的element_no
            max_depth = df['淹没水深(m)'].max() if not df.empty else 0
            if max_depth > 0:
                max_depth_elements = df[df['淹没水深(m)'] == max_depth]['element_no'].tolist()
                for element in max_depth_elements:
                    max_depth_records.append({
                        '编制对象': project_object,
                        '方案名称': sheet_name,
                        '最大值类型': '最大淹没水深',
                        '最大值': max_depth,
                        'element_no': element
                    })
            
            # 计算最大淹没历时及其对应的element_no
            max_duration = df['淹没历时(h)'].max() if '淹没历时(h)' in df.columns and not df.empty else 0
            if max_duration > 0:
                max_duration_elements = df[df['淹没历时(h)'] == max_duration]['element_no'].tolist()
                for element in max_duration_elements:
                    max_duration_records.append({
                        '编制对象': project_object,
                        '方案名称': sheet_name,
                        '最大值类型': '最大淹没历时',
                        '最大值': max_duration,
                        'element_no': element
                    })
        
        # 创建分析结果DataFrame
        if all_analysis_data:
            analysis_df = pd.DataFrame(all_analysis_data)
            
            # 添加最大值记录到分析表格
            for record in max_depth_records:
                analysis_df = pd.concat([analysis_df, pd.DataFrame([record])], ignore_index=True)
            
            for record in max_duration_records:
                analysis_df = pd.concat([analysis_df, pd.DataFrame([record])], ignore_index=True)
            
            # 将所有数据写入一个sheet
            analysis_df.to_excel(writer, sheet_name='分析结果', index=False)
            
            print("成功生成分析表格，所有方案数据已合并到一个sheet中")
        else:
            # 如果没有数据，创建一个默认的工作表
            default_df = pd.DataFrame({'提示': ['未找到有效的分析数据']})
            default_df.to_excel(writer, sheet_name='分析结果', index=False)
            print("未找到有效的分析数据")
        
        # 保存并关闭Excel文件
        writer.close()
        print(f"\n分析表格生成完成！输出文件已保存至：{output_excel}")
        return f"分析表格生成完成！输出文件已保存至：{output_excel}"
        
    except Exception as e:
        error_msg = f"生成分析表格时发生错误: {str(e)}"
        print(error_msg)
        return error_msg


class SHPExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SHP文件提取工具")
        self.root.geometry("700x600")  # 增加高度以容纳新控件
        
        # 输入文件夹路径
        self.input_folder = tk.StringVar()
        # 输出文件路径
        self.output_file = tk.StringVar()
        # 分析表格输出路径
        self.analysis_output_file = tk.StringVar()
        # 编制对象
        self.project_object = tk.StringVar()
        # 网格化区间SH文件路径
        self.mesh_shp_file = tk.StringVar()
        # 裁剪后shp文件输出文件夹路径
        self.clipped_shp_output_folder = tk.StringVar()
        # 是否裁剪网格化区间
        self.clip_mesh = tk.BooleanVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="SHP文件提取工具", font=("Arial", 16, "bold"))
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
        
        # 是否裁剪网格化区间选项
        clip_mesh_check = ttk.Checkbutton(
            main_frame, 
            text="裁剪网格化区间", 
            variable=self.clip_mesh,
            command=self.toggle_mesh_selection
        )
        clip_mesh_check.grid(row=3, column=0, sticky=tk.W, pady=5)
        
        # 网格化区间SH文件选择（初始禁用）
        self.mesh_shp_label = ttk.Label(main_frame, text="网格化区间文件:")
        self.mesh_shp_label.grid(row=4, column=0, sticky=tk.W, pady=5)
        self.mesh_shp_label.config(state='disabled')
        
        self.mesh_shp_entry = ttk.Entry(main_frame, textvariable=self.mesh_shp_file, width=50, state='disabled')
        self.mesh_shp_entry.grid(row=4, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        self.mesh_shp_button = ttk.Button(main_frame, text="浏览...", command=self.select_mesh_shp_file, state='disabled')
        self.mesh_shp_button.grid(row=4, column=2, pady=5)
        
        # 裁剪后shp文件输出文件夹选择（初始禁用）
        self.clipped_shp_output_label = ttk.Label(main_frame, text="裁剪后文件输出:")
        self.clipped_shp_output_label.grid(row=5, column=0, sticky=tk.W, pady=5)
        self.clipped_shp_output_label.config(state='disabled')
        
        self.clipped_shp_output_entry = ttk.Entry(main_frame, textvariable=self.clipped_shp_output_folder, width=50, state='disabled')
        self.clipped_shp_output_entry.grid(row=5, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        self.clipped_shp_output_button = ttk.Button(main_frame, text="浏览...", command=self.select_clipped_shp_output_folder, state='disabled')
        self.clipped_shp_output_button.grid(row=5, column=2, pady=5)
        
        # 编制对象输入
        project_label = ttk.Label(main_frame, text="编制对象:")
        project_label.grid(row=6, column=0, sticky=tk.W, pady=5)
        
        project_entry = ttk.Entry(main_frame, textvariable=self.project_object, width=50)
        project_entry.grid(row=6, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        # 分析表格输出文件选择
        analysis_output_label = ttk.Label(main_frame, text="分析表格输出:")
        analysis_output_label.grid(row=7, column=0, sticky=tk.W, pady=5)
        
        analysis_output_entry = ttk.Entry(main_frame, textvariable=self.analysis_output_file, width=50)
        analysis_output_entry.grid(row=7, column=1, padx=(10, 10), pady=5, sticky=(tk.W, tk.E))
        
        analysis_output_button = ttk.Button(main_frame, text="浏览...", command=self.select_analysis_output_file)
        analysis_output_button.grid(row=7, column=2, pady=5)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=3, pady=20)
        
        # 执行按钮
        execute_button = ttk.Button(button_frame, text="执行提取", command=self.execute_extraction)
        execute_button.grid(row=0, column=0, padx=(0, 10))
        
        # 生成分析表格按钮
        analysis_button = ttk.Button(button_frame, text="生成分析表格", command=self.generate_analysis)
        analysis_button.grid(row=0, column=1, padx=(10, 0))
        
        # 日志文本框
        log_label = ttk.Label(main_frame, text="处理日志:")
        log_label.grid(row=9, column=0, sticky=(tk.W, tk.S), pady=(10, 0))
        
        self.log_text = tk.Text(main_frame, height=12, width=80)
        self.log_text.grid(row=10, column=0, columnspan=3, pady=(5, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=10, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(10, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def toggle_mesh_selection(self):
        """切换网格化区间文件选择的可用性"""
        if self.clip_mesh.get():
            self.mesh_shp_label.config(state='normal')
            self.mesh_shp_entry.config(state='normal')
            self.mesh_shp_button.config(state='normal')
            self.clipped_shp_output_label.config(state='normal')
            self.clipped_shp_output_entry.config(state='normal')
            self.clipped_shp_output_button.config(state='normal')
        else:
            self.mesh_shp_label.config(state='disabled')
            self.mesh_shp_entry.config(state='disabled')
            self.mesh_shp_button.config(state='disabled')
            self.clipped_shp_output_label.config(state='disabled')
            self.clipped_shp_output_entry.config(state='disabled')
            self.clipped_shp_output_button.config(state='disabled')
            
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
            
    def select_mesh_shp_file(self):
        """选择网格化区间SH文件"""
        file_selected = filedialog.askopenfilename(
            filetypes=[("Shapefile", "*.shp"), ("所有文件", "*.*")]
        )
        if file_selected:
            self.mesh_shp_file.set(file_selected)
            
    def select_clipped_shp_output_folder(self):
        """选择裁剪后shp文件输出文件夹"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.clipped_shp_output_folder.set(folder_selected)
            
    def select_analysis_output_file(self):
        """选择分析表格输出文件"""
        file_selected = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_selected:
            self.analysis_output_file.set(file_selected)
            
    def execute_extraction(self):
        """执行提取操作"""
        input_folder = self.input_folder.get()
        output_file = self.output_file.get()
        clip_mesh = self.clip_mesh.get()
        mesh_shp_file = self.mesh_shp_file.get()
        clipped_shp_output_folder = self.clipped_shp_output_folder.get()
        
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
            
        # 如果选择了裁剪网格化区间，但没有选择网格化区间文件
        if clip_mesh and not mesh_shp_file:
            messagebox.showerror("错误", "请选择网格化区间SH文件")
            return
            
        if clip_mesh and not os.path.exists(mesh_shp_file):
            messagebox.showerror("错误", "网格化区间SH文件不存在")
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
            result = extract_shp_to_excel(input_folder, output_file, clip_mesh, mesh_shp_file, clipped_shp_output_folder)
            
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
            
    def generate_analysis(self):
        """生成分析表格"""
        input_file = self.output_file.get()  # 使用提取功能的输出文件作为输入
        output_file = self.analysis_output_file.get()
        project_object = self.project_object.get()
        
        # 验证输入
        if not input_file:
            messagebox.showerror("错误", "请先执行提取操作或选择输入Excel文件")
            return
            
        if not output_file:
            messagebox.showerror("错误", "请选择分析表格输出文件路径")
            return
            
        if not project_object:
            messagebox.showerror("错误", "请输入编制对象")
            return
            
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "输入Excel文件不存在")
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
            
            # 执行分析表格生成功能
            result = generate_analysis_table(input_file, output_file, project_object)
            
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
            messagebox.showinfo("完成", "分析表格生成完成！")
            
        except Exception as e:
            # 恢复原始stdout
            sys.stdout = original_stdout
            
            error_msg = f"生成分析表格时发生错误: {str(e)}"
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