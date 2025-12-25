import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import geopandas as gpd
import pandas as pd
from shapely.geometry import MultiPoint
from shapely.ops import voronoi_diagram
import os

class ThiessenPolygonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("泰森多边形生成工具")
        self.root.geometry("600x600")
        
        # 变量存储文件路径和数据
        self.polygon_shp_path = ""
        self.rain_gauge_shp_path = ""
        self.rain_gauges = None
        self.output_folder = ""
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 第一步：选择面文件
        step1_frame = ttk.LabelFrame(main_frame, text="第一步：选择面文件", padding="10")
        step1_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.polygon_label = ttk.Label(step1_frame, text="未选择面文件")
        self.polygon_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.select_polygon_btn = ttk.Button(step1_frame, text="选择面文件", command=self.select_polygon_file)
        self.select_polygon_btn.grid(row=0, column=1, padx=(10, 0))
        
        # 第二步：选择雨量站点文件
        step2_frame = ttk.LabelFrame(main_frame, text="第二步：选择雨量站点文件", padding="10")
        step2_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.rain_gauge_label = ttk.Label(step2_frame, text="未选择雨量站点文件")
        self.rain_gauge_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.select_rain_gauge_btn = ttk.Button(step2_frame, text="选择雨量站点文件", command=self.select_rain_gauge_file)
        self.select_rain_gauge_btn.grid(row=0, column=1, padx=(10, 0))
        
        # 第三步：输出选项
        step3_frame = ttk.LabelFrame(main_frame, text="第三步：输出选项", padding="10")
        step3_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.output_option_var = tk.BooleanVar()
        self.output_option_check = ttk.Checkbutton(step3_frame, text="需要输出泰森多边形信息表", 
                                                   variable=self.output_option_var, 
                                                   command=self.toggle_output_options)
        self.output_option_check.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        # 名称字段选择部分（初始隐藏）
        self.name_field_frame = ttk.Frame(step3_frame)
        self.name_field_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        self.name_field_frame.grid_remove()  # 初始隐藏
        
        ttk.Label(self.name_field_frame, text="选择站点名称字段:").grid(row=0, column=0, sticky=tk.W)
        
        self.field_combo = ttk.Combobox(self.name_field_frame, state="readonly", width=30)
        self.field_combo.grid(row=0, column=1, padx=(10, 0))
        
        # 第四步：输出路径
        step4_frame = ttk.LabelFrame(main_frame, text="第四步：选择输出路径", padding="10")
        step4_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.output_label = ttk.Label(step4_frame, text="未选择输出路径")
        self.output_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.select_output_btn = ttk.Button(step4_frame, text="选择输出路径", command=self.select_output_path)
        self.select_output_btn.grid(row=0, column=1, padx=(10, 0))
        
        # 执行按钮
        self.execute_btn = ttk.Button(main_frame, text="执行泰森多边形生成", command=self.execute_thiessen)
        self.execute_btn.grid(row=4, column=0, columnspan=2, pady=20)
        self.execute_btn.config(state='disabled')
        
        # 进度和结果显示
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.progress_var = tk.StringVar(value="就绪")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
    def select_polygon_file(self):
        file_path = filedialog.askopenfilename(
            title="选择面文件",
            filetypes=[("Shapefile", "*.shp"), ("All files", "*.*")]
        )
        if file_path:
            self.polygon_shp_path = file_path
            self.polygon_label.config(text=os.path.basename(file_path))
            self.check_ready()
    
    def select_rain_gauge_file(self):
        file_path = filedialog.askopenfilename(
            title="选择雨量站点文件",
            filetypes=[("Shapefile", "*.shp"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.rain_gauges = gpd.read_file(file_path)
                self.rain_gauge_shp_path = file_path
                self.rain_gauge_label.config(text=os.path.basename(file_path))
                
                # 更新字段下拉框
                self.update_field_combo()
                
                self.check_ready()
            except Exception as e:
                messagebox.showerror("错误", f"读取雨量站点文件失败: {str(e)}")
    
    def update_field_combo(self):
        if self.rain_gauges is not None:
            fields = self.rain_gauges.columns.tolist()
            self.field_combo['values'] = fields
            if fields:
                # 尝试默认选择常见的名称字段
                common_name_fields = ['Name', 'name', 'NAME', 'Station', 'station', 'STATION', 'ID', 'id', 'ID_']
                selected_field = None
                for field in common_name_fields:
                    if field in fields:
                        selected_field = field
                        break
                if selected_field is None:
                    selected_field = fields[0]  # 默认选择第一个字段
                self.field_combo.set(selected_field)
    
    def toggle_output_options(self):
        if self.output_option_var.get():
            self.name_field_frame.grid()
        else:
            self.name_field_frame.grid_remove()
        self.check_ready()
    
    def select_output_path(self):
        folder_path = filedialog.askdirectory(title="选择输出文件夹")
        if folder_path:
            self.output_folder = folder_path
            self.output_label.config(text=folder_path)
            self.check_ready()
    
    def check_ready(self):
        # 检查是否所有必需的文件都已选择
        ready = (self.polygon_shp_path and 
                 self.rain_gauge_shp_path and 
                 self.output_folder)
        
        if ready and self.output_option_var.get():
            # 如果需要输出信息表，确保已选择名称字段
            ready = self.field_combo.get() != ""
        
        self.execute_btn.config(state='normal' if ready else 'disabled')
    
    def execute_thiessen(self):
        if not self.polygon_shp_path or not self.rain_gauge_shp_path or not self.output_folder:
            messagebox.showerror("错误", "请完成所有步骤的选择")
            return
        
        try:
            self.progress_var.set("正在处理...")
            self.root.update()
            
            # 读取数据
            polygons = gpd.read_file(self.polygon_shp_path)
            rain_gauges = gpd.read_file(self.rain_gauge_shp_path)
            
            # 检查坐标系统
            if polygons.crs != rain_gauges.crs:
                rain_gauges = rain_gauges.to_crs(polygons.crs)
            
            # 创建输出文件夹
            if not os.path.exists(self.output_folder):
                os.makedirs(self.output_folder)
            
            # 初始化结果DataFrame
            results = pd.DataFrame(columns=["Polygon_Name", "Shape_Name", "Area_km2"])
            
            # 初始化坐标结果DataFrame（如果需要输出）
            coordinates_results = None
            if self.output_option_var.get():
                coordinates_results = pd.DataFrame(columns=["Polygon_Name", "Station_Name", "Station_ID", "Region_and_Coordinates"])
            
            # 遍历面文件中的每个shape
            for idx, polygon in polygons.iterrows():
                self.progress_var.set(f"正在处理多边形 {idx + 1}/{len(polygons)}")
                self.root.update()
                
                # 提取当前shape
                current_polygon = polygon.geometry
                polygon_name = f"polygon_{idx + 1}"
                
                # 导出当前shape为单独的Shapefile
                single_polygon_gdf = gpd.GeoDataFrame([polygon], columns=polygons.columns, crs=polygons.crs)
                # 删除可能的多余字段
                
                single_polygon_path = os.path.join(self.output_folder, f"{polygon_name}.shp")
                single_polygon_gdf.to_file(single_polygon_path, encoding='utf-8')
                
                # 使用全部雨量站生成泰森多边形
                points = MultiPoint(rain_gauges.geometry)
                voronoi_polygons = voronoi_diagram(points)
                
                # 将泰森多边形转换为GeoDataFrame，并保留雨量站的属性
                voronoi_gdf = gpd.GeoDataFrame(geometry=[poly for poly in voronoi_polygons.geoms], crs=polygons.crs)
                voronoi_gdf = voronoi_gdf.sjoin(rain_gauges, how="inner")  # 保留雨量站的属性字段
                
                # 将泰森多边形与当前shape相交
                intersected_polygons = gpd.overlay(voronoi_gdf, single_polygon_gdf, how='intersection')
                
                # 计算相交后的面积并筛选有效多边形
                intersected_polygons['area_km2'] = intersected_polygons.geometry.area / 1e6  # 面积单位：平方千米
                valid_polygons = intersected_polygons[intersected_polygons['area_km2'] > 0]
                
                # 输出相交后的Shapefile
                output_shp_name = f"{polygon_name}_intersected.shp"
                output_shp_path = os.path.join(self.output_folder, output_shp_name)
                valid_polygons.to_file(output_shp_path, encoding='utf-8')
                
                # 将结果添加到结果表中
                for _, row in valid_polygons.iterrows():
                    new_row = pd.DataFrame({
                        "Polygon_Name": [polygon_name],
                        "Shape_Name": [output_shp_name],
                        "ID": [row['ID']] if 'ID' in row else [None],  # 添加ID列
                        "Area_km2": [row['area_km2']]
                    })
                    results = pd.concat([results, new_row], ignore_index=True)
                    
                    # 如果需要输出坐标信息
                    if self.output_option_var.get():
                        # 提取顶点坐标
                        geometry = row.geometry
                        vertex_coords = []
                        
                        if geometry.geom_type == 'Polygon':
                            # 对于单个多边形，获取外边界坐标
                            exterior_coords = list(geometry.exterior.coords)
                            # 计算节点个数（减去重复的第一个点）
                            num_vertices = len(exterior_coords)
                            # 格式化为 "X1,Y1,X2,Y2,X3,Y3..." 的形式
                            coord_str = ",".join([f"{x},{y}" for x, y in exterior_coords])  
                            vertex_coords.append((num_vertices, coord_str))
                        elif geometry.geom_type == 'MultiPolygon':
                            # 对于多个多边形，分别处理每个部分
                            for poly in geometry.geoms:
                                exterior_coords = list(poly.exterior.coords)
                                # 计算节点个数（减去重复的第一个点）
                                num_vertices = len(exterior_coords)
                                coord_str = ",".join([f"{x},{y}" for x, y in exterior_coords]) 
                                vertex_coords.append((num_vertices, coord_str))
                        
                        # 添加到坐标结果表（合并REGION和坐标）
                        for num_vertices, coord_str in vertex_coords:
                            station_name = row.get(self.field_combo.get(), 'Unknown')
                            # 合并REGION和坐标信息
                            combined_info = f"REGION={num_vertices},{coord_str}"
                            coord_row = pd.DataFrame({
                                "Polygon_Name": [polygon_name],
                                "Station_Name": [station_name],
                                "Station_ID": [row.get('ID', 'Unknown')],
                                "Region": [f"REGION={num_vertices}"],
                                "Vertex_Coordinates": [coord_str],
                                "Region_and_Coordinates": [combined_info]
                            })
                            coordinates_results = pd.concat([coordinates_results, coord_row], ignore_index=True)

            # 输出结果到Excel
            results_path = os.path.join(self.output_folder, "results.xlsx")
            results.to_excel(results_path, index=False)
            
            # 如果需要输出坐标信息，也保存到Excel
            if self.output_option_var.get() and coordinates_results is not None:
                coordinates_path = os.path.join(self.output_folder, "coordinates_results.xlsx")
                coordinates_results.to_excel(coordinates_path, index=False)
            
            self.progress_var.set("处理完成！")
            messagebox.showinfo("完成", f"泰森多边形生成完成！\n结果已保存到: {self.output_folder}")
            
        except Exception as e:
            self.progress_var.set("处理失败")
            messagebox.showerror("错误", f"处理过程中出现错误: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ThiessenPolygonApp(root)
    root.mainloop()
