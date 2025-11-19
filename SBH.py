import os
import geopandas as gpd
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def extract_shp_to_excel(main_folder, output_file):
    # 创建Excel写入对象
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
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
            # 读取shp文件并转换为DataFrame
            gdf = gpd.read_file(shp_path)
            
            # 移除几何列（如果需要保留可以注释这行）
            df = pd.DataFrame(gdf.drop(columns='geometry'))
            
            # 将数据写入Excel的sheet
            df.to_excel(writer, sheet_name=folder_name, index=False)
            
            print(f"成功处理文件夹：{folder_name}")
            
        except Exception as e:
            print(f"处理 {folder_name} 时发生错误：{str(e)}")
            continue

    # 保存并关闭Excel文件
    writer.close()
    print(f"\n处理完成！输出文件已保存至：{output_file}")

# 使用示例
if __name__ == "__main__":
    main_folder = r"D:\Your\Main\Folder\Path"  # 修改为你的主文件夹路径
    output_file = r"D:\Output\Combined_Data.xlsx"  # 修改为输出路径
    
    extract_shp_to_excel(main_folder, output_file)