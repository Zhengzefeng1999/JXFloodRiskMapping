import pandas as pd
import os
from pathlib import Path

def split_csv_by_station(input_csv_path, station_column_name, output_folder):
    """
    按测站名称将大型CSV文件分割为多个XLS文件
    
    参数:
    input_csv_path: 输入的CSV文件路径
    station_column_name: 包含测站名称的列名
    output_folder: 输出文件夹路径
    """
    
    # 创建输出文件夹（如果不存在）
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    print("正在读取CSV文件...")
    # 读取CSV文件
    df = pd.read_csv(input_csv_path)
    
    print(f"CSV文件包含 {len(df)} 行数据")
    print(f"共有 {df[station_column_name].nunique()} 个不同的测站")
    
    # 获取所有唯一的测站名称
    unique_stations = df[station_column_name].unique()
    
    print("开始分割数据...")
    
    # 为每个测站创建单独的文件
    for i, station_name in enumerate(unique_stations):
        print(f"处理第 {i+1}/{len(unique_stations)} 个测站: {station_name}")
        
        # 过滤当前测站的数据
        station_data = df[df[station_column_name] == station_name]
        
        # 清理文件名中的特殊字符，确保可以作为文件名使用
        clean_station_name = str(station_name).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
        
        # 定义输出文件路径
        output_file_path = os.path.join(output_folder, f"{clean_station_name}.xlsx")
        
        # 将数据写入XLSX文件
        station_data.to_excel(output_file_path, index=False, engine='openpyxl')
    
    print(f"完成！所有文件已保存到 {output_folder} 文件夹中")

def main():
    # 请根据你的实际情况修改以下参数
    input_csv_path = input("请输入CSV文件路径: ").strip().strip('"')
    station_column_name = input("请输入测站名称列的列名: ").strip()
    output_folder = input("请输入输出文件夹路径: ").strip().strip('"')
    
    # 如果输出文件夹为空，使用默认值
    if not output_folder:
        output_folder = "split_files"
    
    try:
        split_csv_by_station(input_csv_path, station_column_name, output_folder)
    except FileNotFoundError:
        print(f"错误：找不到文件 {input_csv_path}")
    except KeyError:
        print(f"错误：列 '{station_column_name}' 不存在于CSV文件中")
        print("CSV文件的列名如下：")
        # 临时读取列名以便显示
        temp_df = pd.read_csv(input_csv_path, nrows=0)
        print(list(temp_df.columns))
    except Exception as e:
        print(f"发生错误: {str(e)}")

if __name__ == "__main__":
    main()
