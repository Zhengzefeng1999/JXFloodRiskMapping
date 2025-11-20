# SHP to Excel Converter

一个用于批量提取Shapefile文件数据并导出到Excel表格的Python工具。

## 功能特点

- 批量处理多个文件夹中的SHP文件
- 自动遍历子文件夹并提取"2D Zones.shp"文件
- 将每个文件夹的数据导出为Excel中的独立工作表
- 自动处理异常情况并给出友好提示
- 移除几何列，仅保留属性数据

## 依赖环境

- Python 3.6+
- pandas
- geopandas
- openpyxl

## 安装依赖

```bash
pip install pandas geopandas openpyxl
