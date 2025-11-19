import pandas as pd
from pyproj import Transformer

def cgcs2000_to_wgs84(x, y):
    """ 使用中央子午线117°的转换函数 """
    # 定义投影参数（关键参数已修正）
    cgcs2000 = (
        "+proj=tmerc +lat_0=0 +lon_0=117 "  # 中央经度设为117°
        "+k=1 +x_0=500000 +y_0=0 "         # 东偏移500km
        "+ellps=GRS80 +units=m +no_defs"
    )
    transformer = Transformer.from_crs(cgcs2000, "EPSG:4326", always_xy=True)
    lon, lat = transformer.transform(x, y)
    return round(lon, 6), round(lat, 6)  # 保留6位小数匹配验证数据

# 验证数据测试
x_test, y_test = 358864.94, 3181680.15
expected_lon, expected_lat = 115.555190922, 28.743353954
calc_lon, calc_lat = cgcs2000_to_wgs84(x_test, y_test)

print(f"计算结果：经度={calc_lon:.9f}, 纬度={calc_lat:.9f}")
print(f"预期结果：经度={expected_lon:.9f}, 纬度={expected_lat:.9f}")
print(f"经度误差：{abs(calc_lon - expected_lon):.9f} 度")
print(f"纬度误差：{abs(calc_lat - expected_lat):.9f} 度")

# 应用到Excel的完整代码
df = pd.read_excel(r"J:\BM\16、洪水风险图\Model\石鼻河\石鼻河成果汇交.xlsx", engine='openpyxl')
df[['经度', '纬度']] = df.apply(
    lambda row: cgcs2000_to_wgs84(row['X'], row['Y']),
    axis=1,
    result_type='expand'
)
df.to_excel("幸福水库成果汇交.xlsx", index=False, float_format="%.9f", engine='openpyxl')







