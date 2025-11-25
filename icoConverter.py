from PIL import Image

def convert_to_ico(input_path, output_path):
    img = Image.open(input_path)
    img.save(output_path, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])

# 使用示例
convert_to_ico('JXFloodRiskMapping.png', 'JXFloodRiskMapping.ico')