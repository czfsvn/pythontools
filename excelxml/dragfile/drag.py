import sys
#import pandas as pd

def parse_excel(file_path):
    """解析Excel文件"""
    try:
        #df = pd.read_excel(file_path)
        print(f"成功读取文件：{file_path}")
        print("数据预览：")
        #print(df.head())
    except Exception as e:
        print(f"解析失败：{str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 处理拖放的多个文件
        for file in sys.argv[1:]:
            parse_excel(file)
        input("Press Enter to continue...")
    else:
        print("请拖放Excel文件到程序图标上")