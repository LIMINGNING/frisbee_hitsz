import pandas as pd
import os
import glob

def xlsx_to_csv(input_folder='.', output_folder=None):
    """
    将指定文件夹中的所有 .xlsx 文件转换为 .csv 文件
    
    Args:
        input_folder (str): 输入文件夹路径，默认为当前目录
        output_folder (str): 输出文件夹路径，默认为输入文件夹
    """
    if output_folder is None:
        output_folder = input_folder
    
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 查找所有 .xlsx 文件
    xlsx_files = glob.glob(os.path.join(input_folder, '*.xlsx'))
    
    if not xlsx_files:
        print("未找到任何 .xlsx 文件")
        return
    
    print(f"找到 {len(xlsx_files)} 个 .xlsx 文件")
    
    for xlsx_file in xlsx_files:
        try:
            # 获取文件名（不含扩展名）
            base_name = os.path.splitext(os.path.basename(xlsx_file))[0]
            csv_file = os.path.join(output_folder, f"{base_name}.csv")
            
            # 读取 Excel 文件
            df = pd.read_excel(xlsx_file)
            
            # 保存为 CSV 文件
            df.to_csv(csv_file, index=False, encoding='utf-8')
            
            print(f"✓ {xlsx_file} -> {csv_file}")
            
        except Exception as e:
            print(f"✗ 转换失败 {xlsx_file}: {str(e)}")

def xlsx_to_csv_multi_sheet(input_folder='.', output_folder=None):
    """
    将包含多个工作表的 .xlsx 文件转换为多个 .csv 文件
    
    Args:
        input_folder (str): 输入文件夹路径，默认为当前目录
        output_folder (str): 输出文件夹路径，默认为输入文件夹
    """
    if output_folder is None:
        output_folder = input_folder
    
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 查找所有 .xlsx 文件
    xlsx_files = glob.glob(os.path.join(input_folder, '*.xlsx'))
    
    if not xlsx_files:
        print("未找到任何 .xlsx 文件")
        return
    
    print(f"找到 {len(xlsx_files)} 个 .xlsx 文件")
    
    for xlsx_file in xlsx_files:
        try:
            # 获取文件名（不含扩展名）
            base_name = os.path.splitext(os.path.basename(xlsx_file))[0]
            
            # 读取所有工作表
            excel_file = pd.ExcelFile(xlsx_file)
            
            for sheet_name in excel_file.sheet_names:
                # 读取特定工作表
                df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
                
                # 生成 CSV 文件名
                if len(excel_file.sheet_names) == 1:
                    csv_file = os.path.join(output_folder, f"{base_name}.csv")
                else:
                    csv_file = os.path.join(output_folder, f"{base_name}_{sheet_name}.csv")
                
                # 保存为 CSV 文件
                df.to_csv(csv_file, index=False, encoding='utf-8')
                
                print(f"✓ {xlsx_file} (工作表: {sheet_name}) -> {csv_file}")
            
        except Exception as e:
            print(f"✗ 转换失败 {xlsx_file}: {str(e)}")

if __name__ == "__main__":
    # 转换当前目录下的所有 .xlsx 文件
    print("开始转换 .xlsx 文件为 .csv 文件...")
    
    # 选择转换方式：
    # 1. 只转换第一个工作表
    xlsx_to_csv()
    
    # 2. 转换所有工作表（如果需要处理多工作表文件，取消下面的注释）
    # xlsx_to_csv_multi_sheet()
    
    print("转换完成！")