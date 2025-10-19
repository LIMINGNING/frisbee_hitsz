import fitz  # PyMuPDF
import os
from pathlib import Path
from PIL import Image

def extract_last_pages_to_png(pdf_path, output_dir=None):
    """
    提取PDF的最后两页并转换为PNG图片
    
    Args:
        pdf_path (str): PDF文件路径
        output_dir (str, optional): 输出目录，默认为PDF文件所在目录
    """
    # 确保PDF文件存在
    if not os.path.exists(pdf_path):
        print(f"错误: 找不到文件 {pdf_path}")
        return None, None
    
    # 设置输出目录
    if output_dir is None:
        output_dir = os.path.dirname(pdf_path)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        total_pages = len(pdf_document)
        
        if total_pages < 2:
            print(f"警告: PDF只有 {total_pages} 页，无法提取最后两页")
            if total_pages == 1:
                print("将提取唯一的一页")
                pages_to_extract = [0]
            else:
                print("PDF为空")
                return None, None
        else:
            # 获取最后两页的索引 (从0开始)
            pages_to_extract = [total_pages - 2, total_pages - 1]
        
        # 获取PDF文件名（不含扩展名）
        pdf_name = Path(pdf_path).stem
        
        saved_files = []
        
        # 提取并保存页面
        for i, page_num in enumerate(pages_to_extract):
            # 获取页面
            page = pdf_document[page_num]
            
            # 设置渲染参数 (提高图片质量)
            mat = fitz.Matrix(2.0, 2.0)  # 放大2倍，提高清晰度
            pix = page.get_pixmap(matrix=mat)
            
            # 生成输出文件名
            if len(pages_to_extract) == 2:
                page_label = f"倒数第{2-i}页"
                output_filename = f"{pdf_name}_倒数第{2-i}页.png"
            else:
                page_label = "第1页"
                output_filename = f"{pdf_name}_第1页.png"
            
            output_path = os.path.join(output_dir, output_filename)
            
            # 保存为PNG
            pix.save(output_path)
            saved_files.append(output_path)
            print(f"已保存 {page_label} (页面 {page_num + 1}) 到: {output_path}")
        
        # 关闭PDF文档
        pdf_document.close()
        print(f"完成! 共提取 {len(pages_to_extract)} 页")
        
        # 返回保存的文件路径
        if len(saved_files) == 2:
            return saved_files[0], saved_files[1]  # 倒数第2页, 倒数第1页
        elif len(saved_files) == 1:
            return saved_files[0], None
        else:
            return None, None
        
    except Exception as e:
        print(f"错误: {str(e)}")
        return None, None

def merge_images(image1_path, image2_path, output_path, direction='horizontal'):
    """
    合并两个PNG图片
    
    Args:
        image1_path (str): 第一个图片路径
        image2_path (str): 第二个图片路径  
        output_path (str): 输出合并图片路径
        direction (str): 合并方向，'horizontal'(水平) 或 'vertical'(垂直)
    """
    try:
        # 打开两个图片
        img1 = Image.open(image1_path)
        img2 = Image.open(image2_path)
        
        if direction == 'horizontal':
            # 水平合并 (左右拼接)
            # 计算合并后的尺寸
            total_width = img1.width + img2.width
            max_height = max(img1.height, img2.height)
            
            # 创建新的图片
            merged_img = Image.new('RGB', (total_width, max_height), (255, 255, 255))
            
            # 粘贴图片
            merged_img.paste(img1, (0, 0))
            merged_img.paste(img2, (img1.width, 0))
            
        else:  # vertical
            # 垂直合并 (上下拼接)
            # 计算合并后的尺寸
            max_width = max(img1.width, img2.width)
            total_height = img1.height + img2.height
            
            # 创建新的图片
            merged_img = Image.new('RGB', (max_width, total_height), (255, 255, 255))
            
            # 粘贴图片 (居中对齐)
            x1 = (max_width - img1.width) // 2
            x2 = (max_width - img2.width) // 2
            merged_img.paste(img1, (x1, 0))
            merged_img.paste(img2, (x2, img1.height))
        
        # 保存合并后的图片
        merged_img.save(output_path)
        print(f"已保存合并图片到: {output_path}")
        return True
        
    except Exception as e:
        print(f"合并图片时出错: {str(e)}")
        return False

def main():
    # PDF文件路径
    pdf_file = "2025新生杯.pdf"
    
    # 检查文件是否存在
    if not os.path.exists(pdf_file):
        print(f"找不到文件: {pdf_file}")
        print("请确保PDF文件存在于正确的路径")
        return
    
    # 设置输出目录为当前目录
    output_directory = "."
    
    print(f"开始提取 {pdf_file} 的最后两页...")
    page1_path, page2_path = extract_last_pages_to_png(pdf_file, output_directory)
    
    # 如果成功提取了两页，则合并它们
    if page1_path and page2_path:
        print("\n开始合并图片...")
        pdf_name = Path(pdf_file).stem
        
        # 生成合并后的文件名
        merged_horizontal_path = f"{pdf_name}_合并_水平.png"
        merged_vertical_path = f"{pdf_name}_合并_垂直.png"
        
        # 水平合并 (左右拼接)
        print("创建水平合并版本...")
        merge_images(page1_path, page2_path, merged_horizontal_path, 'horizontal')
        
        # 垂直合并 (上下拼接)  
        print("创建垂直合并版本...")
        merge_images(page1_path, page2_path, merged_vertical_path, 'vertical')
        
        print("\n合并完成! 生成了以下文件:")
        print(f"- 单独页面: {page1_path}, {page2_path}")
        print(f"- 水平合并: {merged_horizontal_path}")
        print(f"- 垂直合并: {merged_vertical_path}")
        
    elif page1_path:
        print("只有一页，无需合并")
    else:
        print("没有成功提取页面")

if __name__ == "__main__":
    main()