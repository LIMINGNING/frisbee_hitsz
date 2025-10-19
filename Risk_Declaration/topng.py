import fitz  # PyMuPDF
from PIL import Image
import os

def pdf_to_png(pdf_path, output_path=None, dpi=300):
    """
    将PDF文件转换为PNG图像
    
    Args:
        pdf_path (str): PDF文件路径
        output_path (str): 输出PNG文件路径，默认为None（自动生成）
        dpi (int): 分辨率，默认300
    """
    try:
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        
        # 获取第一页
        page = pdf_document[0]
        
        # 设置缩放比例（dpi/72为标准转换）
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        
        # 渲染页面为图像
        pix = page.get_pixmap(matrix=mat)
        
        # 如果没有指定输出路径，自动生成
        if output_path is None:
            base_name = os.path.splitext(pdf_path)[0]
            output_path = f"{base_name}.png"
        
        # 保存为PNG
        pix.save(output_path)
        
        # 关闭PDF文档
        pdf_document.close()
        
        print(f"✓ 转换成功: {pdf_path} -> {output_path}")
        return output_path
        
    except Exception as e:
        print(f"✗ 转换失败: {str(e)}")
        return None

def pdf_to_transparent_png(pdf_path, output_path=None, dpi=300, opacity=0.1):
    """
    将PDF转换为带透明度的PNG
    
    Args:
        pdf_path (str): PDF文件路径
        output_path (str): 输出PNG文件路径
        dpi (int): 分辨率
        opacity (float): 透明度 (0.0-1.0)
    """
    try:
        # 先转换为普通PNG
        base_name = os.path.splitext(pdf_path)[0]
        temp_png = f"{base_name}.png"  # 明确指定中间文件名
        
        temp_png = pdf_to_png(pdf_path, temp_png, dpi)
        if temp_png is None:
            return None
        
        # 使用PIL调整透明度
        img = Image.open(temp_png).convert("RGBA")
        
        # 调整透明度
        alpha = int(255 * opacity)
        img.putalpha(alpha)
        
        # 如果没有指定输出路径，自动生成
        if output_path is None:
            output_path = f"{base_name}_transparent.png"
        
        # 保存透明PNG
        img.save(output_path, "PNG")
        
        # 不删除中间文件，保留普通PNG
        print(f"✓ 透明PNG创建成功: {output_path}")
        print(f"✓ 普通PNG已保留: {temp_png}")
        return output_path
        
    except Exception as e:
        print(f"✗ 透明PNG创建失败: {str(e)}")
        return None

if __name__ == "__main__":
    # 转换协会徽章
    pdf_file = "协会徽章.pdf"
    
    if os.path.exists(pdf_file):
        print("开始转换...")
        
        # 创建带透明度的PNG（会自动生成普通PNG作为中间步骤）
        transparent_png = pdf_to_transparent_png(pdf_file, "协会徽章_淡.png", dpi=300, opacity=0.1)
        
        print("\n转换完成！生成的文件：")
        print("1. 协会徽章.png - 普通PNG文件")
        print("2. 协会徽章_淡.png - 透明PNG文件")
        print("\n在LaTeX中可以使用任一文件：")
        print("- 直接使用: 协会徽章_淡.png")
        print("- 或配合tikz使用: 协会徽章.png")
        
    else:
        print(f"错误: 找不到文件 {pdf_file}")
        print("请确保PDF文件在当前目录中。")