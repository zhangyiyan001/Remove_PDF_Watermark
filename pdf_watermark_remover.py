#!/usr/bin/env python
# -*- coding: utf-8 -*-

import fitz  # PyMuPDF
import os
import re
import glob
from PIL import Image
import io
import sys

def check_dependencies():
    """检查并提示安装所需的依赖库"""
    missing_packages = []
    
    try:
        import fitz
    except ImportError:
        missing_packages.append("PyMuPDF")
    
    try:
        from PIL import Image
    except ImportError:
        missing_packages.append("pillow")
    
    if missing_packages:
        print("错误: 缺少必要的库。请运行以下命令安装:")
        for package in missing_packages:
            print(f"pip install {package.lower()}")
        sys.exit(1)

def find_pdf_files():
    """查找当前目录下的所有PDF文件"""
    pdf_files = glob.glob("*.pdf")
    return pdf_files

def pdf_to_png(pdf_path, output_folder):
    """
    将PDF文件转换为PNG图片
    
    参数:
        pdf_path: PDF文件路径
        output_folder: 输出图片的文件夹路径
    """
    print("\n=== 第1步: 将PDF转换为PNG图片 ===")
    
    # 检查PDF文件是否存在
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件 '{pdf_path}' 不存在")
        return False
    
    try:
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        total_pages = pdf_document.page_count
        
        # 创建输出文件夹
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"创建输出文件夹: {output_folder}")
        
        # 遍历PDF中的每一页
        for page_num in range(total_pages):
            # 获取当前页面
            page = pdf_document.load_page(page_num)
            
            # 将页面转换为pixmap（图像对象）
            pix = page.get_pixmap(dpi=600)  # 可调整dpi参数以控制输出图像的分辨率
            
            # 输出图片的文件路径
            output_path = os.path.join(output_folder, f"page_{page_num + 1}.png")
            
            # 将pixmap保存为PNG图片
            pix.save(output_path)
            print(f"页面 {page_num + 1}/{total_pages} 已保存为 {output_path}")
        
        print(f"PDF转换完成! 共转换了 {total_pages} 页")
        
        # 关闭PDF文件
        pdf_document.close()
        return True
        
    except Exception as e:
        print(f"转换PDF时出错: {e}")
        return False

def replace_rgb_with_white(image_folder):
    """
    处理图片，将特定RGB范围内的像素替换为白色（去除水印）
    
    参数:
        image_folder: 包含图片的文件夹路径
    """
    print("\n=== 第2步: 处理图片去除水印 ===")
    
    processed_count = 0
    # 遍历文件夹中的所有PNG图片
    for filename in os.listdir(image_folder):
        if filename.endswith(".png") and not filename.endswith("_modify.png"):
            image_path = os.path.join(image_folder, filename)
            
            # 打开图片
            img = Image.open(image_path)
            img = img.convert("RGB")  # 转换为RGB模式（如果不是的话）
            
            # 获取图片的像素数据
            pixels = img.load()
            
            # 遍历所有像素，检查RGB是否在(216, 216, 216) ± 30范围内
            for i in range(img.width):
                for j in range(img.height):
                    r, g, b = pixels[i, j]
                    
                    # 判断像素是否在(216, 216, 216) ± 30范围内
                    if (186 <= r <= 246) and (186 <= g <= 246) and (186 <= b <= 246):
                        # 将符合条件的像素替换为白色
                        pixels[i, j] = (255, 255, 255)

            # 生成新的文件名，添加"_modify"后缀
            new_filename = os.path.splitext(filename)[0] + "_modify.png"
            new_image_path = os.path.join(image_folder, new_filename)
            
            # 保存修改后的图片
            img.save(new_image_path)
            processed_count += 1
            print(f"处理并保存: {new_image_path}")
    
    print(f"图片处理完成! 共处理了 {processed_count} 张图片")
    return processed_count > 0

def clean_images(output_dir):
    """
    删除output_images文件夹中所有不以modify.png结尾的PNG图片
    
    参数:
        output_dir: 包含图片的文件夹路径
    """
    print("\n=== 第3步: 清理未处理的图片 ===")
    
    # 确保output_images文件夹存在
    if not os.path.exists(output_dir):
        print(f"错误: {output_dir}文件夹不存在")
        return False
    
    # 获取所有PNG图片
    all_png_files = glob.glob(os.path.join(output_dir, '*.png'))
    
    # 筛选出不以modify.png结尾的图片
    files_to_delete = [f for f in all_png_files if not f.endswith('modify.png')]
    
    # 删除这些图片
    deleted_count = 0
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
            print(f"已删除: {file_path}")
            deleted_count += 1
        except Exception as e:
            print(f"删除{file_path}时出错: {e}")
    
    print(f"清理完成! 共删除了 {deleted_count} 个文件。")
    print(f"保留了所有以'modify.png'结尾的图片。")
    return True

def create_pdf_from_images(output_dir, pdf_filename):
    """
    将output_images文件夹中所有以modify.png结尾的图片按照文件名中的数字顺序合并成一个PDF文件
    
    参数:
        output_dir: 包含图片的文件夹路径
        pdf_filename: 输出PDF文件名
    """
    print("\n=== 第4步: 将处理后的图片合并为PDF ===")
    
    # 确保output_images文件夹存在
    if not os.path.exists(output_dir):
        print(f"错误: {output_dir}文件夹不存在")
        return False
    
    # 获取所有以modify.png结尾的图片
    image_files = glob.glob(os.path.join(output_dir, '*modify.png'))
    
    if not image_files:
        print(f"错误: 在{output_dir}文件夹中没有找到以modify.png结尾的图片")
        return False
    
    # 使用正则表达式从文件名中提取数字
    def extract_number(filename):
        match = re.search(r'page_(\d+)_modify\.png', os.path.basename(filename))
        if match:
            return int(match.group(1))
        return 0
    
    # 按照提取的数字对文件进行排序
    sorted_image_files = sorted(image_files, key=extract_number)
    
    print(f"找到了 {len(sorted_image_files)} 个图片文件，准备合并为PDF...")
    
    # 打开所有图片
    images = []
    for img_path in sorted_image_files:
        try:
            img = Image.open(img_path).convert('RGB')
            images.append(img)
            print(f"已加载: {img_path}")
        except Exception as e:
            print(f"加载图片 {img_path} 时出错: {e}")
    
    if not images:
        print("错误: 没有成功加载任何图片")
        return False
    
    # 保存为PDF
    try:
        # 第一张图片保存为PDF，后续图片追加
        images[0].save(
            pdf_filename, 
            save_all=True, 
            append_images=images[1:] if len(images) > 1 else []
        )
        print(f"\nPDF创建成功! 文件保存为: {pdf_filename}")
        print(f"共合并了 {len(images)} 张图片。")
        return True
    except Exception as e:
        print(f"创建PDF时出错: {e}")
        return False

def main():
    """主函数，执行完整的处理流程"""
    print("=== PDF水印去除工具 ===")
    
    # 检查依赖
    check_dependencies()
    
    # 查找当前目录下的PDF文件
    pdf_files = find_pdf_files()
    
    if not pdf_files:
        print("错误: 当前目录下没有找到PDF文件")
        return
    
    # 如果找到多个PDF文件，让用户选择
    if len(pdf_files) > 1:
        print("\n在当前目录下找到以下PDF文件:")
        for i, pdf_file in enumerate(pdf_files):
            print(f"{i+1}. {pdf_file}")
        
        try:
            choice = int(input("\n请选择要处理的PDF文件编号 (1-{0}): ".format(len(pdf_files))))
            if choice < 1 or choice > len(pdf_files):
                print("无效的选择，将使用第一个PDF文件")
                pdf_path = pdf_files[0]
            else:
                pdf_path = pdf_files[choice-1]
        except ValueError:
            print("无效的输入，将使用第一个PDF文件")
            pdf_path = pdf_files[0]
    else:
        pdf_path = pdf_files[0]
    
    print(f"\n选择的PDF文件: {pdf_path}")
    
    # 设置默认输出文件夹和PDF文件名
    output_folder = 'output_images'
    output_pdf = os.path.splitext(pdf_path)[0] + "_无水印.pdf"
    
    print(f"输出文件夹: {output_folder}")
    print(f"输出PDF文件: {output_pdf}")
    
    # 执行处理流程
    if pdf_to_png(pdf_path, output_folder):
        if replace_rgb_with_white(output_folder):
            if clean_images(output_folder):
                if create_pdf_from_images(output_folder, output_pdf):
                    print("\n=== 处理完成! ===")
                    print(f"已成功去除水印并创建新的PDF文件: {output_pdf}")
                    return
    
    print("\n处理过程中出现错误，请检查上述错误信息。")

if __name__ == "__main__":
    main() 