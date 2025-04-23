import os
from markdown import markdown
from docx import Document

def convert_md_to_docx(md_file, output_file):
    """
    将 Markdown 文件转换为 Word 文件
    :param md_file: Markdown 文件路径
    :param output_file: 输出的 Word 文件路径
    """
    with open(md_file, 'r', encoding='utf-8') as file:
        md_content = file.read()
    
    # 将 Markdown 转换为 HTML
    html_content = markdown(md_content)
    
    # 创建 Word 文档
    doc = Document()
    doc.add_paragraph(html_content)
    doc.save(output_file)

def convert_all_md_in_folder(folder_path):
    """
    将文件夹中的所有 Markdown 文件转换为 Word 文件
    :param folder_path: 文件夹路径
    """
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.md'):
            md_file_path = os.path.join(folder_path, file_name)
            output_file_path = os.path.join(folder_path, file_name.replace('.md', '.docx'))
            convert_md_to_docx(md_file_path, output_file_path)
            print(f"已转换: {md_file_path} -> {output_file_path}")

if __name__ == "__main__":
    # 获取当前脚本所在的文件夹路径
    current_folder = os.path.dirname(os.path.abspath(__file__))
    convert_all_md_in_folder(current_folder)