import os
import pypandoc

def convert_md_to_docx_with_pypandoc(md_file, output_file):
    """
    使用 pypandoc 将 Markdown 文件转换为 Word 文件
    :param md_file: Markdown 文件路径
    :param output_file: 输出的 Word 文件路径
    """
    try:
        pypandoc.convert_file(md_file, 'docx', outputfile=output_file)
        print(f"已成功转换: {md_file} -> {output_file}")
    except Exception as e:
        print(f"转换失败: {md_file} -> {output_file}, 错误: {e}")

def convert_all_md_in_folder_with_pypandoc(folder_path):
    """
    将文件夹中的所有 Markdown 文件转换为 Word 文件
    :param folder_path: 文件夹路径
    """
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.md'):
            md_file_path = os.path.join(folder_path, file_name)
            output_file_path = os.path.join(folder_path, file_name.replace('.md', '.docx'))
            convert_md_to_docx_with_pypandoc(md_file_path, output_file_path)

if __name__ == "__main__":
    # 动态输入文件夹路径
    folder_path = input("请输入要转换的文件夹路径: ").strip()
    if os.path.isdir(folder_path):
        convert_all_md_in_folder_with_pypandoc(folder_path)
    else:
        print("无效的文件夹路径，请检查后重试。")