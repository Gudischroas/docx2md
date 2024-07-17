import os  # 导入操作系统相关的模块
import pypandoc  # 导入pypandoc模块，用于文件格式转换

def count_files_and_folders(folder_path, recursive):
    """
    统计指定文件夹中的文件和文件夹数量。
    :param folder_path: 文件夹路径
    :param recursive: 是否递归搜索子文件夹
    :return: 文件夹数量和文件数量
    """
    if recursive:
        # 递归搜索子文件夹
        num_folders = sum([len(dirs) for _, dirs, _ in os.walk(folder_path)])
        num_files = sum([len(files) for _, _, files in os.walk(folder_path)])
    else:
        # 仅搜索指定文件夹
        num_folders = len([f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))])
        num_files = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
    return num_folders, num_files

def convert_docx_to_md(folder_path, recursive, media_extraction):
    """
    将指定文件夹中的DOCX文件转换为Markdown文件。
    :param folder_path: 文件夹路径
    :param recursive: 是否递归搜索子文件夹
    :param media_extraction: 是否提取媒体文件
    :return: 处理进度的生成器
    """
    num_folders, num_files = count_files_and_folders(folder_path, recursive)  # 获取文件夹和文件数量
    if recursive:
        docx_files = []
        for root, dirs, files in os.walk(folder_path):
            docx_files += [os.path.join(root, f) for f in files if f.endswith('.docx')]
    else:
        docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.docx')]
    
    num_docx_files = len(docx_files)
    print(f'找到的文件夹总数: {num_folders}')
    print(f'找到的文件总数: {num_files}')
    print(f'找到的DOCX文件总数: {num_docx_files}')
    
    # 指定媒体提取文件夹
    media_folder = os.path.join(folder_path, "media")
    os.makedirs(media_folder, exist_ok = True)  # 创建媒体文件夹

    for i, filename in enumerate(docx_files):
        if os.path.isdir(filename):
            continue

        print(f"处理文件: {filename}")  # 输出文件名称
        md_path = os.path.splitext(filename)[0] + '.md'  # 生成Markdown文件路径
        if media_extraction:
            extra_args = ['--wrap=none', f'--extract-media={media_folder}']
        else:
            extra_args = ['--wrap=none']
        pypandoc.convert_file(filename, 'md', format = 'docx', outputfile = md_path, 
                              extra_args = extra_args, encoding = 'utf-8')  # 转换文件
        yield i + 1, len(docx_files)  # 生成当前进度

def main():
    """
    主函数，负责处理用户输入并调用转换函数。
    """
    folder_path = input("请输入文件夹路径: ")  # 获取用户输入的文件夹路径
    recursive = input("搜索子文件夹? (yes/no): ").strip().lower() == 'yes'  # 获取是否递归搜索子文件夹
    media_extraction = input("提取媒体文件? (yes/no): ").strip().lower() == 'yes'  # 获取是否提取媒体文件

    total_files = 0
    for i, total in convert_docx_to_md(folder_path, recursive, media_extraction):
        progress_num = int(100 * i / total)  # 计算进度百分比
        print(f"进度: {progress_num}%")  # 输出进度
        total_files = total

    print("完成!")
    print(f"已转换{total_files}个DOCX文件.")

if __name__ == "__main__":
    main()  # 调用主函数
