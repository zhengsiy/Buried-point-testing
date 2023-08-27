import os


def get_jk_path():
    # 获取当前文件夹的绝对路径
    current_folder_path = os.path.abspath(os.path.dirname(__file__))

    # 要进入的文件夹的名字
    jk_folder_name = "jk埋点"

    # jk埋点文件夹
    jk_folder_path = os.path.join(current_folder_path, jk_folder_name)

    # 获取文件夹中的所有文件和子文件夹的名称列表
    file_list = os.listdir(jk_folder_path)

    jk_folder_path = os.path.join(jk_folder_path, file_list[0])

    return jk_folder_path
