import os
import pkg_resources


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



def get_ATMlog_path():
    # 获取当前文件夹的绝对路径
    current_folder_path = os.path.abspath(os.path.dirname(__file__))

    # 要进入的文件夹的名字
    ATMlog_folder_name = "ATMlog"

    # ATMlog文件夹
    ATMlog_folder_path = os.path.join(current_folder_path, ATMlog_folder_name)

    # 获取文件夹中的所有文件和子文件夹的名称列表
    file_list = os.listdir(ATMlog_folder_path)

    ATMlog_folder_path = os.path.join(ATMlog_folder_path, file_list[0])

    return ATMlog_folder_path


def get_requirment():
    '''获取requirement文件'''

    # 获取当前环境中的所有安装的包
    installed_packages = pkg_resources.working_set

    # 创建一个空列表来存储包名和版本号
    packages = []

    # 遍历所有安装的包，并将包名和版本号添加到列表中
    for package in installed_packages:
        packages.append(f"{package.key}=={package.version}")

    # 将包名和版本号写入 requirements.txt 文件
    with open("requirements.txt", "w") as file:
        file.write("\n".join(packages))

get_requirment()