import os
import subprocess


class SystemTools:
    """
    操作系统及第三方软件交互的工具类
    """

    @staticmethod
    def open_with_wps(file_path):
        """
        使用指定的 WPS 程序打开目标文件。

        :param file_path: 需要打开的文件绝对路径
        """
        # 指定你电脑上的 WPS 启动程序路径
        wps_exe_path = r'F:\APP\wps\WPS Office\ksolaunch.exe'

        # 1. 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"❌ 错误：要打开的目标文件不存在 -> {file_path}")
            return

        # 2. 检查 WPS 路径是否正确
        if not os.path.exists(wps_exe_path):
            print(f"❌ 错误：找不到 WPS 执行文件，请检查路径 -> {wps_exe_path}")
            return

        print(f"🚀 正在唤起 WPS 打开文件...")

        try:
            # 使用 subprocess.Popen 打开外部程序，这不会阻塞你的 Python 运行
            # 把 wps_exe_path 和 target_file 作为列表传入，可以完美解决路径中有空格的问题
            subprocess.Popen([wps_exe_path, file_path])
            print(f"✅ 文件已在 WPS 中打开！")
        except Exception as e:
            print(f"❌ 自动打开 WPS 失败: {e}")
