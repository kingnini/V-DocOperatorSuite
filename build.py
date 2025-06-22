import os
import PyInstaller.__main__

# 主程序文件
main_script = 'main.py'

# 程序名称
app_name = 'FileManagerTool'

# 添加数据文件（非Python文件）
data_files = [
    ('config.json', '.'),
]

# 构建PyInstaller命令
args = [
    '--name={}'.format(app_name),
    '--onefile',
    '--windowed',
    '--clean',
    '--add-data={}'.format(';'.join([f'{src}{os.pathsep}{dst}' for src, dst in data_files])),
    main_script
]

# 执行打包命令
PyInstaller.__main__.run(args)