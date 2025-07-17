import os
import shutil
import time
import datetime
from file_utils import FileUtils
import csv

class FileManipulator:
    def __init__(self, str_oldpath: str, str_newpath: str, max_file_dict: dict, output_callback=None):
        self.str_oldpath = str_oldpath
        self.str_newpath = str_newpath
        self.max_file_dict = max_file_dict
        self.output_callback = output_callback

    def log(self, message):
        """记录日志信息，如果有回调函数则使用它，否则打印到控制台"""
        if self.output_callback:
            self.output_callback(message)
        else:
            print(message)

    def cp_files(self):
        """
        将文件夹从 'str_oldpath' 复制到 'str_newpath'，文件夹中的文件名会根据各类别中文件名的最高数字索引进行更新。
        如果 'str_newpath' 目录已经存在，会将目标文件复制移动到一个带有时间戳的同名目录，然后清空目标文件。
        """
        self.log("开始复制文件...")
        str_oldpath = self.str_oldpath
        str_newpath = self.str_newpath
        max_file_dict = self.max_file_dict
        
        # 检查源目录是否存在
        if not os.path.exists(str_oldpath):
            self.log(f"源目录不存在：{str_oldpath}")
            return False

        # 检查目标目录，如果已存在，则重命名为原名称_时间戳，并创建空的目标目录
        if os.path.exists(str_newpath):
            # 获取当前时间戳
            timestamp = int(time.time())
            # 构建新的目录名
            new_target_path = f"{str_newpath}_{timestamp}"
            # 重命名现有目录
            os.rename(str_newpath, new_target_path)
            self.log(f"目标目录已存在，已重命名为: {new_target_path}")
            # 创建新的空目录
            os.makedirs(str_newpath)
            self.log(f"已创建新目录: {str_newpath}")
        else:
            # 如果目标目录不存在，直接创建
            os.makedirs(str_newpath)
            self.log(f"已创建目标目录: {str_newpath}")

        files = os.listdir(str_oldpath)

        for f_name in files:
            i_var1 = f_name.rfind("-")  # -所在的位置
            # 新增封面文件判断逻辑处理
            if i_var1 != -1 and os.path.isdir(os.path.join(str_oldpath, f_name)):
                try:
                    file_index = int(f_name[i_var1 + 1:])
                    file_code = f_name[i_var1 + 1:]
                    file_class = f_name[0:i_var1]

                    last_index = int(max_file_dict.get(file_class, 0))

                    if file_index > last_index:
                        max_file_dict[file_class] = file_code
                        self.log(f"更新类别 '{file_class}' 的最大索引为: {file_code}")
                except ValueError:
                    self.log(f"文件名 '{f_name}' 的数字部分无效，跳过处理")

        for key, value in max_file_dict.items():
            # 在新文件名中增加索引数字
            try:
                new_code = "{:04d}".format(int(value) + 1)
            except ValueError:
                self.log(f"类别 '{key}' 的索引值 '{value}' 无效，跳过处理")
                continue

            # 源文件夹和目标文件夹路径
            source_folder = os.path.join(str_oldpath, f"{key}-{value}")
            target_folder = os.path.join(str_newpath, f"{key}-{new_code}")

            # 检查源文件夹是否存在
            if not os.path.exists(source_folder):
                self.log(f"源文件夹不存在：{source_folder}")
                continue

            # 如果目标文件夹不存在，则复制源文件夹
            if not os.path.exists(target_folder):
                try:
                    shutil.copytree(source_folder, target_folder)
                    # 复制文件属性并设置时间戳
                    shutil.copystat(source_folder, target_folder)
                    os.utime(target_folder, (time.time(), time.time()))
                    self.log(f"已复制: {source_folder} -> {target_folder}")
                except Exception as e:
                    self.log(f"复制文件夹时出错: {e}")
            else:
                self.log(f"文件夹已存在于：{target_folder}")

        self.max_file_dict = max_file_dict

        # 复制封面文件
        for f_name in files:
            source_file = os.path.join(str_oldpath, f_name)
            if not os.path.isdir(source_file):
                target_file = os.path.join(str_newpath, f_name)

                # 如果目标文件不存在，则复制
                if not os.path.exists(target_file):
                    try:
                        shutil.copy(source_file, target_file)
                        # 复制文件属性并设置时间戳
                        shutil.copystat(source_file, target_file)
                        os.utime(target_file, (time.time(), time.time()))
                        self.log(f"已复制文件: {f_name}")
                    except Exception as e:
                        self.log(f"复制文件时出错: {e}")
                else:
                    self.log(f"文件已存在于：{target_file}")

        self.log("文件复制完成")
        return True

    def del_files(self):
        """
        删除给定目标路径下的文件

        参数:
            str_tarpath (str): 执行删除操作的目标路径。

        返回:
            None
        """
        self.log("开始删除临时文件...")
        str_tarpath = self.str_newpath

        try:
            # 遍历目标路径下的所有子项：Analysis、Product...
            for folder_item in os.listdir(str_tarpath):
                folder_path = os.path.join(str_tarpath, folder_item)

                # 判断是否为文件夹
                if not os.path.isdir(folder_path):
                    continue  # 跳过非文件夹项

                # 遍历文件夹内的所有子项：Data Pack、证据...
                for item_name in os.listdir(folder_path):
                    item_path = os.path.join(folder_path, item_name)

                    # 判断文件名以 ~$ 开头
                    if item_name.startswith('~$'):
                        # 删除临时文件
                        try:
                            os.remove(item_path)
                            self.log(f"已删除临时文件: {item_name}")
                        except Exception as e:
                            self.log(f"删除临时文件时出错: {e}")
                    elif os.path.isdir(item_path):
                        # 如果是文件夹，则递归删除内部所有内容
                        try:
                            FileUtils.recursively_delete_contents(item_path)
                            self.log(f"已删除文件夹: {item_path}")
                        except Exception as e:
                            self.log(f"删除文件夹时出错: {e}")
        except (FileNotFoundError, PermissionError, OSError) as e:
            self.log(f"发生错误：{e}")
            return False

        self.log("临时文件删除完成")
        return True

    def ren_files(self):
        """
        对指定目录下文件进行重命名：0038-->0039
        """
        self.log("开始重命名文件...")
        str_tarpath = self.str_newpath
        try:
            # 遍历目标路径下的所有子项：Analysis、Product...
            for folder_item in os.listdir(str_tarpath):
                folder_path = os.path.join(str_tarpath, folder_item)

                # 判断是否为文件夹
                if not os.path.isdir(folder_path):
                    # 不是文件夹  则是封面文件
                    newFileName = FileUtils.increment_filename_number(folder_item)
                    new_file_path = os.path.join(str_tarpath, newFileName)
                    try:
                        os.rename(folder_path, new_file_path)  # 使用完整路径重命名
                        self.log(f"已重命名文件: {folder_item} -> {newFileName}")
                    except Exception as e:
                        self.log(f"重命名文件时出错: {e}")
                    continue  # 跳过非文件夹项

                # 遍历文件夹内的所有子项：Data Pack、证据...
                for item_name in os.listdir(folder_path):
                    item_path = os.path.join(folder_path, item_name)

                    # 检查是否为文件，且文件名不以 ~$ 开头
                    if os.path.isfile(item_path) and not item_name.startswith('~$'):
                        # 设置分隔符  括号
                        start_bracket = '(' if item_name.find('(') != -1 else '（'
                        end_bracket = ')' if item_name.find(')') != -1 else '）'

                        # 新的文件名称
                        new_file_name = FileUtils.increment_filename_number(item_name,start_bracket,end_bracket)
                        # 定义新生成文件的路径
                        new_file_path = os.path.join(folder_path, new_file_name)

                        # 重命名文件
                        try:
                            os.rename(item_path, new_file_path)
                            self.log(f"已重命名: {item_name} -> {new_file_name}")
                        except Exception as e:
                            self.log(f"重命名文件时出错: {e}")
        except (FileNotFoundError, PermissionError, OSError) as e:
            self.log(f"发生错误：{e}")
            return False

        self.log("文件重命名完成")
        return True

    def edt_docx(self):
        self.log("开始编辑Word文档...")
        str_tarpath = self.str_newpath

        for head_file_name in os.listdir(str_tarpath):
            current_directory = os.path.join(str_tarpath, head_file_name)

            if not os.path.isdir(current_directory):
                try:
                    FileUtils.edt_docx(str_tarpath, head_file_name)
                    self.log(f"已编辑封面: {head_file_name}")  # 添加日志
                except Exception as e:
                    self.log(f"编辑封面文件时出错: {e}")
                continue

            for file_name in os.listdir(current_directory):
                file_path = os.path.join(current_directory, file_name)
                try:
                    if "REC-Q680003-A2" in file_name:
                        FileUtils.edt_docx(current_directory, file_name)
                        self.log(f"已编辑迁移表: {file_name}")
                    elif "REC-Q680003-A5" in file_name:
                        FileUtils.edt_docx(current_directory, file_name)
                        self.log(f"已编辑申请表: {file_name}")
                except Exception as e:
                    self.log(f"编辑Word文档时出错: {e}")

        self.log("Word文档编辑完成")
        return True
    
    def edt_A2_docx(self, target_dir: str, to_val_date: str, to_prod_date: str = ''):
        """递归修改A2文档中的日期"""        
        if not to_val_date:
            self.log("错误的迁移日期，工作终止")
            return False
        
        # 处理默认的生产日期
        if not to_prod_date:
            to_prod_date = to_val_date
        
        # 遍历目标目录中的所有项目
        for item in os.listdir(target_dir):
            item_path = os.path.join(target_dir, item)
            
            if os.path.isdir(item_path):
                # 递归处理子目录
                self.edt_A2_docx(item_path, to_val_date, to_prod_date)
            else:
                try:
                    if "REC-Q680003-A2" in item:
                        # 处理单个文件
                        FileUtils.edit_A2_docx(target_dir, item, to_val_date, to_prod_date)
                        self.log(f"已修改: {item}")
                except Exception as e:
                    self.log(f"修改《{item}》时出错: {e}")
        return True

    def read_A2_to_csv(self,output_csv):
        # 初始化数据存储
        all_data = []
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        a2_csv_path = os.path.join(output_csv,f"A2_{timestamp}.csv")

        name_contains='REC-Q680003-A2'
        extension='docx'

        self.log(f"开始在目录 '{self.str_newpath}' 中查找A2文档...") 
        pathes = FileUtils.find_files_by_name(self.str_newpath,name_contains,extension)
        self.log(f"共找到 {len(pathes)} 个A2文档。") 

        # 处理每个文档并收集数据
        total = len(pathes)
        for i, path_a2 in enumerate(pathes): 
            self.log(f"正在处理第 {i+1}/{total} 个文件: {os.path.basename(path_a2)}") 
            a2_data = FileUtils.read_A2(path_a2)
            if a2_data:
                all_data.extend(a2_data)
        
        # 写入CSV文件
        if all_data:
            title=['包名称', '记录名称', '迁移验证环境日期', '迁移正式环境日期']
            self.log(f"开始将 {len(all_data)} 条数据写入到 {a2_csv_path}")
            FileUtils.write_to_csv(all_data, a2_csv_path,title)
            self.log(f"成功生成CSV文件: {a2_csv_path}")
        else:
            self.log("未找到有效数据，未生成CSV文件。")

    def read_A5_to_csv(self,output_csv):
        name_contains = "REC-Q680003-A5"
        extension = "docx"
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        # 修改文件路径定义
        a5_tb1_path = os.path.join(
            output_csv, 
            f"A5_tb1_{timestamp}.csv"  # 添加时间戳
        )
        a5_tb2_path = os.path.join(
            output_csv, 
            f"A5_tb2_{timestamp}.csv"  # 添加时间戳
        )

        self.log(f"开始在目录 '{self.str_newpath}' 中查找A5文档...") 
        pathes = FileUtils.find_files_by_name(self.str_newpath, name_contains, extension)
        self.log(f"共找到 {len(pathes)} 个A5文档。") 
        
        # 准备收集数据的列表
        all_tb1_data = []  # 存储所有文档的tb1数据
        all_tb2_data = []  # 存储所有文档的tb2数据
        
        # 定义CSV表头
        tb1_title = ['包名称', '理由', '相关文件']
        tb2_title = ['包名称', '记录名称', '操作类型', '分类','风险评估']

        total = len(pathes)
        for i, path_a5 in enumerate(pathes):
            self.log(f"正在处理第 {i+1}/{total} 个文件: {os.path.basename(path_a5)}")            
            try:
                # 获取当前文档的数据
                result = FileUtils.read_A5(path_a5)
                
                # 添加结果有效性检查
                if len(result) >= 1 and result[0]:
                    all_tb1_data.append(result[0])
                    
                if len(result) >= 2 and result[1]:
                    all_tb2_data.extend(result[1])
                    
            except Exception as e:
                self.log(f"处理文件 {path_a5} 时发生未捕获错误: {str(e)}")
                continue  # 继续处理下一个文件
        
        # 写入CSV文件
        if all_tb1_data:
           FileUtils.write_to_csv(all_tb1_data, a5_tb1_path, tb1_title)
           self.log(f"成功生成A5表1的CSV文件: {a5_tb1_path}") 
        else:
            self.log("警告: 没有收集到表1数据")
        
        if all_tb2_data:
            FileUtils.write_to_csv(all_tb2_data, a5_tb2_path, tb2_title)
            self.log(f"成功生成A5表2的CSV文件: {a5_tb2_path}")
        else:
            self.log("警告: 没有收集到表2数据")
    
    def get_directory_tree(self, path: str, indent=0):
        """
        获取指定路径下的文件树结构。
        :param path: 要打印的目录路径
        :param indent: 当前缩进级别
        :return: 目录树字符串
        """
        tree = ""
        try:
            for entry in os.listdir(path):
                entry_path = os.path.join(path, entry)
                if os.path.isfile(entry_path):
                    tree += ' ' * indent + '|-- ' + entry + '\n'
                elif os.path.isdir(entry_path):
                    tree += ' ' * indent + '|-- ' + entry + '\n'
                    tree += self.get_directory_tree(entry_path, indent + 4)
        except Exception as e:
            self.log(f"获取目录树时出错: {e}")
        return tree

    def execute_operations(self):
        self.log("=" * 50)
        self.log("开始执行文件操作流程")
        self.log("=" * 50)
        
        if not self.cp_files():
            self.log("文件复制失败，中止操作")
            return False
            
        if not self.del_files():
            self.log("文件删除失败，中止操作")
            return False
            
        if not self.ren_files():
            self.log("文件重命名失败，中止操作")
            return False
            
        if not self.edt_docx():
            self.log("文件编辑失败，中止操作")
            return False
            
        self.log("=" * 50)
        self.log("所有操作成功完成")
        self.log("=" * 50)
        return True