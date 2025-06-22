import re
import os
import json
import docx
from docx.enum.text import WD_COLOR_INDEX

class FileUtils:
    """严格禁止实例化的工具类"""
    
    # 定义默认的head_list
    head_list = [
        'Analysis', 'CofA Step', 'Events', 'Format Calculation', 'LIMS Constant',
        'LIMS Users', 'Label Printer', 'Lists', 'Lot Template', 'Product',
        'Report Template', 'Stock', 'Subroutine', 'Suppliers', 'T PH AQL Sample Plan',
        'T PH Item Code', 'T PH Sample Plan', 'T_PH_Grade', 'T_PH_Spec Type', 'T_Report Text',
        'Table Master', 'Table Template', 'Units', 'User Dialog', 'vendor', 'Stage'
    ]

    @staticmethod
    def save_config(config):
        """保存配置到文件"""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    @staticmethod
    def load_config():
        """从文件加载配置"""
        config = {
            'head_list': FileUtils.head_list,
            'default_old_path': '',
            'default_new_path': '',
        }
        
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    config.update(loaded_config)
        except Exception as e:
            print(f"加载配置失败: {e}")
        
        return config

    @staticmethod
    def recursively_delete_contents(folder_path):
        """
        递归删除文件夹内的所有内容。

        参数:
            folder_path (str): 要删除内容的文件夹路径。

        返回:
            None
        """
        # 遍历文件夹中的所有子项
        for item_name in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item_name)

            # 检查是否为文件或文件夹
            if os.path.isfile(item_path):
                # 如果是文件，则删除
                os.remove(item_path)
            elif os.path.isdir(item_path):
                # 如果是文件夹，则递归删除该文件夹及内部所有内容
                FileUtils.recursively_delete_contents(item_path)
                # 删除空文件夹
                os.rmdir(item_path)

    @staticmethod
    def increment_filename_number(filename: str, start_sep: str = '', end_sep: str = '') -> str:
        """
        文件名数字递增方法
        
        参数:
            filename: 原始文件名
            start_sep: 起始定位字符串
            end_sep: 结束定位字符串
            
        返回:
            处理后的新文件名
        """
        if start_sep != '':
            # 找到起始定位字符串第一次出现的位置
            start_index = filename.find(start_sep)
            if start_index == -1:
                return filename  # 未找到起始定位字符串
            
            # 计算起始搜索位置（跳过起始字符串）
            start_search = start_index + len(start_sep)
        else :
            start_search = 0
            
        if end_sep != '':
            # 找到结束定位字符串最后一次出现的位置
            end_index = filename.rfind(end_sep, start_search)
            if end_index == -1:
                return filename  # 未找到结束定位字符串
        else:
            end_index = len(filename)
        
        # 提取目标范围内的内容
        target_str = filename[start_search:end_index]
        
        # 在目标字符串中查找第一个连续数字序列
        match = re.search(r'\d+', target_str)
        if not match:
            return filename  # 未找到数字序列
        
        # 提取数字部分和原始位数
        number_part = match.group()
        original_digits = len(number_part)
        start_pos = match.start()
        end_pos = match.end()
        
        # 转换为整数并+1
        try:
            number_value = int(number_part)
            new_number = number_value + 1
            new_number_str = str(new_number)
        except ValueError:
            return filename
        
        # 保留原始位数格式（前导零）
        if len(new_number_str) < original_digits:
            new_number_str = new_number_str.zfill(original_digits)
        
        # 构建新目标字符串（仅替换数字部分）
        new_target_str = (
            target_str[:start_pos] + 
            new_number_str + 
            target_str[end_pos:]
        )
        
        # 构建完整新文件名
        return (
            filename[:start_search] + 
            new_target_str + 
            filename[end_index:]
        )

    @staticmethod
    def is_str_number(s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    @staticmethod
    def run_paragraph(cells):
        """加红色底纹"""
        for cell in cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.RED
        return

    @staticmethod
    def edt_docx(doc_path: str, doc_name: str):
        file_path = os.path.join(doc_path, doc_name)

        # 使用类变量head_list
        if any(doc_name.startswith(head) for head in FileUtils.head_list):
            # 处理封面文件
            doc = docx.Document(file_path)
            tab = doc.tables[0]

            # 获取目标单元格
            target_cell = tab.rows[2].cells[0]

            # 保存原始格式的文本运行(runs)
            original_runs = target_cell.paragraphs[0].runs.copy()

            # 清除单元格内容
            for paragraph in target_cell.paragraphs:
                p = paragraph._element
                p.getparent().remove(p)

            # 添加新段落
            new_paragraph = target_cell.add_paragraph()

            # 添加新文本运行，并复制原始格式
            if original_runs:
                # 使用第一个原始运行的格式作为基准
                base_run = original_runs[0]
                new_run = new_paragraph.add_run(doc_name.replace(".docx", ""))

                # 复制字体格式
                new_run.font.name = base_run.font.name
                new_run.font.size = base_run.font.size
            else:
                # 如果没有原始运行，直接添加文本
                new_paragraph.add_run(doc_name.replace(".docx", ""))
            doc.save(file_path)
        elif "REC-Q680003-A2" in doc_name:
            # 如果是"REC-Q680003-A2-01  LIMS数据迁移表单"
            doc = docx.Document(file_path)
            tables = doc.tables

            for tab in tables:
                if tab.rows and "数据包名称" in tab.rows[0].cells[0].text:
                    # 修改表头
                    new_text = FileUtils.increment_filename_number(tab.rows[0].cells[1].text)
                    tab.rows[0].cells[1].text = new_text
                else:
                    for t_row in tab.rows:
                        if t_row.cells and FileUtils.is_str_number(t_row.cells[0].text):
                            # 添加红色底纹
                            FileUtils.run_paragraph(t_row.cells)
            doc.save(file_path)

        elif "REC-Q680003-A5" in doc_name:
            # 如果是"REC-Q680003-A5-01  LIMS主数据申请表"
            doc = docx.Document(file_path)
            tables = doc.tables
            rows = tables[0].rows
            rows_index = [2, 3, 5, 7]  # 添加红色底纹的行
            
            new_text = FileUtils.increment_filename_number(rows[0].cells[2].text)
            rows[0].cells[2].text = new_text
            
            for r in rows_index:
                if len(rows) > r:
                    FileUtils.run_paragraph(rows[r].cells)

            if len(tables) > 2:
                for t_row in tables[2].rows:
                    if t_row.cells and t_row.cells[0].text == new_text:
                        FileUtils.run_paragraph(t_row.cells)

            doc.save(file_path)

    @staticmethod
    def edit_A2_docx(current_directory: str, file_name: str, to_val_date: str, to_prod_date: str):
        """批量修改迁移日期"""
        file_path = os.path.join(current_directory, file_name)
        doc = docx.Document(file_path)
        tables = doc.tables

        for tab in tables:
            if tab.rows and "数据包名称" in tab.rows[0].cells[0].text:
                continue
                
            for t_row in tab.rows:
                if not t_row.cells:
                    continue
                    
                if FileUtils.is_str_number(t_row.cells[0].text):
                    if len(t_row.cells) < 5:
                        continue
                        
                    # 获取单元格引用
                    val_data_cell = t_row.cells[3]
                    prod_data_cell = t_row.cells[4]
                    
                    # 处理验证日期单元格
                    # 保存原始段落格式
                    val_alignment = None
                    if val_data_cell.paragraphs:
                        val_alignment = val_data_cell.paragraphs[0].alignment
                    
                    # 清除现有内容
                    val_paragraphs = list(val_data_cell.paragraphs)
                    for para in val_paragraphs:
                        p = para._element
                        p.getparent().remove(p)
                    
                    # 添加新内容并继承原始对齐方式
                    val_para = val_data_cell.add_paragraph()
                    val_run = val_para.add_run(to_val_date)
                    if val_alignment is not None:
                        val_para.alignment = val_alignment
                    
                    # 处理生产日期单元格
                    # 保存原始段落格式
                    prod_alignment = None
                    if prod_data_cell.paragraphs:
                        prod_alignment = prod_data_cell.paragraphs[0].alignment
                    
                    # 清除现有内容
                    prod_paragraphs = list(prod_data_cell.paragraphs)
                    for para in prod_paragraphs:
                        p = para._element
                        p.getparent().remove(p)
                    
                    # 添加新内容并继承原始对齐方式
                    prod_para = prod_data_cell.add_paragraph()
                    prod_run = prod_para.add_run(to_prod_date)
                    if prod_alignment is not None:
                        prod_para.alignment = prod_alignment
                    
                    # 设置字体格式（保留原有逻辑）
                    base_runs = None
                    if val_paragraphs and val_paragraphs[0].runs:
                        base_runs = val_paragraphs[0].runs
                    elif prod_paragraphs and prod_paragraphs[0].runs:
                        base_runs = prod_paragraphs[0].runs
                    
                    if base_runs:
                        base_run = base_runs[0]
                        val_run.font.name = base_run.font.name
                        val_run.font.size = base_run.font.size
                        prod_run.font.name = base_run.font.name
                        prod_run.font.size = base_run.font.size

        # 统一保存文档
        doc.save(file_path)