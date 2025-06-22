import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QGroupBox,
                            QFormLayout, QMessageBox, QTabWidget, QTableWidget, QTableWidgetItem,
                            QHeaderView, QAbstractItemView, QSplitter)
from PyQt5.QtCore import Qt
from file_utils import FileUtils
from file_manipulator import FileManipulator

class FileManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文件操作工具")
        self.setGeometry(100, 100, 800, 600)
        
        # 加载配置
        self.config = FileUtils.load_config()
        FileUtils.head_list = self.config.get('head_list', FileUtils.head_list)
        
        # 主控件和布局 - 使用选项卡
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget, 1)  # 选项卡占据大部分空间
        
        # 创建三个选项卡页面（日志区域已移动到各自页面）
        self.create_main_tab()
        self.create_date_tab()
        self.create_config_tab()
        
        # 初始化文件管理器
        self.file_manipulator = None

    def create_main_tab(self):
        """创建主操作选项卡（包含日志区域）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 使用分割器使日志区域可调整大小
        splitter = QSplitter(Qt.Vertical)
        
        # 上半部分：操作区域
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        
        # 路径输入区域
        path_group = QGroupBox("路径设置")
        path_layout = QFormLayout()
        
        self.old_path_edit = QLineEdit(self.config.get('default_old_path', ''))
        self.old_path_edit.setPlaceholderText("选择源文件夹路径...")
        self.old_path_button = QPushButton("浏览...")
        self.old_path_button.clicked.connect(self.browse_old_path)
        
        old_path_layout = QHBoxLayout()
        old_path_layout.addWidget(self.old_path_edit)
        old_path_layout.addWidget(self.old_path_button)
        
        self.new_path_edit = QLineEdit(self.config.get('default_new_path', ''))
        self.new_path_edit.setPlaceholderText("选择目标文件夹路径...")
        self.new_path_button = QPushButton("浏览...")
        self.new_path_button.clicked.connect(self.browse_new_path)
        
        new_path_layout = QHBoxLayout()
        new_path_layout.addWidget(self.new_path_edit)
        new_path_layout.addWidget(self.new_path_button)
        
        path_layout.addRow("源文件夹:", old_path_layout)
        path_layout.addRow("目标文件夹:", new_path_layout)
        path_group.setLayout(path_layout)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        self.execute_button = QPushButton("执行文件操作")
        self.execute_button.clicked.connect(self.execute_operations)
        self.execute_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        
        self.tree_button = QPushButton("显示目录结构")
        self.tree_button.clicked.connect(self.show_directory_tree)
        self.tree_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        
        button_layout.addWidget(self.execute_button)
        button_layout.addWidget(self.tree_button)
        
        # 添加组件到上半部分
        top_layout.addWidget(path_group)
        top_layout.addLayout(button_layout)
        top_layout.addStretch(1)  # 添加弹性空间
        
        # 下半部分：日志区域（高度增加）
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout()
        self.log_text_main = QTextEdit()  # 主操作页面的日志
        self.log_text_main.setReadOnly(True)
        log_layout.addWidget(self.log_text_main)
        log_group.setLayout(log_layout)
        
        # 添加部件到分割器
        splitter.addWidget(top_widget)
        splitter.addWidget(log_group)
        
        # 设置分割器初始比例（上半部40%，下半部60%）
        splitter.setSizes([int(self.height() * 0.4), int(self.height() * 0.6)])
        
        # 添加分割器到主布局
        layout.addWidget(splitter)
        
        self.tab_widget.addTab(tab, "主操作")

    def create_date_tab(self):
        """创建日期设置选项卡（包含日志区域）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 使用分割器使日志区域可调整大小
        splitter = QSplitter(Qt.Vertical)
        
        # 上半部分：操作区域
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        
        # 目标文件夹路径
        path_group = QGroupBox("目标文件夹")
        path_layout = QHBoxLayout()
        
        self.date_target_edit = QLineEdit(self.config.get('default_new_path', ''))
        self.date_target_edit.setPlaceholderText("选择目标文件夹路径...")
        self.date_target_button = QPushButton("浏览...")
        self.date_target_button.clicked.connect(lambda: self.browse_path(self.date_target_edit))
        
        path_layout.addWidget(self.date_target_edit)
        path_layout.addWidget(self.date_target_button)
        path_group.setLayout(path_layout)
        
        # 日期设置区域
        date_group = QGroupBox("日期设置")
        date_layout = QFormLayout()
        
        self.val_date_edit = QLineEdit()
        self.val_date_edit.setPlaceholderText("YYYY.MM.DD")
        date_layout.addRow("验证环境迁移日期:", self.val_date_edit)
        
        self.prod_date_edit = QLineEdit()
        self.prod_date_edit.setPlaceholderText("YYYY.MM.DD")
        date_layout.addRow("生产环境迁移日期:", self.prod_date_edit)
        
        date_group.setLayout(date_layout)
        
        # 执行按钮
        self.date_execute_button = QPushButton("执行日期设置")
        self.date_execute_button.clicked.connect(self.execute_date_setting)
        self.date_execute_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e68a00;
            }
        """)
        
        # 添加组件到上半部分
        top_layout.addWidget(path_group)
        top_layout.addWidget(date_group)
        top_layout.addWidget(self.date_execute_button)
        top_layout.addStretch(1)  # 添加弹性空间
        
        # 下半部分：日志区域（高度增加）
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout()
        self.log_text_date = QTextEdit()  # 日期设置页面的日志
        self.log_text_date.setReadOnly(True)
        log_layout.addWidget(self.log_text_date)
        log_group.setLayout(log_layout)
        
        # 添加部件到分割器
        splitter.addWidget(top_widget)
        splitter.addWidget(log_group)
        
        # 设置分割器初始比例（上半部40%，下半部60%）
        splitter.setSizes([int(self.height() * 0.4), int(self.height() * 0.6)])
        
        # 添加分割器到主布局
        layout.addWidget(splitter)
        
        self.tab_widget.addTab(tab, "日期设置")

    def create_config_tab(self):
        """创建配置选项卡（不包含日志区域）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 路径设置区域
        path_group = QGroupBox("默认路径设置")
        path_layout = QFormLayout()
        
        self.config_old_path_edit = QLineEdit(self.config.get('default_old_path', ''))
        self.config_old_path_edit.setPlaceholderText("选择源文件夹路径...")
        self.config_old_path_button = QPushButton("浏览...")
        self.config_old_path_button.clicked.connect(lambda: self.browse_path(self.config_old_path_edit))
        
        old_path_layout = QHBoxLayout()
        old_path_layout.addWidget(self.config_old_path_edit)
        old_path_layout.addWidget(self.config_old_path_button)
        
        self.config_new_path_edit = QLineEdit(self.config.get('default_new_path', ''))
        self.config_new_path_edit.setPlaceholderText("选择目标文件夹路径...")
        self.config_new_path_button = QPushButton("浏览...")
        self.config_new_path_button.clicked.connect(lambda: self.browse_path(self.config_new_path_edit))
        
        new_path_layout = QHBoxLayout()
        new_path_layout.addWidget(self.config_new_path_edit)
        new_path_layout.addWidget(self.config_new_path_button)
        
        path_layout.addRow("源文件夹:", old_path_layout)
        path_layout.addRow("目标文件夹:", new_path_layout)
        path_group.setLayout(path_layout)
        
        # head_list 编辑区域
        head_group = QGroupBox("封面文件类型配置 (head_list)")
        head_layout = QVBoxLayout()
        
        # 创建表格
        self.head_list_table = QTableWidget()
        self.head_list_table.setColumnCount(1)
        self.head_list_table.setHorizontalHeaderLabels(["文件类型前缀"])
        self.head_list_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.head_list_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.head_list_table.setEditTriggers(QAbstractItemView.DoubleClicked)
        
        # 添加按钮
        btn_layout = QHBoxLayout()
        self.add_head_btn = QPushButton("添加")
        self.add_head_btn.clicked.connect(self.add_head_item)
        self.remove_head_btn = QPushButton("删除")
        self.remove_head_btn.clicked.connect(self.remove_head_item)
        
        btn_layout.addWidget(self.add_head_btn)
        btn_layout.addWidget(self.remove_head_btn)
        
        head_layout.addWidget(self.head_list_table)
        head_layout.addLayout(btn_layout)
        head_group.setLayout(head_layout)
        
        # 保存按钮
        self.save_config_button = QPushButton("保存配置")
        self.save_config_button.clicked.connect(self.save_config)
        self.save_config_button.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #7b1fa2;
            }
        """)
        
        # 添加组件到布局
        layout.addWidget(path_group)
        layout.addWidget(head_group, 1)  # head_group占据更多空间
        layout.addWidget(self.save_config_button)
        
        # 初始化表格数据
        self.load_head_list_table()
        
        self.tab_widget.addTab(tab, "配置")

    def load_head_list_table(self):
        """加载head_list到表格"""
        self.head_list_table.setRowCount(len(FileUtils.head_list))
        for i, item in enumerate(FileUtils.head_list):
            self.head_list_table.setItem(i, 0, QTableWidgetItem(item))

    def add_head_item(self):
        """添加新的head_list项"""
        row_count = self.head_list_table.rowCount()
        self.head_list_table.insertRow(row_count)
        self.head_list_table.setItem(row_count, 0, QTableWidgetItem(""))

    def remove_head_item(self):
        """删除选中的head_list项"""
        current_row = self.head_list_table.currentRow()
        if current_row >= 0:
            self.head_list_table.removeRow(current_row)

    def save_config(self):
        """保存配置"""
        # 更新head_list
        FileUtils.head_list = []
        for i in range(self.head_list_table.rowCount()):
            item = self.head_list_table.item(i, 0)
            if item and item.text():
                FileUtils.head_list.append(item.text())
        
        # 更新配置字典
        self.config['head_list'] = FileUtils.head_list
        self.config['default_old_path'] = self.config_old_path_edit.text()
        self.config['default_new_path'] = self.config_new_path_edit.text()
        
        # 保存到文件
        FileUtils.save_config(self.config)
        
        # 更新主界面的默认值
        self.old_path_edit.setText(self.config['default_old_path'])
        self.new_path_edit.setText(self.config['default_new_path'])
        self.date_target_edit.setText(self.config['default_new_path'])
        
        QMessageBox.information(self, "配置保存", "配置已成功保存！")

    def browse_old_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择源文件夹")
        if path:
            self.old_path_edit.setText(path)

    def browse_new_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择目标文件夹")
        if path:
            self.new_path_edit.setText(path)
            # 同时更新日期设置页的目标文件夹
            self.date_target_edit.setText(path)

    def browse_path(self, target_edit):
        """通用路径浏览方法"""
        path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if path:
            target_edit.setText(path)

    def log_message(self, message):
        """将消息添加到当前活动选项卡的日志区域"""
        current_tab = self.tab_widget.currentIndex()
        log_text = None
        
        if current_tab == 0:  # 主操作选项卡
            log_text = self.log_text_main
        elif current_tab == 1:  # 日期设置选项卡
            log_text = self.log_text_date
        
        if log_text:
            log_text.append(message)
            # 自动滚动到底部
            log_text.verticalScrollBar().setValue(
                log_text.verticalScrollBar().maximum()
            )

    def execute_operations(self):
        """执行文件操作"""
        old_path = self.old_path_edit.text().strip()
        new_path = self.new_path_edit.text().strip()
        
        if not old_path or not new_path:
            QMessageBox.warning(self, "路径错误", "请选择源文件夹和目标文件夹路径！")
            return
            
        if not os.path.exists(old_path):
            QMessageBox.warning(self, "路径错误", "源文件夹路径不存在！")
            return
            
        # 清空日志
        self.log_text_main.clear()
        
        # 创建文件操作器
        self.file_manipulator = FileManipulator(
            old_path, new_path, {}, self.log_message
        )
        
        # 执行操作
        self.log_message("开始文件操作流程...")
        success = self.file_manipulator.execute_operations()
        
        if success:
            self.log_message("\n✅ 所有操作成功完成！")
        else:
            self.log_message("\n❌ 操作过程中出现错误！")

    def execute_date_setting(self):
        """执行日期设置操作"""
        target_dir = self.date_target_edit.text().strip()
        val_date = self.val_date_edit.text().strip()
        prod_date = self.prod_date_edit.text().strip()
        
        if not target_dir:
            QMessageBox.warning(self, "路径错误", "请选择目标文件夹路径！")
            return
            
        if not val_date:
            QMessageBox.warning(self, "输入错误", "请输入验证环境迁移日期！")
            return
            
        if not os.path.exists(target_dir):
            QMessageBox.warning(self, "路径错误", "目标文件夹路径不存在！")
            return
            
        # 清空日志
        self.log_text_date.clear()
        
        # 创建文件操作器（不需要源路径）
        self.file_manipulator = FileManipulator(
            "", target_dir, {}, self.log_message
        )
        
        # 执行日期设置
        self.log_message("开始设置迁移日期...")
        success = self.file_manipulator.edt_A2_docx(target_dir, val_date, prod_date)
        
        if success:
            self.log_message("\n✅ 日期设置成功完成！")
        else:
            self.log_message("\n❌ 日期设置过程中出现错误！")

    def show_directory_tree(self):
        """显示目录结构"""
        path = self.new_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "路径错误", "请选择目标文件夹路径！")
            return
            
        if not os.path.exists(path):
            QMessageBox.warning(self, "路径错误", "目标文件夹路径不存在！")
            return
            
        # 清空日志
        self.log_text_main.clear()
        
        if not self.file_manipulator:
            self.file_manipulator = FileManipulator(
                "", path, {}, self.log_message
            )
            
        self.log_message("目录结构:\n")
        tree = self.file_manipulator.get_directory_tree(path)
        self.log_message(tree)