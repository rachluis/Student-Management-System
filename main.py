import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QDialog, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QStatusBar, QPushButton, QLineEdit,
    QInputDialog, QMessageBox, QLabel, QHeaderView, QMenu
)
from PyQt6.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import pyqtgraph as pg
from mpl_toolkits.mplot3d import Axes3D
import numpy as np
from PyQt6.QtGui import QPalette, QColor, QIcon
from PyQt6.QtCore import QPoint, Qt, QSize
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("登录")
        self.setFixedSize(400, 250)

        # 读取账号文件
        self.total_failed_attempts = 0
        self.accounts = self.load_accounts()

        # 创建布局
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        # 用户名输入
        self.username_label = QLabel("用户名：")
        self.username_input = QLineEdit()
        self.username_input.setMaxLength(20)
        self.username_input.textChanged.connect(self.validate_input)

        # 密码输入
        self.password_label = QLabel("密码：")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.textChanged.connect(self.validate_input)

        # 登录按钮
        self.login_button = QPushButton("登录")
        self.login_button.setEnabled(False)
        self.login_button.clicked.connect(self.check_login)

        # 添加到布局
        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        layout.addWidget(self.login_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # 添加右下方文字
        self.footer_label = QLabel("刘蕊 0221147010")
        layout.addWidget(self.footer_label, alignment=Qt.AlignmentFlag.AlignRight)

        self.setLayout(layout)
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                          stop:0 #E0F7FA, stop:1 #B2EBF2);
            }
            QLabel {
                color: #006064;
                font-size: 12px;
                font-weight: bold;
            }
            QLineEdit {
                background-color: white;
                border: 2px solid #00838F;
                border-radius: 10px;
                padding: 8px;
                font-size: 12px;
                min-height: 20px;
            }
            QLineEdit:focus {
                border: 2px solid #00ACC1;
                background-color: #F5FCFF;
            }
            QPushButton {
                background-color: #00ACC1;
                color: white;
                border-radius: 15px;
                padding: 10px 20px;
                font-size: 16px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #00838F;
            }
            QPushButton:disabled {
                background-color: #B2DFDB;
                color: #E0E0E0;
            }
            #footer_label {
                color: gray;
                font-size: 10px;
                font-weight: normal;
            }
        """)

    def load_accounts(self):
        file_path = "data/user.txt"
        accounts = {}
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                for line in file:
                    if ":" in line:
                        username, password = line.strip().split(":", 1)
                        accounts[username] = password
                    else:
                        print(f"警告：跳过格式错误的行: {line.strip()}")
            return accounts
        except FileNotFoundError:
            QMessageBox.critical(self, "错误", "账号文件 (user.txt) 未找到！")
            return {}
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取账号文件时出错：{str(e)}")
            return {}

    def validate_input(self):
        username = self.username_input.text()
        password = self.password_input.text()

        if len(username) > 0 and len(password) >= 6:
            self.login_button.setEnabled(True)
        else:
            self.login_button.setEnabled(False)

        if len(password) > 0 and len(password) < 6:
            self.password_label.setText("密码：(最少6个字符)")
        else:
            self.password_label.setText("密码：")

    def check_login(self):
        username = self.username_input.text()
        password = self.password_input.text()

        # 检查总失败次数
        if self.total_failed_attempts >= 5:
            QMessageBox.critical(self, "错误",
                                 "已达到最大尝试次数！\n请重启程序重试。")
            self.close()  # 关闭登录窗口
            return

        # 验证账号密码
        if not self.accounts:
            QMessageBox.critical(self, "错误", "无法验证账号，请检查账号文件！")
            return

        if username in self.accounts and self.accounts[username] == password:
            QMessageBox.information(self, "成功", "登录成功！")
            self.accept()
        else:
            # 更新总失败次数
            self.total_failed_attempts += 1
            remaining_attempts = 5 - self.total_failed_attempts

            if remaining_attempts <= 0:
                QMessageBox.critical(self, "错误",
                                     "已达到最大尝试次数！\n请重启程序重试。")
                self.close()
            else:
                QMessageBox.warning(self, "错误",
                                    f"用户名或密码错误！\n剩余尝试次数：{remaining_attempts}")

            self.password_input.clear()
            self.login_button.setEnabled(False)


class FilterHeader(QHeaderView):
    def __init__(self, parent):
        super().__init__(Qt.Orientation.Horizontal, parent)
        self.setSectionsClickable(True)
        self.sectionClicked.connect(self.on_section_clicked)
        self.filters = {}
        self.parent = parent

    def on_section_clicked(self, logical_index):
        menu = QMenu(self)

        # Get unique values for the column
        column_data = []
        for row in range(self.parent.rowCount()):
            item = self.parent.item(row, logical_index)
            if item:
                column_data.append(item.text())
        unique_values = sorted(set(column_data))

        # Add filter options
        for value in unique_values:
            action = menu.addAction(value)
            action.setCheckable(True)
            if logical_index in self.filters and value in self.filters[logical_index]:
                action.setChecked(True)
            action.triggered.connect(lambda checked, v=value, col=logical_index:
                                     self.apply_filter(col, v, checked))

        # Add "Clear Filters" option
        menu.addSeparator()
        clear_action = menu.addAction("清除筛选")
        clear_action.triggered.connect(lambda: self.clear_filters(logical_index))

        menu.exec(self.mapToGlobal(QPoint(0, self.height())))

    def apply_filter(self, column, value, checked):
        if column not in self.filters:
            self.filters[column] = set()

        if checked:
            self.filters[column].add(value)
        else:
            self.filters[column].discard(value)

        self.update_table()

    def clear_filters(self, column):
        if column in self.filters:
            del self.filters[column]
        self.update_table()

    def update_table(self):
        for row in range(self.parent.rowCount()):
            show_row = True
            for column, values in self.filters.items():
                if not values:  # If no filters for this column
                    continue
                item = self.parent.item(row, column)
                if item and item.text() not in values:
                    show_row = False
                    break
            self.parent.setRowHidden(row, not show_row)

class StatisticsWindow(QMainWindow):
    def __init__(self, df):
        super().__init__()
        self.setWindowTitle("数据统计")
        self.setGeometry(100, 100, 800, 600)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 创建 matplotlib 画布
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)

        # 创建 pyqtgraph 部件用于柱状图
        self.plot_widget = pg.PlotWidget()
        layout.addWidget(self.plot_widget)

        # 按钮切换图表
        button_layout = QHBoxLayout()
        self.pie_button = QPushButton("性别分布饼图")
        self.bar_button = QPushButton("院系人数柱状图")
        self.pie_button.clicked.connect(lambda: self.plot_pie_chart(df))
        self.bar_button.clicked.connect(lambda: self.plot_bar_chart(df))
        button_layout.addWidget(self.pie_button)
        button_layout.addWidget(self.bar_button)
        layout.addLayout(button_layout)

        # 默认显示饼图
        self.plot_pie_chart(df)

    def plot_pie_chart(self, df):
        self.plot_widget.hide()
        self.canvas.show()

        gender_counts = df["性别"].value_counts()
        labels = gender_counts.index.tolist()
        sizes = gender_counts.values.tolist()

        self.figure.clear()
        ax = self.figure.add_subplot(111)

        def func(pct, allvals):
            absolute = int(np.round(pct / 100. * np.sum(allvals)))
            return f"{pct:.1f}%\n({absolute}人)"

        patches, texts, autotexts = ax.pie(sizes,
                                           labels=labels,
                                           autopct=lambda pct: func(pct, sizes),
                                           startangle=90,
                                           colors=['#FF9999', '#66B2FF'])

        plt.setp(autotexts, size=9, weight="bold")
        plt.setp(texts, size=10)

        ax.set_title("性别分布统计", pad=20, size=12, weight="bold")
        self.canvas.draw()

    def plot_bar_chart(self, df):
        self.canvas.hide()
        self.plot_widget.show()

        self.plot_widget.clear()
        dept_counts = df["院系"].value_counts()
        labels = dept_counts.index.tolist()
        values = dept_counts.values.tolist()

        x = range(len(labels))
        bar = pg.BarGraphItem(x=x, height=values, width=0.6, brush='#00ACC1')
        self.plot_widget.addItem(bar)

        # 添加数值标签
        for i, v in enumerate(values):
            text = pg.TextItem(str(v), anchor=(0.5, 1.0))
            text.setPos(i, v)
            self.plot_widget.addItem(text)

        # 设置坐标轴
        self.plot_widget.getAxis('bottom').setTicks([[(i, label) for i, label in enumerate(labels)]])
        self.plot_widget.setTitle("院系人数分布")
        self.plot_widget.setLabel('left', '人数')

        # 调整显示范围
        self.plot_widget.setRange(xRange=[-0.5, len(labels) - 0.5],
                                  yRange=[0, max(values) * 1.2])


def is_chinese(text):
    if not text:
        return False
    for char in text:
        if not '\u4e00' <= char <= '\u9fff':
            return False
    return True

def validate_name(name):
    if not name or not is_chinese(name) or len(name) < 2 or len(name) > 10:
        return False, "姓名必须是2-10个汉字"
    return True, ""

def validate_gender(gender):
    if gender not in ["男", "女"]:
        return False, "性别必须是男或女"
    return True, ""

def validate_department(department, valid_departments):
    if not department or department not in valid_departments:
        return False, f"院系必须是以下之一：{', '.join(valid_departments)}"
    return True, ""

def validate_major(major):
    if not major or not is_chinese(major) or len(major) < 2 or len(major) > 15:
        return False, "专业必须是2-15个汉字"
    return True, ""

def validate_ethnicity(ethnicity):
    if ethnicity and (not is_chinese(ethnicity) or len(ethnicity) > 6):
        return False, "民族必须是汉字且不超过6个字"
    return True, ""

def validate_province(province):
    if province and (not is_chinese(province) or len(province) > 10):
        return False, "省份必须是汉字且不超过10个字"
    return True, ""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("学生基本信息管理")
        self.setGeometry(100, 100, 1000, 700)

        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 搜索栏
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入姓名搜索")
        self.search_button = QPushButton("搜索")
        self.search_button.clicked.connect(self.search_data)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        layout.addLayout(search_layout)

        # 创建表格
        self.table = QTableWidget()
        header = FilterHeader(self.table)
        self.table.setHorizontalHeader(header)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table)

        # 设置列宽比例
        column_widths = {
            "姓名": 100,
            "性别": 60,
            "民族": 80,
            "院系": 150,
            "专业": 150,
            "省份": 100
        }

        # 应用列宽设置
        def setup_columns(df):
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)
            for col, column_name in enumerate(df.columns):
                width = column_widths.get(column_name, 100)  # 默认宽度100
                self.table.setColumnWidth(col, width)

        # 保存setup_columns方法供后续使用
        self.setup_columns = setup_columns




        # 操作按钮
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("新增")
        self.add_button.setIcon(QIcon("icons/add.png"))
        self.add_button.setIconSize(QSize(16, 16))

        self.edit_button = QPushButton("编辑")
        self.edit_button.setIcon(QIcon("icons/edit.png"))
        self.edit_button.setIconSize(QSize(16, 16))

        self.delete_button = QPushButton("删除")
        self.delete_button.setIcon(QIcon("icons/delete.png"))
        self.delete_button.setIconSize(QSize(16, 16))

        self.stats_button = QPushButton("查看统计")
        self.stats_button.setIcon(QIcon("icons/stats.png"))
        self.stats_button.setIconSize(QSize(16, 16))

        self.search_button.setIcon(QIcon("icons/search.png"))
        self.search_button.setIconSize(QSize(16, 16))



        self.add_button.clicked.connect(self.add_record)
        self.edit_button.clicked.connect(self.edit_record)
        self.delete_button.clicked.connect(self.delete_record)
        self.stats_button.clicked.connect(self.show_statistics)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.stats_button)
        layout.addLayout(button_layout)

        # Apply styles
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QTableWidget {
                background-color: white;
                gridline-color: #d0d0d0;
                border: 1px solid #c0c0c0;
                border-radius: 5px;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 8px;
                border: 1px solid #c0c0c0;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton {
                background-color: #00ACC1;
                color: white;
                border-radius: 5px;
                padding: 8px 15px;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #00838F;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #c0c0c0;
                border-radius: 5px;
                background-color: white;
                font-size: 13px;
            }
            QStatusBar {
                background-color: #e0e0e0;
                color: #333;
                font-size: 12px;
                padding: 5px;
            }
        """)

        # 加载学生数据
        self.file_path = "data/学生数据示例.xlsx"
        self.df = self.load_student_data()
        self.display_data(self.df)
        self.update_status_bar()

    def load_student_data(self):
        try:
            return pd.read_excel(self.file_path)
        except FileNotFoundError:
            QMessageBox.critical(self, "错误", "学生数据文件未找到！")
            return pd.DataFrame()

    def display_data(self, df):
        # 临时关闭排序，提高性能
        self.table.setSortingEnabled(False)
        self.table.setUpdatesEnabled(False)

        try:
            # 设置行数和列数，应用列宽设置
            self.setup_columns(df)
            self.table.setRowCount(len(df))

            # 使用批量更新来提高性能
            items = []
            for i, row in enumerate(df.itertuples(index=False)):
                for j, value in enumerate(row):
                    display_value = "无" if pd.isna(value) else str(value)
                    item = QTableWidgetItem(display_value)
                    items.append((i, j, item))

            # 批量设置表格项
            for i, j, item in items:
                self.table.setItem(i, j, item)

        finally:
            # 恢复更新和排序
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)
            # 强制刷新表格
            self.table.viewport().update()

    def update_status_bar(self):
        record_count = self.table.rowCount()
        self.status_bar.showMessage(f"当前记录数：{record_count}")

    def search_data(self):
        search_term = self.search_input.text().strip().lower()

        if not search_term:
            # 显示所有数据
            self.display_data(self.df)
            return

        try:
            # 改进搜索逻辑，使用更精确的匹配
            mask = self.df.apply(lambda row: any(
                str(cell).strip().lower().find(search_term) >= 0
                for cell in row if pd.notna(cell)  # 只搜索非空值
            ), axis=1)

            filtered_df = self.df[mask]

            if filtered_df.empty:
                QMessageBox.information(self, "提示", "未找到匹配的记录")
                return

            # 显示筛选后的数据
            self.table.setUpdatesEnabled(False)
            try:
                self.display_filtered_data(filtered_df, mask)
            finally:
                self.table.setUpdatesEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"搜索时出错：{str(e)}")

        self.update_status_bar()

    def display_filtered_data(self, filtered_df, mask):
        """显示筛选后的数据，保持原始数据索引"""
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(filtered_df))
        self.table.setColumnCount(len(filtered_df.columns))
        self.table.setHorizontalHeaderLabels(filtered_df.columns)

        # 创建映射表：显示行号 -> 原始数据索引
        self.row_map = dict(enumerate(self.df.index[mask]))
        self.reverse_row_map = {v: k for k, v in self.row_map.items()}

        for display_row, (_, row) in enumerate(filtered_df.iterrows()):
            for col, value in enumerate(row):
                display_value = "无" if pd.isna(value) else str(value)
                item = QTableWidgetItem(display_value)
                self.table.setItem(display_row, col, item)

        self.table.setSortingEnabled(True)

    def add_record(self):
        valid_departments = set(self.df["院系"].unique())
        new_values = {}
        columns_order = ["姓名", "性别", "民族", "院系", "专业", "省份"]

        for column in columns_order:
            while True:
                if column == "姓名":
                    value, ok = QInputDialog.getText(self, "新增记录", "请输入姓名（必填，2-10个汉字）：")
                    if not ok:  # 用户点击取消，直接退出整个方法
                        return
                    valid, msg = validate_name(value)
                    required = True

                elif column == "性别":
                    value, ok = QInputDialog.getText(self, "新增记录", "请输入性别（必填，男/女）：")
                    if not ok:
                        return
                    valid, msg = validate_gender(value)
                    required = True

                elif column == "院系":
                    value, ok = QInputDialog.getText(self, "新增记录",
                                                     f"请输入院系（必填，{', '.join(valid_departments)}）：")
                    if not ok:
                        return
                    valid, msg = validate_department(value, valid_departments)
                    required = True

                elif column == "专业":
                    value, ok = QInputDialog.getText(self, "新增记录", "请输入专业（必填，2-20个汉字）：")
                    if not ok:
                        return
                    valid, msg = validate_major(value)
                    required = True

                elif column == "民族":
                    value, ok = QInputDialog.getText(self, "新增记录", "请输入民族（选填，限10个汉字以内）：")
                    if not ok:
                        return
                    if not value:  # 允许为空
                        new_values[column] = None
                        break
                    valid, msg = validate_ethnicity(value)
                    required = False

                else:  # 省份
                    value, ok = QInputDialog.getText(self, "新增记录", "请输入省份（选填，限10个汉字以内）：")
                    if not ok:
                        return
                    if not value:  # 允许为空
                        new_values[column] = None
                        break
                    valid, msg = validate_province(value)
                    required = False

                if not valid:
                    reply = QMessageBox.warning(self, "警告",
                                                f"{msg}\n是否重新输入？",
                                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if reply == QMessageBox.StandardButton.No:
                        return
                    continue

                new_values[column] = value
                break

            # 检查必填字段
            if required and column not in new_values:
                return

        try:
            # 添加新记录
            new_row = pd.DataFrame([{
                col: new_values.get(col, None) for col in self.df.columns
            }])

            self.df = pd.concat([self.df, new_row], ignore_index=True)
            self.df.to_excel(self.file_path, index=False)
            self.display_data(self.df)
            self.update_status_bar()
            QMessageBox.information(self, "成功", "记录添加成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存数据失败：{str(e)}")

    def get_original_row_index(self, display_row):
        """获取显示行号对应的原始数据索引"""
        return getattr(self, 'row_map', {}).get(display_row, display_row)

    def edit_record(self):
        display_row = self.table.currentRow()
        selected_col = self.table.currentColumn()
        if display_row < 0:
            QMessageBox.warning(self, "警告", "请先选择要编辑的记录！")
            return

        try:
            # 获取原始数据行索引
            selected = self.get_original_row_index(display_row)
            valid_departments = set(self.df["院系"].unique())
            current_record = self.df.iloc[selected]

            # 如果选择了特定列（单击某个单元格）
            if selected_col >= 0:
                column_name = self.df.columns[selected_col]
                current_value = "无" if pd.isna(current_record[column_name]) else str(current_record[column_name])

                # 根据不同列使用相应的验证规则
                while True:
                    if column_name == "姓名":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         "请输入姓名（2-10个汉字）：", text=current_value)
                        if not ok:
                            return
                        valid, msg = validate_name(value)
                    elif column_name == "性别":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         "请输入性别（男/女）：", text=current_value)
                        if not ok:
                            return
                        valid, msg = validate_gender(value)
                    elif column_name == "院系":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         f"请输入院系（{', '.join(valid_departments)}）：",
                                                         text=current_value)
                        if not ok:
                            return
                        valid, msg = validate_department(value, valid_departments)
                    elif column_name == "专业":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         "请输入专业（2-20个汉字）：", text=current_value)
                        if not ok:
                            return
                        valid, msg = validate_major(value)
                    elif column_name == "民族":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         "请输入民族（选填，限10个汉字以内）：", text=current_value)
                        if not ok:
                            return
                        if not value:  # 允许为空
                            break
                        valid, msg = validate_ethnicity(value)
                    elif column_name == "省份":
                        value, ok = QInputDialog.getText(self, "编辑记录",
                                                         "请输入省份（选填，限10个汉字以内）：", text=current_value)
                        if not ok:
                            return
                        if not value:  # 允许为空
                            break
                        valid, msg = validate_province(value)

                    if not valid:
                        QMessageBox.warning(self, "警告", msg)
                        continue
                    break

                try:
                    # 更新单个字段的数据
                    value = None if value == "" else value  # 空字符串转换为 None
                    self.df.at[selected, column_name] = value

                    # 保存到文件
                    QApplication.processEvents()  # 处理待处理的事件
                    self.df.to_excel(self.file_path, index=False)

                    # 只更新修改的单元格
                    display_value = "无" if pd.isna(value) else str(value)
                    self.table.item(display_row, selected_col).setText(display_value)

                    self.update_status_bar()
                    QMessageBox.information(self, "成功", "记录更新成功！")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"保存数据失败：{str(e)}")

            # 如果选择了整行（全部编辑）
            else:
                columns_order = ["姓名", "性别", "民族", "院系", "专业", "省份"]
                new_values = {}

                for column in columns_order:
                    current_value = "无" if pd.isna(current_record[column]) else str(current_record[column])

                    while True:
                        if column == "姓名":
                            value, ok = QInputDialog.getText(self, "编辑记录",
                                                             "请输入姓名（2-10个汉字）：", text=current_value)
                            if not ok:
                                return
                            valid, msg = validate_name(value)
                        elif column == "性别":
                            value, ok = QInputDialog.getText(self, "编辑记录",
                                                             "请输入性别（男/女）：", text=current_value)
                            if not ok:
                                return
                            valid, msg = validate_gender(value)
                        elif column == "院系":
                            value, ok = QInputDialog.getText(self, "编辑记录",
                                                             f"请输入院系（{', '.join(valid_departments)}）：",
                                                             text=current_value)
                            if not ok:
                                return
                            valid, msg = validate_department(value, valid_departments)
                        elif column == "专业":
                            value, ok = QInputDialog.getText(self, "编辑记录",
                                                             "请输入专业（2-20个汉字）：", text=current_value)
                            if not ok:
                                return
                            valid, msg = validate_major(value)
                        elif column in ["民族", "省份"]:
                            prompt = "民族" if column == "民族" else "省份"
                            value, ok = QInputDialog.getText(self, "编辑记录",
                                                             f"请输入{prompt}（选填，限10个汉字以内）：",
                                                             text=current_value)
                            if not ok:
                                return
                            if not value:  # 允许为空
                                value = None
                                break
                            valid, msg = validate_ethnicity(value) if column == "民族" else validate_province(value)

                        if not valid:
                            QMessageBox.warning(self, "警告", msg)
                            continue
                        break

                    new_values[column] = value

                try:
                    # 更新整行数据
                    for column, value in new_values.items():
                        value = None if value == "" else value
                        self.df.at[selected, column] = value

                    # 保存到文件
                    QApplication.processEvents()  # 处理待处理的事件
                    self.df.to_excel(self.file_path, index=False)

                    # 更新表格显示
                    for col, column_name in enumerate(self.df.columns):
                        value = new_values.get(column_name)
                        display_value = "无" if pd.isna(value) else str(value)
                        self.table.item(display_row, col).setText(display_value)

                    self.update_status_bar()
                    QMessageBox.information(self, "成功", "记录更新成功！")

                except Exception as e:
                    QMessageBox.critical(self, "错误", f"保存数据失败：{str(e)}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"编辑记录时出错：{str(e)}")

        finally:
            # 强制刷新表格
            self.table.viewport().update()

    def delete_record(self):
        display_row = self.table.currentRow()
        selected_col = self.table.currentColumn()

        if display_row < 0:
            QMessageBox.warning(self, "警告", "请先选择要删除的记录！")
            return

        # 获取原始数据索引
        original_row = self.get_original_row_index(display_row)

        # 如果选中了具体的单元格
        if selected_col >= 0:
            column_name = self.df.columns[selected_col]
            current_value = self.table.item(display_row, selected_col).text()

            # 检查是否是必填字段
            if column_name in ["姓名", "性别", "院系", "专业"]:
                QMessageBox.warning(self, "警告", f"{column_name}是必填信息，不能删除！")
                return

            # 检查当前值，不区分是否为空，使用统一的提示框样式
            msg_box = QMessageBox()
            msg_box.setWindowTitle("确认操作")
            if current_value == "无":
                msg_box.setText(f"当前{column_name}信息为空，是否删除整行数据？")
                btn_yes = msg_box.addButton("是", QMessageBox.ButtonRole.YesRole)
                btn_no = msg_box.addButton("否", QMessageBox.ButtonRole.NoRole)
                msg_box.setDefaultButton(btn_no)
            else:
                msg_box.setText(f"请选择操作：")
                btn_delete_cell = msg_box.addButton(f"1. 删除此{column_name}信息",
                                                    QMessageBox.ButtonRole.AcceptRole)
                btn_delete_row = msg_box.addButton("2. 删除整行数据",
                                                   QMessageBox.ButtonRole.DestructiveRole)
                msg_box.setDefaultButton(btn_delete_cell)

            msg_box.exec()
            clicked_button = msg_box.clickedButton()

            # 如果点击了关闭按钮（×）
            if clicked_button is None:
                return

            try:
                if current_value == "无":
                    if clicked_button == btn_yes:
                        # 删除整行
                        self.df = self.df.drop(original_row).reset_index(drop=True)
                        self.df.to_excel(self.file_path, index=False)
                        self.display_data(self.df)
                        success_msg = "整行记录已删除"
                    else:
                        return
                else:
                    if clicked_button == btn_delete_cell:
                        # 仅删除单元格内容
                        self.df.at[original_row, column_name] = None
                        self.table.item(display_row, selected_col).setText("无")
                        success_msg = f"{column_name}信息已删除"
                    elif clicked_button == btn_delete_row:
                        # 删除整行
                        self.df = self.df.drop(original_row).reset_index(drop=True)
                        self.df.to_excel(self.file_path, index=False)
                        self.display_data(self.df)
                        success_msg = "整行记录已删除"
                    else:
                        return

                self.update_status_bar()
                QMessageBox.information(self, "成功", success_msg)

            except Exception as e:
                QMessageBox.critical(self, "错误", f"删除失败：{str(e)}")

        else:
            # 删除整行
            msg_box = QMessageBox()
            msg_box.setWindowTitle("确认删除")
            msg_box.setText("确定要删除这条记录吗？\n此操作不可撤销！")
            btn_yes = msg_box.addButton("1. 是", QMessageBox.ButtonRole.YesRole)
            btn_no = msg_box.addButton("2. 否", QMessageBox.ButtonRole.NoRole)
            msg_box.setDefaultButton(btn_no)

            msg_box.exec()
            if msg_box.clickedButton() == btn_yes:
                try:
                    self.df = self.df.drop(original_row).reset_index(drop=True)
                    self.df.to_excel(self.file_path, index=False)
                    self.display_data(self.df)
                    self.update_status_bar()
                    QMessageBox.information(self, "成功", "记录删除成功！")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"删除记录失败：{str(e)}")

    def show_statistics(self):
        self.stats_window = StatisticsWindow(self.df)
        self.stats_window.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_dialog = LoginDialog()
    if login_dialog.exec() == QDialog.DialogCode.Accepted:
        main_window = MainWindow()
        main_window.show()
    sys.exit(app.exec())