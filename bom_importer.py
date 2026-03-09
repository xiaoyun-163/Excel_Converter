import sys
import win32com.client
import pythoncom
import logging
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QComboBox, QLabel, QMessageBox, QProgressBar)
from PyQt6.QtCore import Qt

# 设置日志配置
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class BOMImporterTool(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CATIA 导入工具")
        self.resize(350, 250)
        self.layout = QVBoxLayout()
        self.next_x = 20.0
        self.next_y = 20.0
        self.import_count = 0

        if not self.check_catia():
            sys.exit()

        self.catia_selector = QComboBox()
        self.layout.addWidget(QLabel("1. 选择目标 CATIA 图纸:"))
        self.layout.addWidget(self.catia_selector)

        self.btn_get_range = QPushButton("2. 检测 Excel 选区范围")
        self.btn_get_range.clicked.connect(self.detect_range)
        self.layout.addWidget(self.btn_get_range)

        self.range_label = QLabel("请先在 Excel 中选定区域并检测...")
        self.layout.addWidget(self.range_label)

        self.progress = QProgressBar()
        self.layout.addWidget(QLabel("导入进度:"))
        self.layout.addWidget(self.progress)

        self.btn_import = QPushButton("3. 确认无误，开始导入")
        self.btn_import.clicked.connect(self.import_data)
        self.btn_import.setEnabled(False)
        self.layout.addWidget(self.btn_import)

        self.setLayout(self.layout)
        self.refresh_catia_list()
        self.selected_rows = 0
        self.selected_cols = 0

    def detect_range(self):
        try:
            excel_app = win32com.client.GetActiveObject("Excel.Application")
            selection = excel_app.Selection
            self.selected_rows = selection.Rows.Count
            self.selected_cols = selection.Columns.Count
            self.range_label.setText(f"检测到区域: {self.selected_rows} 行 x {self.selected_cols} 列")
            self.btn_import.setEnabled(True)
        except Exception:
            QMessageBox.critical(self, "错误", "未能连接 Excel，请确保已打开 Excel 并选定区域！")

    def check_catia(self):
        try:
            win32com.client.GetActiveObject("CATIA.Application")
            return True
        except:
            QMessageBox.critical(None, "错误", "未检测到 CATIA 程序，请先启动 CATIA 后再运行本程序！")
            return False

    def refresh_catia_list(self):
        try:
            catia = win32com.client.GetActiveObject("CATIA.Application")
            for i in range(1, catia.Documents.Count + 1):
                doc = catia.Documents.Item(i)
                if doc.Name.endswith(".CATDrawing"):
                    self.catia_selector.addItem(doc.Name)
        except:
            pass

    def import_data(self):
        try:
            # 1. 获取 Excel 数据对象
            excel_app = win32com.client.GetActiveObject("Excel.Application")
            selection = excel_app.Selection
            
            # 2. 获取 CATIA 对象
            catia = win32com.client.GetActiveObject("CATIA.Application")
            target_doc = catia.Documents.Item(self.catia_selector.currentText())
            target_doc.Activate()
            view = target_doc.Sheets.ActiveSheet.Views.ActiveView
            
            # 定义基础尺寸
            r_int = int(self.selected_rows)
            c_int = int(self.selected_cols)
            h_val = 12.0
            w_val = 25.0
            offset_x = 30.0 
            offset_y = 20.0
            
            # 计算本次的放置坐标
            current_x = 20.0 + (self.import_count * offset_x)
            current_y = 20.0 + (self.import_count * offset_y)
            
            # 3. 创建表格
            try:
                table = view.Tables.Add(current_x, current_y, r_int, c_int, h_val, w_val)
            except Exception:
                table = view.Tables.Add(current_x, current_y, r_int, c_int)
                for i in range(1, c_int + 1): table.Columns.Item(i).Width = w_val
                for j in range(1, r_int + 1): table.Rows.Item(j).Height = h_val
            
            # 4. 填充数据 (包含合并单元格逻辑)
            total_cells = self.selected_rows * self.selected_cols
            for r in range(self.selected_rows):
                for c in range(self.selected_cols):
                    excel_cell = selection.Cells(r + 1, c + 1)
                    
                    # --- 改进逻辑开始 ---
                    # 直接获取 .Text 属性。如果它包含 ###，这是因为 Excel 界面显示限制，
                    # 我们通过调用 excel_cell.EntireColumn.AutoFit() 瞬间自动调整列宽，
                    # 这样 .Text 就会更新为正确显示的字符，然后再读取。
                    if "#" in str(excel_cell.Text):
                        excel_cell.EntireColumn.AutoFit()
                    
                    text_to_fill = str(excel_cell.Text)
                    # --- 改进逻辑结束 ---
                    
                    # 处理合并单元格与符号替换
                    if excel_cell.MergeCells:
                        if excel_cell.Address != excel_cell.MergeArea.Cells(1, 1).Address:
                            text_to_fill = ""
                    
                    table.GetCellObject(r + 1, c + 1).Text = text_to_fill.replace('*', '×')
                    
                    # 更新进度条
                    idx = (r * self.selected_cols) + (c + 1)
                    if idx % 5 == 0 or idx == total_cells:
                        self.progress.setValue(int((idx / total_cells) * 100))
                        QApplication.processEvents()
            
            # 5. 成功提示
            self.import_count += 1
            QMessageBox.information(self, "成功", f"表格已导入！第 {self.import_count} 次偏移。")
            self.progress.setValue(0)
            
        except Exception as e:
            # 2. 写入日志
            logging.error(f"导入过程中发生错误: {str(e)}", exc_info=True)
            self.progress.setValue(0)
            QMessageBox.critical(self, "错误", f"导入失败: {str(e)} \n(详细信息已记录在 error_log.txt)")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    tool = BOMImporterTool()
    tool.show()
    sys.exit(app.exec())