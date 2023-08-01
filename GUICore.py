from shutil import copy as shucopy

from GUI import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox
import os
import pandas as pd

log_list = []


class MainWindow(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.yesFirstLine.setChecked(True)

        self.openExclePathButton.clicked.connect(self.openExcPath)

        self.openToPathButton.clicked.connect(self.openToPath)

        self.openOrgDirPathButton.clicked.connect(self.openOrgDirPath)

        self.runButton.clicked.connect(self.copyRun)

        self.export_log_button.clicked.connect(self.export_log)

    def export_log(self):
        dir = QFileDialog()
        dir.setFileMode(QFileDialog.DirectoryOnly)
        dir.setDirectory('c:\\users\\')
        try:
            if dir.exec_():
                save_path = os.path.join(dir.selectedFiles()[0], "logs.txt")
        except:
            print(1)
            return 0
        with open(save_path, 'w') as f:
            count = 1
            for log in log_list:
                f.write(log + '\n')
                if len(log_list) <= 100:
                    print("导出中...." + str(round(count / len(log_list) * 100, 2)) + "%/100%")
                else:
                    if (count % 5 == 0):
                        print("导出中...." + str(round(count / len(log_list) * 100,2)) + "%/100%")
                count += 1
        QMessageBox.warning(None, "提示", "导出完成！", QMessageBox.Ok)


    def copyRun(self):
        log_list.clear()
        gl_fileName = self.excelPathLine.text()  #
        # gl_lineName = self.columnNameLine.text()
        gl_toDirPath = self.toDirPathLine.text()
        gl_orgDirPath = self.orgDirLine.text()
        gl_columnNumber = self.columnLine.text()  #
        if (self.yesFirstLine.isChecked()):
            gl_getFirstLine = 0  # 取
        elif (self.noFirstLine.isChecked()):
            gl_getFirstLine = 1  # 不取

        # 1. 读取Excel指定列到all_files列表
        file_path = gl_fileName

        skip_first_row = gl_getFirstLine

        try:
            df = pd.read_excel(file_path, header=None)
        except FileNotFoundError:
            QMessageBox.critical(None, "警告", "未找到文件，请正确填写表格文件所在路径", QMessageBox.Ok)
            return 0
        except ValueError:
            QMessageBox.critical(None, "警告", "请填写表格文件所在路径", QMessageBox.Ok)
            return 0

        try:
            column_num = int(gl_columnNumber) - 1
        except ValueError:
            QMessageBox.critical(None, "警告", "请填写表格中文件名所在列", QMessageBox.Ok)
            return 0

        if skip_first_row == 1:
            all_files = df.iloc[1:, column_num].tolist()
        else:
            all_files = df.iloc[:, column_num].tolist()
        print(all_files)
        # 2. 复制原文件夹中的文件到目标文件夹
        src_folder = gl_orgDirPath

        try:
            os.listdir(src_folder)
        except:
            QMessageBox.critical(None, "警告", "原文件路径不存在", QMessageBox.Ok)
            return 0

        dst_folder = gl_toDirPath

        try:
            os.listdir(dst_folder)
        except:
            QMessageBox.critical(None, "警告", "复制目标路径不存在", QMessageBox.Ok)
            return 0

        lens = len(all_files)
        logs = 1
        for file in all_files:

            src_path = os.path.join(src_folder, file)

            if os.path.exists(src_path):
                dst_path = os.path.join(dst_folder, file)

                shucopy(src_path, dst_path)
                temp = "(" + str(logs) + "/" + str(lens) + ")" + file + "  移动成功！"
                print(temp)  # 加个log功能
                log_list.append(temp)
            else:
                temp = "(" + str(logs) + "/" + str(lens) + ")" + file + "  移动失败，请检查原文件路径是否正确。"
                print(temp)
                log_list.append(temp)
            logs += 1

        self.export_log_button.setEnabled(True)
        QMessageBox.warning(None, "提示", "运行完成！", QMessageBox.Ok)


    def openToPath(self):
        dir = QFileDialog()
        dir.setFileMode(QFileDialog.DirectoryOnly)
        dir.setDirectory('c:\\users\\')

        if dir.exec_():
            self.toDirPathLine.setText(dir.selectedFiles()[0])

    def openOrgDirPath(self):
        dir = QFileDialog()
        dir.setFileMode(QFileDialog.DirectoryOnly)
        dir.setDirectory('c:\\users\\')

        if dir.exec_():
            self.orgDirLine.setText(dir.selectedFiles()[0])

    def openExcPath(self):
        dir = QFileDialog()
        dir.setNameFilter('表格文件(*.xlsx *.xls)')
        dir.setFileMode(QFileDialog.ExistingFile)
        dir.setDirectory('c:\\users\\')

        if dir.exec_():
            self.excelPathLine.setText(dir.selectedFiles()[0])
