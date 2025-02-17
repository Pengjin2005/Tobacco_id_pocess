from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog

import pandas as pd
from openpyxl import load_workbook
import sys
import os
import datetime
import json

with open("C:/nova/tobacco_process_copy/config.ini", "r", encoding="utf-8") as f:
    config = json.load(f)


class Log:
    def __init__(self):
        self.log_file = config["log_file"]
        with open(self.log_file, "w") as f:
            f.write("Start Log at " + str(datetime.datetime.now()) + "\n")

    def info(self, msg: str):
        with open(self.log_file, "a") as f:
            f.write("INFO: " + msg + " " + str(datetime.datetime.now()) + "\n")

    def warning(self, msg: str):
        with open(self.log_file, "a") as f:
            f.write("WARNING: " + msg + " " + str(datetime.datetime.now()) + "\n")

    def error(self, msg: str):
        with open(self.log_file, "a") as f:
            f.write("ERROR: " + msg + " " + str(datetime.datetime.now()) + "\n")


class TobaccoProcess:
    def __init__(self, path):
        self.log = Log()
        self.path = path
        # self.path = path.replace("\\", "\\\\")
        try:
            self.data = pd.read_csv(self.path, encoding="gb2312")
            self.n_data = self.data.copy()
            self.mapping_dict = dict()
            self.log.info("Read raw data from " + self.path + "successfully")
        except Exception as e:
            self.log.error(
                "Read raw data from " + self.path + "failed: " + str(e) + "\n"
            )

        try:
            mapping_data = pd.read_excel(config["mapping_file"])
            for i in range(len(mapping_data)):
                self.mapping_dict[mapping_data["大条码"][i]] = mapping_data["小条码"][i]
            self.log.info("Read mapping data from mapping.xlsx successfully")
        except Exception as e:
            self.log.error(
                "Read mapping data from mapping.xlsx failed: " + str(e) + "\n"
            )

    def mapping(self):
        try:
            for i in range(len(self.data)):
                if self.data["条码"][i] in self.mapping_dict:
                    self.n_data.loc[i, "条码"] = str(
                        self.mapping_dict[self.data["条码"][i]]
                    )

            self.n_data["批发价"] = self.n_data["批发价"].apply(lambda x: x / 10)
            self.n_data["零售价"] = self.n_data["零售价"].apply(lambda x: x / 10)
            self.n_data["需求量"] = self.n_data["需求量"].apply(lambda x: x * 10)
            self.n_data["订购量"] = self.n_data["订购量"].apply(lambda x: x * 10)

            self.log.info("Mapping data successfully")
        except Exception as e:
            self.log.error("Mapping data failed: " + str(e) + "\n")

    def save_data(self):
        try:
            self.n_data = self.n_data.drop(columns=["厂家名称"])
            self.n_data = self.n_data.astype(str)
            self.n_data.to_excel(config["result_file"], index=False)

            workbook = load_workbook(config["result_file"])
            worksheet = workbook.active

            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # 获取列字母
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column].width = adjusted_width

            workbook.save(config["result_file"])
            self.log.info(f"Save data to {config['result_file']} successfully")
        except Exception as e:
            self.log.error(
                f"Save data to {config['result_file']} failed: " + str(e) + "\n"
            )

        try:
            os.startfile(f"{config['result_file']}")
            self.log.info(f"Open {config['result_file']} successfully")
        except Exception as e:
            self.log.error(f"Open {config['result_file']} failed: " + str(e) + "\n")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("卷烟条码编辑")
        self.setGeometry(500, 200, 500, 200)
        self.button = QPushButton("请选择下载的文件", self)
        self.button.setGeometry(50, 50, 400, 100)
        self.button.clicked.connect(self.select_file)

    def select_file(self):

        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("卷烟条码编辑")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)

        file_path, _ = file_dialog.getOpenFileName(
            None, "选择文件", config["default_path"]
        )
        try:
            print("file path", file_path)
            tobacco = TobaccoProcess(file_path)
            tobacco.mapping()
            tobacco.save_data()
        except:
            pass


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
