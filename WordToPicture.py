import os
import shutil

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QMessageBox
from docx import Document
import re
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Word文档照片提取程序")
        self.setGeometry(500, 500, 700, 600)
        self.open_button = QPushButton("选择Word文档", self)
        self.open_button.setFixedSize(300, 100)
        self.open_button.setStyleSheet("""
            QPushButton {
                color: white; 
                border: none; 
                padding: 15px 32px; 
                text-align: center; 
                text-decoration: none; 
                display: inline-block; 
                font-size: 30px; 
                margin: 4px 2px; 
                transition-duration: 0.4s; 
                cursor: pointer;
            }

            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.open_button.setGeometry(200, 200, 200, 30)
        self.open_button.clicked.connect(self.open_document)

    def open_document(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "选择Word文档", "", "Word 文档 (*.docx)")
        print(file_path)
        if file_path:
            self.extract_images(file_path)

    def extract_images(self, file_path):
        doc = Document(file_path)
        images_folder = os.path.splitext(file_path)[0]
        print(images_folder)
        if os.path.exists(images_folder):
            reply = QMessageBox.question(self, "文件夹已存在", "与Word文档同名的文件夹已存在，是否继续？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return
        os.makedirs(images_folder, exist_ok=True)
        shutil.copy(file_path, images_folder)
        dict_rel = doc.part._rels  # rels其实是个目录
        for rel in dict_rel:
            rel = dict_rel[rel]
            print("rel", rel.target_ref)
            if "image" in rel.target_ref:
                # create_dir(desc_path)
                img_name = re.findall("/(.*)", rel.target_ref)[0]  # windos:/
                img_name = f'{img_name}'
                print(img_name)
                with open(f'{images_folder}/{img_name}', "wb") as f:
                    f.write(rel.target_part.blob)
        print("照片提取完成！")

if __name__ == "__main__":
    app = QApplication([])
    app.setStyleSheet("""
        QPushButton {
            background-color: #4CAF50; 
            color: white; 
        }
    """)
    window = MainWindow()
    window.show()
    app.exec_()
