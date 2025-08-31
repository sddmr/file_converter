import sys
import os
import pandas as pd
import yaml
from PIL import Image
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QFileDialog, QMessageBox, QComboBox, QGroupBox
)
from PySide6.QtGui import QIcon, QFont
from PySide6.QtCore import Qt, QSize

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None


class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Converter")
        self.setGeometry(300, 300, 550, 400)
        self.setStyleSheet("background-color: #121212; color: white;")

        self.file_path = ""
        self.output_dir = ""
        self.target_ext = "Seçenek Yok"
        self.lang = "tr"

        main_layout = QVBoxLayout()

        lang_layout = QHBoxLayout()
        self.tr_btn = QPushButton()
        self.tr_btn.setIcon(QIcon("turkey_flag.png"))
        self.tr_btn.setIconSize(QSize(32, 32))
        self.tr_btn.setFixedSize(40, 40)
        self.tr_btn.setStyleSheet(self.icon_button_style())
        self.tr_btn.clicked.connect(lambda: self.set_language("tr"))

        self.en_btn = QPushButton()
        self.en_btn.setIcon(QIcon("uk_flag.png"))
        self.en_btn.setIconSize(QSize(32, 32))
        self.en_btn.setFixedSize(40, 40)
        self.en_btn.setStyleSheet(self.icon_button_style())
        self.en_btn.clicked.connect(lambda: self.set_language("en"))

        lang_layout.addStretch()
        lang_layout.addWidget(self.tr_btn)
        lang_layout.addWidget(self.en_btn)
        main_layout.addLayout(lang_layout)

        self.file_group = QGroupBox("Dönüştürülecek Dosya")
        self.file_group.setStyleSheet(self.groupbox_style())
        file_layout = QVBoxLayout()
        self.file_label = QLabel("Henüz dosya seçilmedi")
        self.file_label.setFont(QFont("Arial", 10))
        file_btn = QPushButton("Dosya Seç")
        file_btn.setStyleSheet(self.button_style())
        file_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(file_btn)
        self.file_group.setLayout(file_layout)
        main_layout.addWidget(self.file_group)

        self.format_group = QGroupBox("Dönüştürme Formatı")
        self.format_group.setStyleSheet(self.groupbox_style())
        format_layout = QVBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItem("Seçenek Yok")
        format_layout.addWidget(self.format_combo)
        self.format_group.setLayout(format_layout)
        main_layout.addWidget(self.format_group)

        self.output_group = QGroupBox("Kaydedilecek Klasör")
        self.output_group.setStyleSheet(self.groupbox_style())
        output_layout = QVBoxLayout()
        self.output_entry = QLineEdit()
        self.output_entry.setStyleSheet("background-color: #1e1e1e; color: white;")
        output_btn = QPushButton("Klasör Seç")
        output_btn.setStyleSheet(self.button_style())
        output_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(self.output_entry)
        output_layout.addWidget(output_btn)
        self.output_group.setLayout(output_layout)
        main_layout.addWidget(self.output_group)

        convert_btn = QPushButton("DÖNÜŞTÜR VE KAYDET")
        convert_btn.setStyleSheet(self.convert_button_style())
        convert_btn.clicked.connect(self.convert_file)
        main_layout.addWidget(convert_btn)

        self.setLayout(main_layout)

    def button_style(self):
        return """
            QPushButton {
                background-color: #2e2e2e; 
                color: white; 
                border-radius: 5px; 
                padding: 5px 10px;
            }
            QPushButton:hover {
                background-color: #444444;
            }
        """

    def convert_button_style(self):
        return """
            QPushButton {
                background-color: #4CAF50; 
                color: white; 
                font-weight: bold; 
                border-radius: 6px; 
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """

    def groupbox_style(self):
        return """
            QGroupBox {
                border: 1px solid #444; 
                border-radius: 5px; 
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin; 
                left: 10px; 
                padding: 0 3px 0 3px;
            }
        """

    def icon_button_style(self):
        return """
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: #333333;
                border-radius: 5px;
            }
        """

    def set_language(self, lang):
        self.lang = lang
        if lang == "tr":
            self.file_group.setTitle("Dönüştürülecek Dosya")
            self.format_group.setTitle("Dönüştürme Formatı")
            self.output_group.setTitle("Kaydedilecek Klasör")
            self.file_label.setText("Henüz dosya seçilmedi")
        elif lang == "en":
            self.file_group.setTitle("File to Convert")
            self.format_group.setTitle("Target Format")
            self.output_group.setTitle("Output Folder")
            self.file_label.setText("No file selected")

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Dosya seç" if self.lang == "tr" else "Select File",
            "",
            "Metin Dosyaları (*.txt);;CSV Dosyaları (*.csv);;JSON Dosyaları (*.json);;YAML Dosyaları (*.yaml);;"
            "Excel Dosyaları (*.xlsx);;PNG Dosyaları (*.png);;JPG Dosyaları (*.jpg *.jpeg);;WEBP Dosyaları (*.webp);;"
            "Word Dosyaları (*.docx);;Tüm Dosyalar (*.*)"
        )
        if path:
            self.file_path = path
            self.file_label.setText(os.path.basename(path))

            ext = os.path.splitext(path)[1].lower()
            options = []

            if ext in [".txt", ".csv", ".json", ".yaml"]:
                options = ["txt", "csv", "json", "yaml"]
            elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
                options = ["png", "jpg", "webp"]
            elif ext == ".docx":
                options = ["pdf"]
            elif ext == ".xlsx":
                options = ["csv"]

            self.format_combo.clear()
            if options:
                self.format_combo.addItems(options)
                self.target_ext = options[0]
            else:
                self.format_combo.addItem("Seçenek Yok")
                self.target_ext = "Seçenek Yok"

            self.format_combo.currentTextChanged.connect(lambda val: setattr(self, "target_ext", val))

    def browse_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "Kaydedilecek klasörü seç" if self.lang == "tr" else "Select Output Folder")
        if path:
            self.output_dir = path
            self.output_entry.setText(path)

    def convert_file(self):
        input_path = self.file_path
        output_dir = self.output_dir
        target_ext = self.target_ext.lower()

        if not input_path or not os.path.exists(input_path):
            QMessageBox.critical(self, "Hata", "Lütfen geçerli bir dosya seçin." if self.lang == "tr" else "Please select a valid file.")
            return
        if not output_dir or not os.path.exists(output_dir):
            QMessageBox.critical(self, "Hata", "Lütfen geçerli bir çıktı klasörü seçin." if self.lang == "tr" else "Please select a valid output folder.")
            return
        if target_ext == "seçenek yok":
            QMessageBox.critical(self, "Hata", "Dönüştürme formatı seçilemedi." if self.lang == "tr" else "Target format not selected.")
            return

        name, ext = os.path.splitext(os.path.basename(input_path))
        ext = ext.lower()

        try:
            out_path = os.path.join(output_dir, f"{name}.{target_ext}")

            if ext in [".txt", ".csv", ".json", ".yaml", ".xlsx"]:
                df = None
                if ext == ".txt":
                    with open(input_path, "r", encoding="utf-8") as f:
                        lines = [line.strip() for line in f.readlines()]
                    df = pd.DataFrame(lines, columns=["line"])
                elif ext == ".csv":
                    df = pd.read_csv(input_path)
                elif ext == ".json":
                    df = pd.read_json(input_path)
                elif ext == ".yaml":
                    with open(input_path, "r", encoding="utf-8") as f:
                        data = yaml.safe_load(f)
                    df = pd.DataFrame(data)
                elif ext == ".xlsx":
                    df = pd.read_excel(input_path)

                if target_ext == "txt":
                    df.to_csv(out_path, index=False, header=False, encoding='utf-8')
                elif target_ext == "csv":
                    df.to_csv(out_path, index=False, encoding='utf-8')
                elif target_ext == "json":
                    df.to_json(out_path, orient="records", indent=4, force_ascii=False)
                elif target_ext == "yaml":
                    data_dict = df.to_dict('records')
                    with open(out_path, "w", encoding="utf-8") as f:
                        yaml.dump(data_dict, f, allow_unicode=True, default_flow_style=False)

            elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
                im = Image.open(input_path)
                if target_ext == "jpg":
                    if im.mode in ('RGBA', 'LA'):
                        im = im.convert('RGB')
                    im.save(out_path, "JPEG")
                else:
                    im.save(out_path, target_ext.upper())

            elif ext == ".docx":
                if docx2pdf_convert:
                    docx2pdf_convert(input_path, out_path)
                else:
                    QMessageBox.critical(self, "Hata", "`docx2pdf` modülü yüklü değil. 'pip install docx2pdf' yapın.")
                    return

            QMessageBox.information(self, "Başarılı", f"Dönüştürme tamamlandı:\n{out_path}")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dönüştürme sırasında bir hata oluştu:\n{str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec())
