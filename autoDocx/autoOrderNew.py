from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QTabWidget, QTextEdit, QHBoxLayout
from PyQt5.QtCore import Qt
import sys
from docx import Document
from datetime import datetime

class PurchaseRequestApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("구매요구서 자동화")
        self.setGeometry(100, 100, 600, 400)

        self.tab_widget = QTabWidget(self)
        self.setup_tab1()
        self.setup_tab2()
        self.setup_tab3()

        layout = QVBoxLayout()
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def setup_tab1(self):
        tab1 = QWidget()
        self.tab_widget.addTab(tab1, "수동입력")

        label = QLabel("구매요구서 양식:")
        self.file_path = QLineEdit(self)
        self.file_path.setReadOnly(True)
        button = QPushButton("Browse", self)
        button.clicked.connect(self.upload_form)
                       
        institution_label = QLabel("기관명:")
        self.institution_entry = QLineEdit(self)
        self.institution_entry.setFixedWidth(150)

        institution_code_label = QLabel("기관코드:")
        self.institution_code_entry = QLineEdit(self)
        self.institution_code_entry.setFixedWidth(150)
        
        isotope_fname_label = QLabel("핵종명[앞]:")
        self.isotope_fname_entry = QLineEdit(self)
        self.isotope_fname_entry.setFixedWidth(150)
        
        quantity_label = QLabel("핵종수량:")
        self.quantity_entry = QLineEdit(self)
        self.quantity_entry.setFixedWidth(150)

        isotope_ename_label = QLabel("핵종명[뒤]:")
        self.isotope_ename_entry = QLineEdit(self)
        self.isotope_ename_entry.setFixedWidth(150)
        
        model_name_label = QLabel("모델명:")
        self.model_name_entry = QLineEdit(self)
        self.model_name_entry.setFixedWidth(150)

        manpower_labels = [f"투입인력 : {i:02d}:{j:02d}" for i in range(8, 19) for j in range(0, 60, 30) if not (i == 17 and j == 30)]
        self.manpower_entries = [QLineEdit(self) for _ in range(19)]
        
        for entry in self.manpower_entries:
            entry.setFixedWidth(50)  # 크기를 50으로 고정
            
        
        purchase_date_label = QLabel("구입 예정일:")
        self.purchase_date_entry = QLineEdit(self)
        self.purchase_date_entry.setFixedWidth(150)

        generate_button = QPushButton("파일생성하기", self)
        generate_button.clicked.connect(self.generate_word_file)

        tab1_layout = QVBoxLayout()
        
        hbox_file_path = QHBoxLayout()
        hbox_file_path.addWidget(label)
        hbox_file_path.addWidget(self.file_path)
        hbox_file_path.addWidget(button)        
        tab1_layout.addLayout(hbox_file_path)
        
        hbox_file_path2 = QHBoxLayout()
        hbox_file_path2.addWidget(institution_label)
        hbox_file_path2.addWidget(self.institution_entry)
        tab1_layout.addLayout(hbox_file_path2)
        
        hbox_file_path3 = QHBoxLayout()
        hbox_file_path3.addWidget(institution_code_label)
        hbox_file_path3.addWidget(self.institution_code_entry)
        tab1_layout.addLayout(hbox_file_path3)
        
        hbox_file_path4 = QHBoxLayout()
        hbox_file_path4.addWidget(isotope_fname_label)
        hbox_file_path4.addWidget(self.isotope_fname_entry)
        tab1_layout.addLayout(hbox_file_path4)
        
        hbox_file_path5 = QHBoxLayout()
        hbox_file_path5.addWidget(quantity_label)
        hbox_file_path5.addWidget(self.quantity_entry)
        tab1_layout.addLayout(hbox_file_path5)
        
        hbox_file_path6 = QHBoxLayout()
        hbox_file_path6.addWidget(isotope_ename_label)
        hbox_file_path6.addWidget(self.isotope_ename_entry)
        tab1_layout.addLayout(hbox_file_path6)
        
        hbox_file_path7 = QHBoxLayout()
        hbox_file_path7.addWidget(model_name_label)
        hbox_file_path7.addWidget(self.model_name_entry)
        tab1_layout.addLayout(hbox_file_path7)
        
        for i in range(19):
            hbox_file_pathMan = QHBoxLayout()
            hbox_file_pathMan.addWidget(QLabel(manpower_labels[i]))
            hbox_file_pathMan.addWidget(self.manpower_entries[i])
            hbox_file_pathMan.addStretch(1)
            tab1_layout.addLayout(hbox_file_pathMan)

        hbox_file_path8 = QHBoxLayout()
        hbox_file_path8.addWidget(purchase_date_label)
        hbox_file_path8.addWidget(self.purchase_date_entry)
        tab1_layout.addLayout(hbox_file_path8)

        tab1_layout.addWidget(generate_button)

        tab1.setLayout(tab1_layout)

    def upload_form(self):
        file_dialog = QFileDialog()
        file_path = file_dialog.getOpenFileName(self, "Select Word file", "", "Word files (*.docx)")[0]
        if file_path:
            self.file_path.setText(file_path)

    def generate_word_file(self):
        template_path = self.file_path.text()
        if template_path:
            document = Document(template_path)
            placeholders = {
                "instNm": self.institution_entry.text(),
                "instCd": self.institution_code_entry.text(),
                "isotopeFront": self.isotope_fname_entry.text(),
                "Quantity": self.quantity_entry.text(),
                "isotopeEnd": self.isotope_ename_entry.text(),
                "model": self.model_name_entry.text(),
                "manPw1": self.manpower_entries[0].text(),
                "manPw2": self.manpower_entries[1].text(),
                "manPw3": self.manpower_entries[2].text(),
                "manPw4": self.manpower_entries[3].text(),
                "manPw5": self.manpower_entries[4].text(),
                "manPw6": self.manpower_entries[5].text(),
                "manPw7": self.manpower_entries[6].text(),
                "manPw8": self.manpower_entries[7].text(),
                "manPw9": self.manpower_entries[8].text(),
                "manPw10": self.manpower_entries[9].text(),
                "manPw11": self.manpower_entries[10].text(),
                "manPw12": self.manpower_entries[11].text(),
                "manPw13": self.manpower_entries[12].text(),
                "manPw14": self.manpower_entries[13].text(),
                "manPw15": self.manpower_entries[14].text(),
                "manPw16": self.manpower_entries[15].text(),
                "manPw17": self.manpower_entries[16].text(),
                "manPw18": self.manpower_entries[17].text(),
                "purchDt": self.purchase_date_entry.text()
            }

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in placeholders.items():
                            self.replace_text_in_cell(cell, key, value)

            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            new_file_path = f"구매요구서_{timestamp}.docx"
            document.save(new_file_path)
            print(f"파일이 저장되었습니다: {new_file_path}")

    def replace_text_in_cell(self, cell, key, value):
        cell.text = cell.text.replace(f"{key}", value)

    def setup_tab2(self):
        tab2 = QWidget()
        self.tab_widget.addTab(tab2, "데이터 등록")
        # Implement Tab 2 UI and functionality here

    def setup_tab3(self):
        tab3 = QWidget()
        self.tab_widget.addTab(tab3, "검색 입력")
        # Implement Tab 3 UI and functionality here

def main():
    app = QApplication(sys.argv)
    window = PurchaseRequestApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
