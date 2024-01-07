import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import sqlite3
from docx import Document
from datetime import datetime
from copy import deepcopy

class PurchaseRequestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Request App")

        self.tab_control = ttk.Notebook(root)

        # Tab 1: Manual Input
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab1, text="수동입력")
        self.setup_tab1()

        # Tab 2: Data Registration
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab2, text="데이터 등록")
        self.setup_tab2()

        # Tab 3: Search and Update
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab3, text="검색 입력")
        self.setup_tab3()

        self.tab_control.pack(expand=1, fill='both')

    def setup_tab1(self):
       label = ttk.Label(self.tab1, text="구매요구서 양식:")
       label.grid(column=0, row=0)

       self.file_path = tk.StringVar()

       entry = ttk.Entry(self.tab1, textvariable=self.file_path, state="readonly")
       entry.grid(column=1, row=0)

       button = ttk.Button(self.tab1, text="Browse", command=self.upload_form)
       button.grid(column=2, row=0)

        # 사용자 입력 부분 
       ttk.Label(self.tab1, text="기관명:").grid(column=0, row=1)
       self.institution_entry = ttk.Entry(self.tab1)
       self.institution_entry.grid(column=1, row=1)

       ttk.Label(self.tab1, text="기관코드:").grid(column=0, row=2)
       self.institution_code_entry = ttk.Entry(self.tab1)
       self.institution_code_entry.grid(column=1, row=2)

       ttk.Label(self.tab1, text="핵종명[앞]:").grid(column=0, row=3)
       self.isotope_fname_entry = ttk.Entry(self.tab1)
       self.isotope_fname_entry.grid(column=1, row=3)

       ttk.Label(self.tab1, text="핵종수량:").grid(column=0, row=4)
       self.quantity_entry = ttk.Entry(self.tab1)
       self.quantity_entry.grid(column=1, row=4)

       
       ttk.Label(self.tab1, text="핵종명[뒤]:").grid(column=0, row=5)
       self.isotope_ename_entry = ttk.Entry(self.tab1)
       self.isotope_ename_entry.grid(column=1, row=5)

       ttk.Label(self.tab1, text="모델명:").grid(column=0, row=6)
       self.model_name_entry = ttk.Entry(self.tab1)
       self.model_name_entry.grid(column=1, row=6)

       ttk.Label(self.tab1, text="투입인력 08:30 :").grid(column=0, row=7)
       self.manpower_entry1 = ttk.Entry(self.tab1)
       self.manpower_entry1.grid(column=1, row=7)

       ttk.Label(self.tab1, text="투입인력 09:00 :").grid(column=0, row=8)
       self.manpower_entry2 = ttk.Entry(self.tab1)
       self.manpower_entry2.grid(column=1, row=8)

       ttk.Label(self.tab1, text="투입인력 09:30 :").grid(column=0, row=9)
       self.manpower_entry3 = ttk.Entry(self.tab1)
       self.manpower_entry3.grid(column=1, row=9)
       
       ttk.Label(self.tab1, text="투입인력 10:00 :").grid(column=0, row=10)
       self.manpower_entry4 = ttk.Entry(self.tab1)
       self.manpower_entry4.grid(column=1, row=10)

       ttk.Label(self.tab1, text="투입인력 10:30 :").grid(column=0, row=11)
       self.manpower_entry5 = ttk.Entry(self.tab1)
       self.manpower_entry5.grid(column=1, row=11)

       ttk.Label(self.tab1, text="투입인력 11:00 :").grid(column=0, row=12)
       self.manpower_entry6 = ttk.Entry(self.tab1)
       self.manpower_entry6.grid(column=1, row=12)

       ttk.Label(self.tab1, text="투입인력 11:30 :").grid(column=0, row=13)
       self.manpower_entry7 = ttk.Entry(self.tab1)
       self.manpower_entry7.grid(column=1, row=13)

       ttk.Label(self.tab1, text="투입인력 12:00 :").grid(column=0, row=14)
       self.manpower_entry8 = ttk.Entry(self.tab1)
       self.manpower_entry8.grid(column=1, row=14)

       ttk.Label(self.tab1, text="투입인력 12:30 :").grid(column=0, row=15)
       self.manpower_entry9 = ttk.Entry(self.tab1)
       self.manpower_entry9.grid(column=1, row=15)

       ttk.Label(self.tab1, text="투입인력 13:00 :").grid(column=0, row=16)
       self.manpower_entry10 = ttk.Entry(self.tab1)
       self.manpower_entry10.grid(column=1, row=16)

       ttk.Label(self.tab1, text="투입인력 13:30 :").grid(column=0, row=17)
       self.manpower_entry11 = ttk.Entry(self.tab1)
       self.manpower_entry11.grid(column=1, row=17)

       ttk.Label(self.tab1, text="투입인력 14:00 :").grid(column=0, row=18)
       self.manpower_entry12 = ttk.Entry(self.tab1)
       self.manpower_entry12.grid(column=1, row=18)

       ttk.Label(self.tab1, text="투입인력 14:30 :").grid(column=0, row=19)
       self.manpower_entry13 = ttk.Entry(self.tab1)
       self.manpower_entry13.grid(column=1, row=19)

       ttk.Label(self.tab1, text="투입인력 15:00 :").grid(column=0, row=20)
       self.manpower_entry14 = ttk.Entry(self.tab1)
       self.manpower_entry14.grid(column=1, row=20)

       ttk.Label(self.tab1, text="투입인력 15:30 :").grid(column=0, row=21)
       self.manpower_entry15 = ttk.Entry(self.tab1)
       self.manpower_entry15.grid(column=1, row=21)

       ttk.Label(self.tab1, text="투입인력 16:00 :").grid(column=0, row=22)
       self.manpower_entry16 = ttk.Entry(self.tab1)
       self.manpower_entry16.grid(column=1, row=22)

       ttk.Label(self.tab1, text="투입인력 16:30 :").grid(column=0, row=23)
       self.manpower_entry17 = ttk.Entry(self.tab1)
       self.manpower_entry17.grid(column=1, row=23) 
       
       ttk.Label(self.tab1, text="투입인력 17:00 :").grid(column=0, row=24)
       self.manpower_entry18 = ttk.Entry(self.tab1)
       self.manpower_entry18.grid(column=1, row=24) 
       
       ttk.Label(self.tab1, text="구입 예정일:").grid(column=0, row=25)
       self.purchase_date_entry = ttk.Entry(self.tab1)
       self.purchase_date_entry.grid(column=1, row=25)

       # 워드파일 생성
       generate_button = ttk.Button(self.tab1, text="파일생성하기", command=self.generate_word_file)
       generate_button.grid(column=1, row=26)

    def upload_form(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            self.file_path.set(file_path)

    def generate_word_file(self):
        template_path = self.file_path.get()
        if template_path:
            # 양식 파일 
            document = Document(template_path)

            # 파일내용 변환 
            placeholders = {
                "instNm": self.institution_entry.get(),
                "instCd": self.institution_code_entry.get(),
                "isotopeFront": self.isotope_fname_entry.get(),
                "Quantity": self.quantity_entry.get (),
                "isotopeEnd": self.isotope_ename_entry.get(),
                "model": self.model_name_entry.get(),
                "manPw1": self.manpower_entry1.get(),
                "manPw2": self.manpower_entry2.get(),
                "manPw3": self.manpower_entry3.get(),
                "manPw4": self.manpower_entry4.get(),
                "manPw5": self.manpower_entry5.get(),
                "manPw6": self.manpower_entry6.get(),
                "manPw7": self.manpower_entry7.get(),
                "manPw8": self.manpower_entry8.get(),
                "manPw9": self.manpower_entry9.get(),
                "manPw10": self.manpower_entry10.get(),
                "manPw11": self.manpower_entry11.get(),
                "manPw12": self.manpower_entry12.get(),
                "manPw13": self.manpower_entry13.get(),
                "manPw14": self.manpower_entry14.get(),
                "manPw15": self.manpower_entry15.get(),
                "manPw16": self.manpower_entry16.get(),
                "manPw17": self.manpower_entry17.get(),
                "manPw18": self.manpower_entry18.get(),
                "purchDt": self.purchase_date_entry.get()
            }
            # 각 표에서 검색 단어를 치환합니다.
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in placeholders.items():
                            # cell.text = cell.text.replace(f"{key}", value)
                            self.replace_text_in_cell(cell, key, value)
            
            
            # for paragraph in document.paragraphs:
            #     for key, value in placeholders.items():
            #         print(f"입력값 ={key}  변환값 = {value}")
            #         paragraph.text = paragraph.text.replace(f"{key}", value)

            # 변경된 파일 저장 
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S") 
            new_file_path = f"구매요구서_{timestamp}.docx"
            document.save(new_file_path)
            print(f"Word document modified and saved as: {new_file_path}")

    def replace_text_in_cell(self, cell, key, value):
      # 셀 내용을 치환
        cell.text = cell.text.replace(f"{key}", value)
      
      
      
        # # 문단을 복제하여 스타일과 형식을 그대로 유지
        # new_cell = deepcopy(cell)
        
        # # 기존 문단 비우기
        # for paragraph in cell.paragraphs:
        #     for run in paragraph.runs:
        #         run.clear()
                
        # # 새로운 내용으로 채우기
        # new_cell.text = new_cell.text.replace(f"{key}", value)

        #  # 부모인 행에 대해 새로운 셀을 추가
        # row = cell._element.getparent()
        # if row is not None:
        #     row.append(new_cell._element)
        # else:
        #     # 부모 행이 없는 경우, 로그 등을 통해 디버깅 정보를 확인할 수 있습니다.
        #     print("Warning: Parent row not found.")

        # # 기존 셀 제거
        # if cell._element.getparent() is not None:
        #     cell._element.getparent().remove(cell._element)
        # else:
        #     # 부모 셀이 없는 경우, 로그 등을 통해 디버깅 정보를 확인할 수 있습니다.
        #     print("Warning: Parent cell not found.")
        pass

    def setup_tab2(self):
        # Implement Tab 2 UI and functionality here
        pass

    def setup_tab3(self):
        # Implement Tab 3 UI and functionality here
        pass

def main():
    root = tk.Tk()
    app = PurchaseRequestApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()