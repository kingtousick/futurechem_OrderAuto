import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import sqlite3
from docx import Document
from datetime import datetime
from copy import deepcopy
import os
import sys 


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class PurchaseRequestApp:
    def __init__(self, root):
        
        self.root = root
        self.root.title("구매요구서 자동화")
        self.manpower_entries = [] 

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
       ttk.Label(self.tab1, text="기관명:").grid(column=0, row=1, sticky=tk.E)
       self.institution_entry = ttk.Entry(self.tab1)
       self.institution_entry.grid(column=1, row=1)
       
       ttk.Label(self.tab1, text="기관코드:").grid(column=0, row=2, sticky=tk.E)
       self.institution_code_entry = ttk.Entry(self.tab1)
       self.institution_code_entry.grid(column=1, row=2)

       ttk.Label(self.tab1, text="핵종명[앞]:").grid(column=0, row=3, sticky=tk.E)
       self.isotope_fname_entry = ttk.Entry(self.tab1)
       self.isotope_fname_entry.grid(column=1, row=3)

       ttk.Label(self.tab1, text="핵종수량:").grid(column=0, row=4, sticky=tk.E)
       self.quantity_entry = ttk.Entry(self.tab1)
       self.quantity_entry.grid(column=1, row=4)

       
       ttk.Label(self.tab1, text="핵종명[뒤]:").grid(column=0, row=5, sticky=tk.E)
       self.isotope_ename_entry = ttk.Entry(self.tab1)
       self.isotope_ename_entry.grid(column=1, row=5)

       ttk.Label(self.tab1, text="모델명:").grid(column=0, row=6, sticky=tk.E)
       self.model_name_entry = ttk.Entry(self.tab1)
       self.model_name_entry.grid(column=1, row=6)
       
     
       manpower_labels = [f"투입인력 : {i:02d}:{j:02d}" for i in range(8, 18) for j in range(0, 60, 30) if not ((i == 8 and j == 0) or (i == 17 and j == 30))]
       for i in range(len(manpower_labels)):
            ttk.Label(self.tab1, text=manpower_labels[i]).grid(column=0, row=7+i)
            entry = ttk.Entry(self.tab1)
            entry.grid(column=1, row=7+i)
            self.manpower_entries.append(entry)
       
       ttk.Label(self.tab1, text="구입 예정일:").grid(column=0, row=25, sticky=tk.E)
       self.purchase_date_entry = ttk.Entry(self.tab1)
       self.purchase_date_entry.grid(column=1, row=25)

       # 워드파일 생성
       generate_button = ttk.Button(self.tab1, text="파일생성하기", command=self.generate_word_file)
       generate_button.grid(column=1, row=27)

    def upload_form(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            self.file_path.set(file_path)
                                              
    def generate_word_file(self):
        template_path = self.file_path.get()
        if template_path:
            # 양식 파일 
            original_document = Document(template_path)
        
            # 문서를 복사하여 기존 양식 유지
            document = deepcopy(original_document)

            # 파일내용 변환 
            placeholders = {
                "instNm": self.institution_entry.get(),
                "instCd": self.institution_code_entry.get(),
                "isotopeFront": self.isotope_fname_entry.get(),
                "Quantity": self.quantity_entry.get (),
                "isotopeEnd": self.isotope_ename_entry.get(),
                "model": self.model_name_entry.get(),
                "purchDt": self.purchase_date_entry.get()
            }
            # manPw 엔트리들을 동적으로 추가
            for i in range(1, 19):
                placeholders[f"manPw{i}"] = self.manpower_entries[i-1].get()
                print(f"manPw{i}={self.manpower_entries[i-1].get()}")

            # 각 표에서 검색 단어를 치환합니다.
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in placeholders.items():
                            if key == cell.text:
                                cell.text = cell.text.replace(f"{key}", value)
                                print(f"After replace - {key}={value}, cell.text={cell.text}")            
            

            # 변경된 파일 저장 
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S") 
            new_file_path = f"구매요구서_{timestamp}.docx"
            document.save(new_file_path)
            print(f"파일이 저장되었습니다 : {new_file_path}")
        
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
    
