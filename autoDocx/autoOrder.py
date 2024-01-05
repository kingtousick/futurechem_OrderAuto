import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import sqlite3
from docx import Document

class PurchaseRequestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Request App")

        self.tab_control = ttk.Notebook(root)

        # Tab 1: Manual Input
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab1, text="Manual Input")
        self.setup_tab1()

        # Tab 2: Data Registration
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab2, text="Data Registration")
        self.setup_tab2()

        # Tab 3: Search and Update
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab3, text="Search and Update")
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

       ttk.Label(self.tab1, text="핵종명:").grid(column=0, row=3)
       self.isotope_name_entry = ttk.Entry(self.tab1)
       self.isotope_name_entry.grid(column=1, row=3)

       ttk.Label(self.tab1, text="행종수량:").grid(column=0, row=4)
       self.quantity_entry = ttk.Entry(self.tab1)
       self.quantity_entry.grid(column=1, row=4)

       ttk.Label(self.tab1, text="모델명:").grid(column=0, row=5)
       self.model_name_entry = ttk.Entry(self.tab1)
       self.model_name_entry.grid(column=1, row=5)

       ttk.Label(self.tab1, text="투입인력:").grid(column=0, row=6)
       self.manpower_entry = ttk.Entry(self.tab1)
       self.manpower_entry.grid(column=1, row=6)

       ttk.Label(self.tab1, text="구입 예정일:").grid(column=0, row=7)
       self.purchase_date_entry = ttk.Entry(self.tab1)
       self.purchase_date_entry.grid(column=1, row=7)

       # 워드파일 생성
       generate_button = ttk.Button(self.tab1, text="Generate Word File", command=self.generate_word_file)
       generate_button.grid(column=1, row=8)

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
                "Institution": self.institution_entry.get(),
                "InstitutionCode": self.institution_code_entry.get(),
                "Isotope": f"{self.isotope_name_entry.get()} - Quantity: {self.quantity_entry.get()}",
                "Model": self.model_name_entry.get(),
                "Manpower": self.manpower_entry.get(),
                "PurchaseDate": self.purchase_date_entry.get()
            }

            for paragraph in document.paragraphs:
                for key, value in placeholders.items():
                    paragraph.text = paragraph.text.replace(f"{{{key}}}", value)

            # 변경된 파일 저장 
            new_file_path = f"Purchase_Request_{self.institution_entry.get()}_{self.isotope_name_entry.get()}_{self.model_name_entry.get()}.docx"
            document.save(new_file_path)
            print(f"Word document modified and saved as: {new_file_path}")

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