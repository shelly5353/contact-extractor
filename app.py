import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from contact_extractor import ContactExtractor, save_contacts_to_excel
import logging
import threading
from datetime import datetime

class ContactExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("מערכת חילוץ אנשי קשר")
        self.root.geometry("800x600")
        
        # הגדרת כיוון RTL
        self.root.tk.call('tk', 'scaling', 1.5)  # הגדלת הממשק
        
        # יצירת מסגרת ראשית
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # כפתור לבחירת תיקיית מקור
        self.source_button = ttk.Button(self.main_frame, text="בחר תיקיית מקור", command=self.choose_source_dir)
        self.source_button.grid(row=0, column=1, pady=10, padx=5, sticky=tk.E)
        
        # תיבת טקסט להצגת הנתיב שנבחר
        self.source_path = tk.StringVar()
        self.source_entry = ttk.Entry(self.main_frame, textvariable=self.source_path, state='readonly', width=50)
        self.source_entry.grid(row=0, column=0, pady=10, padx=5, sticky=tk.W)
        
        # כפתור לבחירת קובץ יעד
        self.dest_button = ttk.Button(self.main_frame, text="בחר קובץ יעד", command=self.choose_dest_file)
        self.dest_button.grid(row=1, column=1, pady=10, padx=5, sticky=tk.E)
        
        # תיבת טקסט להצגת נתיב קובץ היעד
        self.dest_path = tk.StringVar()
        self.dest_entry = ttk.Entry(self.main_frame, textvariable=self.dest_path, state='readonly', width=50)
        self.dest_entry.grid(row=1, column=0, pady=10, padx=5, sticky=tk.W)
        
        # כפתור להתחלת העיבוד
        self.process_button = ttk.Button(self.main_frame, text="התחל עיבוד", command=self.start_processing)
        self.process_button.grid(row=2, column=0, columnspan=2, pady=20)
        
        # פס התקדמות
        self.progress = ttk.Progressbar(self.main_frame, length=300, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=2, pady=10)
        
        # אזור לוג
        self.log_frame = ttk.LabelFrame(self.main_frame, text="לוג פעילות", padding="5")
        self.log_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # תיבת טקסט ללוג
        self.log_text = tk.Text(self.log_frame, height=15, width=70, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # סרגל גלילה ללוג
        self.scrollbar = ttk.Scrollbar(self.log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text['yscrollcommand'] = self.scrollbar.set
        
        # הגדרת מצב התחלתי
        self.processing = False
        self.extractor = ContactExtractor()
        
        # הגדרת handler ללוג
        self.log_handler = TextHandler(self.log_text)
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)
        
        # כיוון RTL לכל הרכיבים
        for child in self.main_frame.winfo_children():
            try:
                child.configure(direction='rtl')
            except:
                pass

    def choose_source_dir(self):
        """פתיחת דיאלוג לבחירת תיקיית מקור"""
        dir_path = filedialog.askdirectory(title="בחר תיקיית מקור")
        if dir_path:
            self.source_path.set(dir_path)
            self.log_text.insert(tk.END, f"נבחרה תיקיית מקור: {dir_path}\n")

    def choose_dest_file(self):
        """פתיחת דיאלוג לבחירת קובץ יעד"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="בחר קובץ יעד"
        )
        if file_path:
            self.dest_path.set(file_path)
            self.log_text.insert(tk.END, f"נבחר קובץ יעד: {file_path}\n")

    def start_processing(self):
        """התחלת תהליך העיבוד"""
        if not self.source_path.get() or not self.dest_path.get():
            messagebox.showerror("שגיאה", "יש לבחור תיקיית מקור וקובץ יעד")
            return
            
        if self.processing:
            messagebox.showinfo("מידע", "העיבוד כבר מתבצע")
            return
            
        self.processing = True
        self.process_button.state(['disabled'])
        self.progress['value'] = 0
        
        # התחלת העיבוד בthread נפרד
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        """עיבוד הקבצים"""
        try:
            # איסוף כל הקבצים לעיבוד
            files = []
            for root, _, filenames in os.walk(self.source_path.get()):
                for filename in filenames:
                    if filename.endswith(('.xlsx', '.xls', '.doc', '.docx', '.pdf')):
                        files.append(os.path.join(root, filename))
            
            if not files:
                messagebox.showinfo("מידע", "לא נמצאו קבצים לעיבוד בתיקייה שנבחרה")
                self.processing = False
                self.process_button.state(['!disabled'])
                return
            
            # עיבוד כל הקבצים
            all_contacts = {}
            total_files = len(files)
            
            for i, file_path in enumerate(files, 1):
                try:
                    logging.info(f"מעבד קובץ {i}/{total_files}: {os.path.basename(file_path)}")
                    
                    # עדכון פס ההתקדמות
                    progress = (i / total_files) * 100
                    self.root.after(0, lambda p=progress: self.progress.configure(value=p))
                    
                    # עיבוד הקובץ
                    result = self.extractor.extract_from_docx(file_path) if file_path.endswith('.docx') else \
                            self.extractor.extract_from_doc(file_path) if file_path.endswith('.doc') else \
                            self.extractor.extract_from_xlsx(file_path) if file_path.endswith(('.xlsx', '.xls')) else \
                            self.extractor.extract_from_pdf(file_path)
                    
                    if result:
                        # הוספת אנשי הקשר למילון
                        for contact in result:
                            if contact.name in all_contacts:
                                all_contacts[contact.name].merge(contact)
                            else:
                                all_contacts[contact.name] = contact
                    
                except Exception as e:
                    logging.error(f"שגיאה בעיבוד הקובץ {os.path.basename(file_path)}: {str(e)}")
            
            # שמירת התוצאות
            save_contacts_to_excel(all_contacts, self.dest_path.get())
            
            messagebox.showinfo("סיום", f"העיבוד הסתיים בהצלחה!\nנמצאו {len(all_contacts)} אנשי קשר")
            logging.info(f"העיבוד הסתיים. נמצאו {len(all_contacts)} אנשי קשר")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה במהלך העיבוד: {str(e)}")
            logging.error(f"שגיאה כללית: {str(e)}")
            
        finally:
            self.processing = False
            self.root.after(0, lambda: self.process_button.state(['!disabled']))
            self.root.after(0, lambda: self.progress.configure(value=0))

class TextHandler(logging.Handler):
    """מחלקה לטיפול בלוג והצגתו בממשק"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n")
        self.text_widget.see(tk.END)
        self.text_widget.update()

def main():
    root = tk.Tk()
    root.title("מערכת חילוץ אנשי קשר")
    
    # הגדרת סגנון
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 10))
    style.configure('TLabel', font=('Arial', 10))
    style.configure('TEntry', font=('Arial', 10))
    
    app = ContactExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 