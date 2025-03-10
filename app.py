import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from contact_extractor import ContactExtractor, save_contacts_to_excel
import logging
import pandas as pd

class ContactExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("מחלץ אנשי קשר")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # הגדרת משתנים
        self.source_folder = tk.StringVar()
        self.target_file = tk.StringVar()
        self.output_mode = tk.StringVar(value="new")  # ברירת מחדל: קובץ חדש
        
        # הגדרת לוגר
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
        
        # יצירת הממשק
        self.create_widgets()
        
        # הגדרת הודעות לוג
        self.log_messages = []
        
    def create_widgets(self):
        # מסגרת ראשית
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # בחירת תיקיית מקור
        source_frame = ttk.LabelFrame(main_frame, text="תיקיית מקור", padding="5")
        source_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(source_frame, text="תיקייה:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(source_frame, textvariable=self.source_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(source_frame, text="בחר תיקייה", command=self.select_source_folder).grid(row=0, column=2)
        
        # בחירת קובץ יעד
        target_frame = ttk.LabelFrame(main_frame, text="קובץ יעד", padding="5")
        target_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(target_frame, text="קובץ:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(target_frame, textvariable=self.target_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(target_frame, text="בחר קובץ", command=self.select_target_file).grid(row=0, column=2)
        
        # בחירת מצב פלט
        output_frame = ttk.LabelFrame(main_frame, text="מצב פלט", padding="5")
        output_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Radiobutton(output_frame, text="יצור קובץ חדש", variable=self.output_mode, value="new").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(output_frame, text="עדכון קובץ קיים", variable=self.output_mode, value="update").grid(row=0, column=1, sticky=tk.W)
        
        # מסגרת התקדמות
        progress_frame = ttk.LabelFrame(main_frame, text="התקדמות", padding="5")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.status_label = ttk.Label(progress_frame, text="")
        self.status_label.grid(row=1, column=0, columnspan=2, sticky=tk.W)
        
        # מסגרת לוג
        log_frame = ttk.LabelFrame(main_frame, text="לוג פעילות", padding="5")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = tk.Text(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # כפתור התחלה
        ttk.Button(main_frame, text="התחל עיבוד", command=self.process_files).grid(row=5, column=0, columnspan=2, pady=10)
        
        # הגדרת הרחבה של החלון
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)
            self.log_message(f"נבחרה תיקיית מקור: {folder}")
            
    def select_target_file(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="אנשי_קשר.xlsx"
        )
        if file:
            self.target_file.set(file)
            self.log_message(f"נבחר קובץ יעד: {file}")
            
    def log_message(self, message):
        self.log_messages.append(message)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def process_files(self):
        source_folder = self.source_folder.get()
        target_file = self.target_file.get()
        
        if not source_folder or not target_file:
            messagebox.showerror("שגיאה", "נא לבחור תיקיית מקור וקובץ יעד")
            return
            
        try:
            # איפוס סרגל ההתקדמות
            self.progress_var.set(0)
            self.status_label.config(text="מתחיל בעיבוד...")
            self.root.update()
            
            # יצירת מחלץ אנשי קשר
            extractor = ContactExtractor()
            
            # חילוץ אנשי קשר מכל הקבצים
            contacts = []
            files = [f for f in os.listdir(source_folder) if f.endswith(('.xlsx', '.xls', '.doc', '.docx'))]
            total_files = len(files)
            
            for i, file in enumerate(files):
                file_path = os.path.join(source_folder, file)
                self.log_message(f"מעבד קובץ: {file}")
                
                try:
                    file_contacts = extractor.extract_contacts(file_path)
                    contacts.extend(file_contacts)
                    self.log_message(f"נמצאו {len(file_contacts)} אנשי קשר בקובץ {file}")
                except Exception as e:
                    self.log_message(f"שגיאה בעיבוד הקובץ {file}: {str(e)}")
                
                # עדכון התקדמות
                progress = (i + 1) / total_files * 100
                self.progress_var.set(progress)
                self.status_label.config(text=f"עובד קובץ {i+1} מתוך {total_files}")
                self.root.update()
            
            if not contacts:
                messagebox.showwarning("אזהרה", "לא נמצאו אנשי קשר בקבצים")
                return
                
            # שמירת התוצאות
            if self.output_mode.get() == "new":
                # יצירת קובץ חדש
                extractor.save_contacts_to_excel(contacts, target_file)
                self.log_message(f"נשמרו {len(contacts)} אנשי קשר לקובץ חדש: {target_file}")
            else:
                # עדכון קובץ קיים
                if os.path.exists(target_file):
                    existing_contacts = extractor.load_contacts_from_excel(target_file)
                    all_contacts = existing_contacts + contacts
                    # הסרת כפילויות
                    unique_contacts = extractor.remove_duplicates(all_contacts)
                    extractor.save_contacts_to_excel(unique_contacts, target_file)
                    self.log_message(f"עודכנו {len(unique_contacts)} אנשי קשר בקובץ: {target_file}")
                else:
                    messagebox.showerror("שגיאה", "קובץ היעד לא קיים")
                    return
            
            self.status_label.config(text="העיבוד הושלם בהצלחה!")
            messagebox.showinfo("הצלחה", f"נשמרו {len(contacts)} אנשי קשר בהצלחה")
            
        except Exception as e:
            self.log_message(f"שגיאה: {str(e)}")
            messagebox.showerror("שגיאה", f"אירעה שגיאה: {str(e)}")

def main():
    root = tk.Tk()
    app = ContactExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 