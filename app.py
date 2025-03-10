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
        
        # Configure logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
        
        # Set fixed database path
        self.database_path = '/Users/shellysmac/Documents/Worksking/יעל ישראל/UP/Data_Base.xlsx'
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Source folder selection
        ttk.Label(main_frame, text="תיקיית מקור:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.source_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="בחר תיקייה", command=self.select_source_folder).grid(row=0, column=2, padx=5, pady=5)
        
        # Process button
        ttk.Button(main_frame, text="התחל עיבוד", command=self.process_files).grid(row=1, column=0, columnspan=3, pady=20)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="התקדמות", padding="5")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status_var = tk.StringVar(value="מוכן")
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="לוג", padding="5")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Log text
        self.log_text = tk.Text(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for log
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = scrollbar.set
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_path.set(folder)
            self.log(f"נבחרה תיקיית מקור: {folder}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def process_files(self):
        source_folder = self.source_path.get()
        
        if not source_folder:
            messagebox.showerror("שגיאה", "נא לבחור תיקיית מקור")
            return
        
        try:
            self.status_var.set("מעבד קבצים...")
            self.progress_var.set(0)
            self.root.update()
            
            # Initialize contact extractor
            extractor = ContactExtractor()
            all_contacts = {}
            
            # Load existing contacts from database if it exists
            if os.path.exists(self.database_path):
                try:
                    df = pd.read_excel(self.database_path)
                    for _, row in df.iterrows():
                        contact = Contact(
                            name=row['שם'],
                            phone=row['טלפון'].split('; ')[0] if pd.notna(row['טלפון']) else None,
                            email=row['אימייל'].split('; ')[0] if pd.notna(row['אימייל']) else None,
                            address=row['כתובת'].split('; ')[0] if pd.notna(row['כתובת']) else None,
                            source_file=row['קובץ מקור'] if pd.notna(row['קובץ מקור']) else None
                        )
                        if pd.notna(row['טלפון']):
                            contact.phones.update(row['טלפון'].split('; '))
                        if pd.notna(row['אימייל']):
                            contact.emails.update(row['אימייל'].split('; '))
                        if pd.notna(row['כתובת']):
                            contact.addresses.update(row['כתובת'].split('; '))
                        
                        # Create unique key
                        name_key = contact.name.lower().strip()
                        phone_key = list(contact.phones)[0] if contact.phones else ""
                        email_key = list(contact.emails)[0] if contact.emails else ""
                        key = f"{name_key}_{phone_key}_{email_key}"
                        all_contacts[key] = contact
                    
                    self.log(f"נטענו {len(all_contacts)} אנשי קשר מהקובץ הקיים")
                except Exception as e:
                    self.log(f"שגיאה בטעינת הקובץ הקיים: {str(e)}")
            
            # Get list of files
            files = []
            for root, _, filenames in os.walk(source_folder):
                for filename in filenames:
                    if filename.lower().endswith(('.xlsx', '.xls', '.doc', '.docx')):
                        files.append(os.path.join(root, filename))
            
            total_files = len(files)
            processed_files = 0
            total_contacts = 0
            merged_contacts = 0
            
            # Process each file
            for file_path in files:
                self.log(f"מעבד קובץ: {file_path}")
                
                try:
                    # Extract contacts based on file type
                    if file_path.lower().endswith(('.xlsx', '.xls')):
                        contacts = extractor.extract_from_xlsx(file_path)
                    elif file_path.lower().endswith(('.doc', '.docx')):
                        contacts = extractor.extract_from_doc(file_path)
                    else:
                        continue
                    
                    # Add contacts to dictionary (handling duplicates)
                    for contact in contacts:
                        total_contacts += 1
                        
                        # Create a unique key based on name and contact details
                        name_key = contact.name.lower().strip()
                        phone_key = list(contact.phones)[0] if contact.phones else ""
                        email_key = list(contact.emails)[0] if contact.emails else ""
                        
                        # Try different combinations for matching
                        possible_keys = [
                            f"{name_key}_{phone_key}",
                            f"{name_key}_{email_key}",
                            f"{name_key}_{phone_key}_{email_key}"
                        ]
                        
                        # Check if this contact matches any existing contact
                        found_match = False
                        for key in possible_keys:
                            if key in all_contacts:
                                all_contacts[key].merge(contact)
                                merged_contacts += 1
                                found_match = True
                                break
                        
                        # If no match found, add as new contact
                        if not found_match:
                            all_contacts[possible_keys[0]] = contact
                    
                    processed_files += 1
                    progress = (processed_files / total_files) * 100
                    self.progress_var.set(progress)
                    self.root.update()
                    
                except Exception as e:
                    self.log(f"שגיאה בעיבוד הקובץ {file_path}: {str(e)}")
                    continue
            
            # Save contacts to Excel
            if all_contacts:
                save_contacts_to_excel(all_contacts, self.database_path)
                self.log(f"נשמרו {len(all_contacts)} אנשי קשר לקובץ {self.database_path}")
                self.log(f"סה\"כ אנשי קשר שנמצאו: {total_contacts}")
                self.log(f"כפילויות שמוזגו: {merged_contacts}")
            else:
                self.log("לא נמצאו אנשי קשר")
            
            self.status_var.set("העיבוד הושלם")
            messagebox.showinfo("הצלחה", "העיבוד הושלם בהצלחה")
            
        except Exception as e:
            self.log(f"שגיאה: {str(e)}")
            self.status_var.set("שגיאה")
            messagebox.showerror("שגיאה", f"אירעה שגיאה: {str(e)}")

def main():
    root = tk.Tk()
    app = ContactExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 