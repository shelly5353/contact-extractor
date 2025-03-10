import logging
import re
from typing import Dict, List, Optional, Set, Tuple
from docx import Document
import pdfplumber
import pandas as pd
import json

class Contact:
    def __init__(self, name: str = None, phone: str = None, email: str = None, address: str = None, source_file: str = None):
        self.name = name
        self.phones: Set[str] = {phone} if phone else set()
        self.emails: Set[str] = {email} if email else set()
        self.addresses: Set[str] = {address} if address else set()
        self.source_file = source_file

    def add_phone(self, phone: str) -> None:
        if phone:
            self.phones.add(phone)

    def add_email(self, email: str) -> None:
        if email:
            self.emails.add(email)

    def add_address(self, address: str) -> None:
        if address:
            self.addresses.add(address)

    def merge(self, other: 'Contact') -> None:
        if other.name:
            self.name = other.name
        self.phones.update(other.phones)
        self.emails.update(other.emails)
        self.addresses.update(other.addresses)
        if other.source_file:
            self.source_file = other.source_file

    def is_valid(self) -> bool:
        return bool(self.name and (self.phones or self.emails))

class ContactExtractor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Israeli address pattern with variations
        self.street_prefixes = r'(?:רח\'|רחוב|שד\'|שדרות|דרך|סמטת|סמ\'|שכונת|שכ\')'
        self.city_names = r'(?:תל[- ]אביב|רמת[- ]גן|חיפה|ירושלים|באר[- ]שבע|רחובות|פתח[- ]תקווה|רעננה|הרצליה|גבעתיים|חולון|בת[- ]ים|נתניה|אשדוד|אשקלון|רמלה|לוד|כפר[- ]סבא|רמת[- ]השרון|ראש[- ]העין|פ"ת|פ״ת)'
        
        # Words that indicate a role/title rather than a name
        self.role_words = [
            'מנהל', 'מנהלת', 'יועץ', 'יועצת', 'עובד', 'עובדת', 'אחראי', 'אחראית',
            'מזכיר', 'מזכירה', 'ראש', 'סגן', 'סגנית', 'מפקח', 'מפקחת', 'רכז', 'רכזת',
            'משפטי', 'משפטית', 'כספים', 'פרויקט', 'שירות', 'לקוחות', 'תפעול', 'מחקר',
            'פיתוח', 'משאבי', 'אנוש', 'סניף', 'מחלקה', 'צוות', 'גיוס', 'סוכנים',
            'מכירות', 'שיווק', 'תמיכה', 'הדרכה', 'אזור', 'מרכז', 'צפון', 'דרום',
            'מערב', 'מזרח', 'סוכן', 'סוכנת', 'נציג', 'נציגה', 'מוקד', 'שלוחה',
            'בטיפול', 'במעקב', 'הועבר', 'לא', 'כרגע', 'רלוונטי', 'נרשם', 'נרשמה',
            'מעוניין', 'ניסיתי', 'רחוק', 'קורס', 'לימודי', 'בוקר', 'השקעות'
        ]
        
        # Contact label patterns
        self.contact_labels = [
            'טלפון', 'נייד', 'טל', 'פקס', 'דוא"ל', 'אימייל', 'מייל', 'כתובת',
            'שם', 'איש קשר', 'פרטי התקשרות', 'פרטים', 'פרטי קשר', 'תפקיד',
            'משרד', 'סניף', 'מחלקה', 'יחידה', 'אגף', 'מטה', 'הנהלה'
        ]

        # Compile regex patterns
        self.name_pattern = re.compile(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}')
        self.phone_pattern = re.compile(r'(?:\+972|0)(?:[23489]|5[0-9]|7[0-9])-?\d{7}')
        self.email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
        self.address_pattern = re.compile(f"{self.street_prefixes}\\s+[\\u0590-\\u05FF\\s,]+\\d+|{self.city_names}")
        
        # Initialize contacts list
        self.contacts = []

    def _clean_text(self, text: str) -> str:
        """Clean and normalize text."""
        # Replace various unicode quotation marks and apostrophes
        text = re.sub(r'[""״\'׳]', '"', text)
        # Replace multiple spaces and newlines
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def _is_likely_name_line(self, line: str) -> bool:
        """Check if a line is likely to contain a name."""
        # Skip empty lines or lines that are too short
        if not line or len(line) < 3:
            return False
            
        # Skip lines that are too long (likely not a name)
        if len(line.split()) > 8:  # Increased to allow for titles and roles
            return False
            
        # Skip lines that look like addresses or contain numbers
        if (re.search(self.street_prefixes, line) or 
            re.search(r'\d{3,}', line) or
            re.search(r'קומה|דירה|כניסה|בניין|מתחם', line)):
            return False
            
        # Skip lines that contain contact information
        if any(pattern in line.lower() for pattern in [
            'טלפון:', 'נייד:', 'טל:', 'פקס:', 'דוא"ל:', 'אימייל:', 'מייל:', 'כתובת:',
            'ת.ד.', 'מיקוד', 'ת.ז.', '@', 'שלוחה', 'מוקד:', 'אזור', 'ראש העין'
        ]):
            return False
            
        # Skip lines that are only role words or non-name words
        words = line.split()
        non_name_words = [w for w in words if w not in self.role_words and 
                         w not in self.contact_labels]
        if not non_name_words:
            return False
            
        # Check if line starts with a title or contains name indicators
        if any(indicator in line for indicator in ['שם:', 'איש קשר:', 'נציג:']):
            return True
            
        # Check if line contains Hebrew name pattern (at least two Hebrew words)
        if re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}', line):
            # Additional check for common name formats
            if re.search(r'(?:^|\s)(?:מר|גב\'|ד"ר|עו"ד|רו"ח)\s+[\u0590-\u05FF]+', line):
                return True
            # Check for name followed by role
            if re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,2}\s*[-,]\s*[^,\n]+$', line):
                return True
            # Check for standalone name
            if (len(line.split()) <= 4 and 
                all(re.search(r'[\u0590-\u05FF]', word) for word in line.split())):
                return True
            
        return False

    def _extract_name_from_line(self, line: str) -> Optional[str]:
        """Extract a name from a line of text."""
        # Remove status indicators and comments
        line = re.sub(r'^(?:בטיפול|במעקב|הועבר ל|לא כרגע|לא רלוונטי|נרשמה?|מעוניין|ניסיתי|רחוק)\s*-?\s*', '', line)
        line = re.sub(r'\s*-\s*.*$', '', line)  # Remove everything after a dash
        
        # Remove common prefixes and labels
        line = re.sub(r'^(?:שם|איש קשר|נציג|עובד|פרטי|תפקיד|מייל|טלפון)\s*:?\s*', '', line)
        line = re.sub(r'^(?:לכבוד|לידי|עבור)\s+', '', line)
        
        # Clean up the line
        line = line.strip()
        if not line:
            return None
            
        # Try to match the name pattern
        match = re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}', line)
        if not match:
            return None
            
        # Get the matched text
        name = match.group(0)
        
        # Handle cases with commas or dashes
        parts = re.split(r'[-,]', name)
        name = parts[0].strip()
        
        # If we have additional parts that look like part of the name, include them
        if len(parts) > 1:
            second_part = parts[1].strip()
            # Check if the second part looks like a name continuation
            if (len(second_part.split()) <= 2 and 
                re.search(r'[\u0590-\u05FF]', second_part) and
                not any(word in second_part for word in self.role_words)):
                name = f"{name} {second_part}"
        
        # Skip if it's just a title or common word
        if name in ['מר', 'גב\'', 'ד"ר', 'פרופ\'', 'עו"ד', 'רו"ח', 'אדון', 'גברת']:
            return None
            
        # Skip if it contains only role words or common words
        words = name.split()
        non_role_words = [w for w in words if w not in self.role_words and w not in self.contact_labels]
        if not non_role_words:
            return None
            
        # Verify the name contains Hebrew characters and meets length requirements
        if not re.search(r'[\u0590-\u05FF]', name) or len(name) < 3:
            return None
            
        # Verify maximum number of words and minimum word length
        name_words = name.split()
        if (len(name_words) > 4 or 
            any(len(word) < 2 for word in name_words if re.search(r'[\u0590-\u05FF]', word))):
            return None
            
        # Skip if it's a heading or common phrase
        if name.lower() in [
            'רשימת אנשי קשר', 'אנשי קשר', 'פרטי קשר', 'פרטי התקשרות',
            'מחלקת שירות', 'צוות פיתוח', 'הנהלת חשבונות'
        ]:
            return None
            
        return name

    def _extract_name_and_role(self, text: str) -> Optional[Tuple[str, Optional[str]]]:
        """Extract name and role from text."""
        # Try to find name and role in format: "name - role"
        match = re.match(r'^(.+?)\s*[-,]\s*(.+?)(?=(?:,|\s*(?:טלפון|נייד|טל|פקס|דוא"ל|אימייל|מייל|כתובת)\s*:|$))', text)
        if match:
            name = match.group(1).strip()
            role = match.group(2).strip()
            if self._is_likely_name_line(name):
                name = self._extract_name_from_line(name)
                if name:
                    return name, role
        
        # Try to find just a name
        if self._is_likely_name_line(text):
            name = self._extract_name_from_line(text)
            if name:
                return name, None
        
        return None

    def _extract_phones(self, text: str) -> List[str]:
        """Extract phone numbers from text."""
        phones = set()
        text = self._clean_text(text)
        
        # Define patterns that match Israeli phone numbers
        patterns = [
            r'0[23489]-\d{7}',  # Landline
            r'05[0-9]-\d{7}',   # Mobile
            r'07[0-9]-\d{7}',   # Special
            r'\+972[23489]\d{8}',  # International format landline
            r'\+972-?5[0-9]-?\d{7}'  # International format mobile
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                phone = match.group()
                # Clean up phone number
                phone = re.sub(r'[\s\-]', '', phone)
                if phone.startswith('972'):
                    phone = '+' + phone
                elif phone.startswith('05'):
                    phone = '+972' + phone[1:]
                phones.add(phone)
                
        return sorted(list(phones))

    def _extract_emails(self, text: str) -> List[str]:
        """Extract email addresses from text."""
        emails = set()
        text = self._clean_text(text)
        
        matches = re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        for match in matches:
            email = match.group().lower()
            if self._is_valid_email(email):
                emails.add(email)
                
        return sorted(list(emails))

    def _is_valid_email(self, email: str) -> bool:
        """Validate email address."""
        # Basic validation
        if not re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email):
            return False
            
        # Check for common invalid patterns
        invalid_patterns = [
            r'example\.com$',
            r'test\.com$',
            r'domain\.com$',
            r'@.*@',
            r'\.{2,}',
            r'^[0-9]+@'
        ]
        
        for pattern in invalid_patterns:
            if re.search(pattern, email):
                return False
                
        return True

    def _extract_addresses(self, text: str) -> List[str]:
        """Extract addresses from text."""
        addresses = []
        matches = re.finditer(self.address_pattern, text)
        for match in matches:
            address = match.group().strip()
            addresses.append(address)
        return list(set(addresses))  # Remove duplicates

    def extract_from_xlsx(self, file_path: str) -> List[Contact]:
        """Extract contacts from Excel file."""
        try:
            contacts = []
            
            # Read all sheets in Excel file
            excel_file = pd.ExcelFile(file_path)
            
            for sheet_name in excel_file.sheet_names:
                try:
                    # Read each sheet
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    # Skip empty sheets
                    if df.empty:
                        continue
                        
                    # Try to identify columns
                    headers = [str(col).lower() for col in df.columns]
                    
                    # Find relevant columns
                    name_col = None
                    phone_col = None
                    email_col = None
                    address_col = None
                    
                    for idx, header in enumerate(headers):
                        if any(word in header for word in ['שם', 'איש קשר', 'נציג']):
                            name_col = df.columns[idx]
                        elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון']):
                            phone_col = df.columns[idx]
                        elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל']):
                            email_col = df.columns[idx]
                        elif any(word in header for word in ['כתובת', 'עיר', 'ישוב']):
                            address_col = df.columns[idx]
                    
                    # Process each row
                    for _, row in df.iterrows():
                        contact = Contact(source_file=f"{file_path} - {sheet_name}")
                        
                        # Extract name
                        if name_col and pd.notna(row[name_col]):
                            name = str(row[name_col]).strip()
                            if name and any(c.isalpha() for c in name):
                                contact.name = name
                        
                        # Extract phone
                        if phone_col and pd.notna(row[phone_col]):
                            phone = str(row[phone_col]).strip()
                            extracted_phones = self._extract_phones(phone)
                            for phone in extracted_phones:
                                contact.add_phone(phone)
                        
                        # Extract email
                        if email_col and pd.notna(row[email_col]):
                            email = str(row[email_col]).strip()
                            extracted_emails = self._extract_emails(email)
                            for email in extracted_emails:
                                contact.add_email(email)
                        
                        # Extract address
                        if address_col and pd.notna(row[address_col]):
                            address = str(row[address_col]).strip()
                            if address and any(c.isalpha() for c in address):
                                contact.add_address(address)
                        
                        # Also check entire row text for additional information
                        row_text = ' '.join(str(val) for val in row.values if pd.notna(val))
                        for phone in self._extract_phones(row_text):
                            contact.add_phone(phone)
                        for email in self._extract_emails(row_text):
                            contact.add_email(email)
                        
                        if contact.is_valid():
                            contacts.append(contact)
                            
                except Exception as e:
                    self.logger.error(f"Error processing sheet {sheet_name} in file {file_path}: {str(e)}")
                    continue
            
            return contacts

        except Exception as e:
            self.logger.error(f"Error processing Excel file {file_path}: {str(e)}")
            return []

    def extract_from_doc(self, file_path: str) -> List[Contact]:
        """Extract contacts from DOC/DOCX file."""
        try:
            contacts = []
            
            # For DOCX files
            if file_path.lower().endswith('.docx'):
                doc = Document(file_path)
                
                # Process tables
                for table in doc.tables:
                    # Get headers from first row
                    headers = []
                    if table.rows:
                        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
                    
                    # Find relevant columns
                    name_col = -1
                    phone_col = -1
                    email_col = -1
                    address_col = -1
                    
                    for idx, header in enumerate(headers):
                        if any(word in header for word in ['שם', 'איש קשר', 'נציג']):
                            name_col = idx
                        elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון']):
                            phone_col = idx
                        elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל']):
                            email_col = idx
                        elif any(word in header for word in ['כתובת', 'עיר', 'ישוב']):
                            address_col = idx
                    
                    # Process each row
                    for row in table.rows[1:]:  # Skip header row
                        contact = Contact(source_file=file_path)
                        row_cells = row.cells
                        row_text = ' '.join(cell.text.strip() for cell in row_cells if cell.text.strip())
                        
                        # Extract from specific columns if found
                        if name_col >= 0 and len(row_cells) > name_col:
                            name = row_cells[name_col].text.strip()
                            if name and any(c.isalpha() for c in name):
                                contact.name = name
                        
                        if phone_col >= 0 and len(row_cells) > phone_col:
                            for phone in self._extract_phones(row_cells[phone_col].text):
                                contact.add_phone(phone)
                        
                        if email_col >= 0 and len(row_cells) > email_col:
                            for email in self._extract_emails(row_cells[email_col].text):
                                contact.add_email(email)
                        
                        if address_col >= 0 and len(row_cells) > address_col:
                            address = row_cells[address_col].text.strip()
                            if address and any(c.isalpha() for c in address):
                                contact.add_address(address)
                        
                        # Also check entire row text
                        for phone in self._extract_phones(row_text):
                            contact.add_phone(phone)
                        for email in self._extract_emails(row_text):
                            contact.add_email(email)
                        
                        if contact.is_valid():
                            contacts.append(contact)
                
                # Process paragraphs
                text = "\n".join(para.text for para in doc.paragraphs)
                contacts.extend(self._extract_contacts_from_text(text))
            
            return contacts

        except Exception as e:
            self.logger.error(f"Error extracting from DOC/DOCX {file_path}: {str(e)}")
            return []

    def extract_from_pdf(self, file_path: str) -> List[Contact]:
        """Extract contacts from PDF file."""
        try:
            contacts = []
            text = ""
            
            # Try pdfplumber
            with pdfplumber.open(file_path) as pdf:
                # Process each page
                for page in pdf.pages:
                    # Extract text
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                    
                    # Extract tables
                    tables = page.extract_tables()
                    for table in tables:
                        # Get headers from first row
                        headers = [str(cell).strip().lower() if cell else '' for cell in table[0]]
                        
                        # Find relevant columns
                        name_col = -1
                        phone_col = -1
                        email_col = -1
                        address_col = -1
                        
                        for idx, header in enumerate(headers):
                            if any(word in header for word in ['שם', 'איש קשר', 'נציג']):
                                name_col = idx
                            elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון']):
                                phone_col = idx
                            elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל']):
                                email_col = idx
                            elif any(word in header for word in ['כתובת', 'עיר', 'ישוב']):
                                address_col = idx
                        
                        # Process each row
                        for row in table[1:]:  # Skip header row
                            contact = Contact(source_file=file_path)
                            row_text = ' '.join(str(cell).strip() for cell in row if cell)
                            
                            # Extract from specific columns if found
                            if name_col >= 0 and len(row) > name_col and row[name_col]:
                                name = str(row[name_col]).strip()
                                if name and any(c.isalpha() for c in name):
                                    contact.name = name
                            
                            if phone_col >= 0 and len(row) > phone_col and row[phone_col]:
                                for phone in self._extract_phones(str(row[phone_col])):
                                    contact.add_phone(phone)
                            
                            if email_col >= 0 and len(row) > email_col and row[email_col]:
                                for email in self._extract_emails(str(row[email_col])):
                                    contact.add_email(email)
                            
                            if address_col >= 0 and len(row) > address_col and row[address_col]:
                                address = str(row[address_col]).strip()
                                if address and any(c.isalpha() for c in address):
                                    contact.add_address(address)
                            
                            # Also check entire row text
                            for phone in self._extract_phones(row_text):
                                contact.add_phone(phone)
                            for email in self._extract_emails(row_text):
                                contact.add_email(email)
                            
                            if contact.is_valid():
                                contacts.append(contact)
            
            # Process extracted text
            if text:
                contacts.extend(self._extract_contacts_from_text(text))
            
            return contacts

        except Exception as e:
            self.logger.error(f"Error extracting from PDF {file_path}: {str(e)}")
            return []

    def _extract_contacts_from_text(self, text: str) -> List[Contact]:
        """Extract contacts from raw text."""
        contacts = []
        
        # Split text into lines
        lines = text.split('\n')
        current_contact = None
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_contact and current_contact.is_valid():
                    contacts.append(current_contact)
                current_contact = None
                continue
            
            # Try to extract name and role
            name_and_role = self._extract_name_and_role(line)
            if name_and_role:
                if current_contact and current_contact.is_valid():
                    contacts.append(current_contact)
                name, role = name_and_role
                current_contact = Contact(name=name, source_file=None)
                if role:
                    current_contact.add_role(role)
            
            # Extract other information
            if current_contact:
                for phone in self._extract_phones(line):
                    current_contact.add_phone(phone)
                for email in self._extract_emails(line):
                    current_contact.add_email(email)
                for address in self._extract_addresses(line):
                    current_contact.add_address(address)
        
        # Add last contact if valid
        if current_contact and current_contact.is_valid():
            contacts.append(current_contact)
        
        return contacts

    def load_contacts_from_excel(self, file_path):
        """טוען אנשי קשר מקובץ Excel קיים"""
        try:
            df = pd.read_excel(file_path)
            contacts = []
            
            for _, row in df.iterrows():
                contact = Contact(
                    name=row['שם'],
                    phone=row['טלפון'].split('; ')[0] if pd.notna(row['טלפון']) else None,
                    email=row['אימייל'].split('; ')[0] if pd.notna(row['אימייל']) else None,
                    address=row['כתובת'].split('; ')[0] if pd.notna(row['כתובת']) else None
                )
                
                if pd.notna(row['טלפון']):
                    contact.phones.update(row['טלפון'].split('; '))
                if pd.notna(row['אימייל']):
                    contact.emails.update(row['אימייל'].split('; '))
                if pd.notna(row['כתובת']):
                    contact.addresses.update(row['כתובת'].split('; '))
                
                contacts.append(contact)
            
            self.logger.info(f"נטענו {len(contacts)} אנשי קשר מהקובץ {file_path}")
            return contacts
            
        except Exception as e:
            self.logger.error(f"שגיאה בטעינת אנשי קשר מהקובץ {file_path}: {str(e)}")
            return []
            
    def remove_duplicates(self, contacts):
        """מסיר כפילויות מרשימת אנשי קשר"""
        unique_contacts = {}
        
        for contact in contacts:
            # יצירת מפתח ייחודי
            name_key = contact.name.lower().strip()
            phone_key = list(contact.phones)[0] if contact.phones else ""
            email_key = list(contact.emails)[0] if contact.emails else ""
            key = f"{name_key}_{phone_key}_{email_key}"
            
            if key in unique_contacts:
                # מיזוג אנשי קשר זהים
                unique_contacts[key].merge(contact)
            else:
                unique_contacts[key] = contact
        
        return list(unique_contacts.values())
        
    def extract_contacts(self, file_path):
        """חולץ אנשי קשר מכל סוגי הקבצים הנתמכים"""
        try:
            if file_path.lower().endswith(('.xlsx', '.xls')):
                return self.extract_from_xlsx(file_path)
            elif file_path.lower().endswith(('.doc', '.docx')):
                return self.extract_from_doc(file_path)
            else:
                self.logger.warning(f"סוג קובץ לא נתמך: {file_path}")
                return []
        except Exception as e:
            self.logger.error(f"שגיאה בחילוץ אנשי קשר מהקובץ {file_path}: {str(e)}")
            return []

def save_contacts_to_excel(contacts: Dict[str, Contact], output_path: str) -> None:
    """Save contacts to Excel file."""
    try:
        # Convert contacts to list of dictionaries
        data = []
        for contact in contacts.values():
            data.append({
                'שם': contact.name,
                'טלפון': '; '.join(contact.phones),
                'אימייל': '; '.join(contact.emails),
                'כתובת': '; '.join(contact.addresses),
                'קובץ מקור': contact.source_file
            })
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame(data)
        
        # Create Excel writer with xlsxwriter engine
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='אנשי קשר')
        
        # Get the workbook and the worksheet
        workbook = writer.book
        worksheet = writer.sheets['אנשי קשר']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center',
            'bg_color': '#D7E4BC',
            'border': 1
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })
        
        # Set column widths based on content
        for idx, col in enumerate(df.columns):
            # Get max length in the column
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            # Add some padding
            max_length += 2
            # Set column width (max 100 characters)
            worksheet.set_column(idx, idx, min(max_length, 100), cell_format)
        
        # Format header row
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Freeze the first row
        worksheet.freeze_panes(1, 0)
        
        # Save the file
        writer.close()
        
    except Exception as e:
        logging.error(f"Error saving contacts to Excel: {str(e)}") 