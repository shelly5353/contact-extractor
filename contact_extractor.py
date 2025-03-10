import logging
import re
from typing import Dict, List, Optional, Set, Tuple
from docx import Document
import pdfplumber
from openpyxl import load_workbook
import json
import os
import pandas as pd
import PyPDF2

class Contact:
    def __init__(self, name: str = None, phone: str = None, email: str = None, address: str = None, source_file: str = None):
        self.name = name
        self.phones: Set[str] = {phone} if phone else set()
        self.emails: Set[str] = {email} if email else set()
        self.addresses: Set[str] = {address} if address else set()
        self.source_file = source_file
        self.role = None
        self.logger = logging.getLogger(__name__)
        self.logger.debug(f"יצירת איש קשר חדש: {name}")

    def add_phone(self, phone: str) -> None:
        if phone:
            # ניקוי מספר הטלפון
            phone = re.sub(r'[\s\-\(\)]', '', phone)
            # בדיקת תקינות בסיסית - לפחות 9 ספרות
            if len(re.findall(r'\d', phone)) >= 9:
                if phone.startswith('972'):
                    phone = '+' + phone
                elif phone.startswith('05'):
                    phone = '+972' + phone[1:]
                elif phone.startswith('0'):
                    phone = '+972' + phone[1:]
                self.phones.add(phone)
                self.logger.debug(f"הוספת מספר טלפון: {phone} לאיש קשר {self.name}")
            else:
                self.logger.debug(f"דילוג על מספר טלפון לא תקין: {phone}")

    def add_email(self, email: str) -> None:
        if email:
            email = email.lower().strip()
            if re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email):
                self.emails.add(email)
                self.logger.debug(f"הוספת כתובת מייל: {email} לאיש קשר {self.name}")
            else:
                self.logger.debug(f"דילוג על כתובת מייל לא תקינה: {email}")

    def add_address(self, address: str) -> None:
        if address:
            address = address.strip()
            # בדיקה שהכתובת מכילה לפחות 2 מילים ומספר או אותיות בעברית
            words = address.split()
            if len(words) >= 2 and (any(c.isdigit() for c in address) or re.search(r'[\u0590-\u05FF]', address)):
                self.addresses.add(address)
                self.logger.debug(f"הוספת כתובת: {address} לאיש קשר {self.name}")
            else:
                self.logger.debug(f"דילוג על כתובת לא תקינה: {address}")

    def add_role(self, role: str) -> None:
        if role:
            role = role.strip()
            if len(role) >= 2 and any(c.isalpha() for c in role):
                self.role = role
                self.logger.debug(f"הוספת תפקיד: {role} לאיש קשר {self.name}")
            else:
                self.logger.debug(f"דילוג על תפקיד לא תקין: {role}")

    def merge(self, other: 'Contact') -> None:
        if other.name and (not self.name or len(other.name) > len(self.name)):
            self.name = other.name
            self.logger.debug(f"עדכון שם בזמן מיזוג: {self.name}")
        self.phones.update(other.phones)
        self.emails.update(other.emails)
        self.addresses.update(other.addresses)
        if other.role and (not self.role or len(other.role) > len(self.role)):
            self.role = other.role
            self.logger.debug(f"עדכון תפקיד בזמן מיזוג: {self.role}")
        if other.source_file:
            self.source_file = other.source_file
        self.logger.debug(f"מיזוג הושלם עבור איש קשר: {self.name}")

    def is_valid(self) -> bool:
        """בודק אם איש הקשר תקין"""
        # בדיקת שם
        if not self.name or not any(c.isalpha() for c in self.name):
            self.logger.debug(f"איש קשר לא תקין - שם חסר או לא תקין: {self.name}")
            return False
            
        # בדיקת אורך שם
        if len(self.name.strip()) < 2:
            self.logger.debug(f"איש קשר לא תקין - שם קצר מדי: {self.name}")
            return False
            
        # בדיקת פרטי קשר - חייב לפחות טלפון או אימייל
        has_contact = bool(self.phones or self.emails)
        if not has_contact:
            self.logger.debug(f"איש קשר לא תקין - אין פרטי קשר: {self.name}")
            return False
            
        # בדיקת תקינות טלפונים
        for phone in self.phones:
            if not re.match(r'^\+?972\d{8,9}$', re.sub(r'[\s\-]', '', phone)):
                self.logger.debug(f"איש קשר לא תקין - מספר טלפון לא תקין: {phone}")
                return False
                
        # בדיקת תקינות אימיילים
        for email in self.emails:
            if not re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email):
                self.logger.debug(f"איש קשר לא תקין - כתובת אימייל לא תקינה: {email}")
                return False
                
        self.logger.debug(f"איש קשר תקין: {self.name} (טלפונים: {len(self.phones)}, אימיילים: {len(self.emails)})")
        return True

class ContactExtractor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Israeli address pattern with variations
        self.street_prefixes = r'(?:רח\'|רחוב|שד\'|שדרות|דרך|סמטת|סמ\'|שכונת|שכ\'|בול\'|בולוורד|כיכר|מתחם)'
        self.city_names = r'(?:תל[- ]אביב|רמת[- ]גן|חיפה|ירושלים|באר[- ]שבע|רחובות|פתח[- ]תקווה|רעננה|הרצליה|גבעתיים|חולון|בת[- ]ים|נתניה|אשדוד|אשקלון|רמלה|לוד|כפר[- ]סבא|רמת[- ]השרון|ראש[- ]העין|פ"ת|פ״ת|רמת גן|רמת-גן|תל אביב|תל-אביב|באר שבע|באר-שבע)'
        
        # Words that indicate a role/title rather than a name
        self.role_words = [
            'מנהל', 'מנהלת', 'יועץ', 'יועצת', 'עובד', 'עובדת', 'אחראי', 'אחראית',
            'מזכיר', 'מזכירה', 'ראש', 'סגן', 'סגנית', 'מפקח', 'מפקחת', 'רכז', 'רכזת',
            'משפטי', 'משפטית', 'כספים', 'פרויקט', 'שירות', 'לקוחות', 'תפעול', 'מחקר',
            'פיתוח', 'משאבי', 'אנוש', 'סניף', 'מחלקה', 'צוות', 'גיוס', 'סוכנים',
            'מכירות', 'שיווק', 'תמיכה', 'הדרכה', 'אזור', 'מרכז', 'צפון', 'דרום',
            'מערב', 'מזרח', 'סוכן', 'סוכנת', 'נציג', 'נציגה', 'מוקד', 'שלוחה',
            'בטיפול', 'במעקב', 'הועבר', 'לא', 'כרגע', 'רלוונטי', 'נרשם', 'נרשמה',
            'מעוניין', 'ניסיתי', 'רחוק', 'קורס', 'לימודי', 'בוקר', 'השקעות',
            'מנכ"ל', 'מנכ"לית', 'סמנכ"ל', 'סמנכ"לית', 'מנהל', 'מנהלת', 'מנהלים',
            'מנהלות', 'מנהלי', 'מנהלות', 'מנהל', 'מנהלת', 'מנהלים', 'מנהלות',
            'מנהלי', 'מנהלות', 'מנהל', 'מנהלת', 'מנהלים', 'מנהלות', 'מנהלי', 'מנהלות'
        ]
        
        # Contact label patterns
        self.contact_labels = [
            'טלפון', 'נייד', 'טל', 'פקס', 'דוא"ל', 'אימייל', 'מייל', 'כתובת',
            'שם', 'איש קשר', 'פרטי התקשרות', 'פרטים', 'פרטי קשר', 'תפקיד',
            'משרד', 'סניף', 'מחלקה', 'יחידה', 'אגף', 'מטה', 'הנהלה',
            'טלפון:', 'נייד:', 'טל:', 'פקס:', 'דוא"ל:', 'אימייל:', 'מייל:', 'כתובת:',
            'שם:', 'איש קשר:', 'פרטי התקשרות:', 'פרטים:', 'פרטי קשר:', 'תפקיד:',
            'משרד:', 'סניף:', 'מחלקה:', 'יחידה:', 'אגף:', 'מטה:', 'הנהלה:'
        ]

        # Compile regex patterns
        self.name_pattern = re.compile(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}')
        self.phone_pattern = re.compile(r'(?:\+972|05\d|\+972-\d{2}|0\d{1,2}[-.]?)\d{7,8}')
        self.email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
        self.address_pattern = re.compile(f"{self.street_prefixes}\\s+[\\u0590-\\u05FF\\s,]+\\d+|{self.city_names}")
        
        # Initialize contacts list
        self.contacts = []

    def scan_directory(self, directory_path: str) -> List[Contact]:
        """סורק תיקייה ומחלץ אנשי קשר מכל הקבצים הנתמכים"""
        contacts = []
        try:
            for root, dirs, files in os.walk(directory_path):
                for file in files:
                    if file.lower().endswith(tuple(f'.{ext}' for ext in ['xlsx', 'xls', 'doc', 'docx', 'pdf'])):
                        file_path = os.path.join(root, file)
                        try:
                            self.logger.info(f"מעבד קובץ: {file_path}")
                            file_contacts = self.extract_contacts(file_path)
                            contacts.extend(file_contacts)
                            self.logger.info(f"נמצאו {len(file_contacts)} אנשי קשר בקובץ {file_path}")
                        except Exception as e:
                            self.logger.error(f"שגיאה בעיבוד הקובץ {file_path}: {str(e)}")
                            continue
        except Exception as e:
            self.logger.error(f"שגיאה בסריקת התיקייה {directory_path}: {str(e)}")
        
        self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בסך הכל בתיקייה {directory_path}")
        return contacts

    def _clean_text(self, text: str) -> str:
        """מנקה ומנרמל טקסט"""
        if not text:
            return ""
            
        # המרה למחרוזת
        text = str(text)
        
        # החלפת גרשיים שונים
        text = re.sub(r'[""״\'׳]', '"', text)
        
        # החלפת מקפים שונים
        text = re.sub(r'[-–—]', '-', text)
        
        # החלפת פסיקים שונים
        text = re.sub(r'[,،]', ',', text)
        
        # החלפת נקודות שונים
        text = re.sub(r'[\.٫]', '.', text)
        
        # החלפת נקודותיים שונים
        text = re.sub(r'[:׃]', ':', text)
        
        # החלפת סוגריים שונים
        text = re.sub(r'[\(\)]', '(', text)
        text = re.sub(r'[\[\]]', '[', text)
        text = re.sub(r'[\{\}]', '{', text)
        
        # החלפת רווחים מיותרים
        text = re.sub(r'\s+', ' ', text)
        
        # הסרת רווחים בתחילת וסוף הטקסט
        text = text.strip()
        
        return text

    def _is_likely_name_line(self, line: str) -> bool:
        """בודק אם שורה נראית כמו שם"""
        # דילוג על שורות ריקות או קצרות מדי
        if not line or len(line) < 3:
            return False
            
        # דילוג על שורות ארוכות מדי (כנראה לא שם)
        if len(line.split()) > 8:  # הוגדל כדי לאפשר תארים ותפקידים
            return False
            
        # דילוג על שורות שנראות כמו כתובות או מכילות מספרים
        if (re.search(self.street_prefixes, line) or 
            re.search(r'\d{3,}', line) or
            re.search(r'קומה|דירה|כניסה|בניין|מתחם', line)):
            return False
            
        # דילוג על שורות שמכילות פרטי קשר
        if any(pattern in line.lower() for pattern in [
            'טלפון:', 'נייד:', 'טל:', 'פקס:', 'דוא"ל:', 'אימייל:', 'מייל:', 'כתובת:',
            'ת.ד.', 'מיקוד', 'ת.ז.', '@', 'שלוחה', 'מוקד:', 'אזור', 'ראש העין'
        ]):
            return False
            
        # דילוג על שורות שמכילות רק מילות תפקיד או מילים נפוצות
        words = line.split()
        non_name_words = [w for w in words if w not in self.role_words and 
                         w not in self.contact_labels]
        if not non_name_words:
            return False
            
        # בדיקה אם השורה מתחילה בתואר או מכילה אינדיקטורים לשם
        if any(indicator in line for indicator in ['שם:', 'איש קשר:', 'נציג:']):
            return True
            
        # בדיקה אם השורה מכילה תבנית שם עברי (לפחות שתי מילים בעברית)
        if re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}', line):
            # בדיקה נוספת לפורמטים נפוצים של שמות
            if re.search(r'(?:^|\s)(?:מר|גב\'|ד"ר|עו"ד|רו"ח)\s+[\u0590-\u05FF]+', line):
                return True
            # בדיקה לשם ואחריו תפקיד
            if re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,2}\s*[-,]\s*[^,\n]+$', line):
                return True
            # בדיקה לשם עצמאי
            if (len(line.split()) <= 4 and 
                all(re.search(r'[\u0590-\u05FF]', word) for word in line.split())):
                return True
            
        return False

    def _extract_name_from_line(self, line: str) -> Optional[str]:
        """מחלץ שם משורת טקסט"""
        # הסרת סטטוסים והערות
        line = re.sub(r'^(?:בטיפול|במעקב|הועבר ל|לא כרגע|לא רלוונטי|נרשמה?|מעוניין|ניסיתי|רחוק)\s*-?\s*', '', line)
        line = re.sub(r'\s*-\s*.*$', '', line)  # הסרת כל מה שאחר המקף
        
        # הסרת תחיליות ותוויות נפוצות
        line = re.sub(r'^(?:שם|איש קשר|נציג|עובד|פרטי|תפקיד|מייל|טלפון)\s*:?\s*', '', line)
        line = re.sub(r'^(?:לכבוד|לידי|עבור)\s+', '', line)
        
        # ניקוי השורה
        line = line.strip()
        if not line:
            return None
            
        # ניסיון להתאים את תבנית השם
        match = re.search(r'[\u0590-\u05FF]+(?:\s+[\u0590-\u05FF]+){1,3}', line)
        if not match:
            return None
            
        # קבלת הטקסט המתאים
        name = match.group(0)
        
        # טיפול במקרים עם פסיקים או מקפים
        parts = re.split(r'[-,]', name)
        name = parts[0].strip()
        
        # אם יש חלקים נוספים שנראים כמו חלק מהשם, כוללים אותם
        if len(parts) > 1:
            second_part = parts[1].strip()
            # בדיקה אם החלק השני נראה כמו המשך שם
            if (len(second_part.split()) <= 2 and 
                re.search(r'[\u0590-\u05FF]', second_part) and
                not any(word in second_part for word in self.role_words)):
                name = f"{name} {second_part}"
        
        # דילוג אם זה רק תואר או מילה נפוצה
        if name in ['מר', 'גב\'', 'ד"ר', 'פרופ\'', 'עו"ד', 'רו"ח', 'אדון', 'גברת']:
            return None
            
        # דילוג אם מכיל רק מילות תפקיד או מילים נפוצות
        words = name.split()
        non_role_words = [w for w in words if w not in self.role_words and w not in self.contact_labels]
        if not non_role_words:
            return None
            
        # אימות שהשם מכיל אותיות עבריות ומענה על דרישות אורך
        if not re.search(r'[\u0590-\u05FF]', name) or len(name) < 3:
            return None
            
        # אימות מספר מילים מקסימלי ואורך מינימלי של מילה
        name_words = name.split()
        if (len(name_words) > 4 or 
            any(len(word) < 2 for word in name_words if re.search(r'[\u0590-\u05FF]', word))):
            return None
            
        # דילוג אם זה כותרת או ביטוי נפוץ
        if name.lower() in [
            'רשימת אנשי קשר', 'אנשי קשר', 'פרטי קשר', 'פרטי התקשרות',
            'מחלקת שירות', 'צוות פיתוח', 'הנהלת חשבונות', 'פרטי התקשרות',
            'פרטי קשר', 'פרטים', 'פרטי התקשרות', 'פרטי קשר', 'פרטים',
            'פרטי התקשרות', 'פרטי קשר', 'פרטים', 'פרטי התקשרות', 'פרטי קשר'
        ]:
            return None
            
        self.logger.debug(f"נמצא שם: {name}")
        return name

    def _extract_name_and_role(self, text: str) -> Optional[Tuple[str, Optional[str]]]:
        """חולץ שם ותפקיד מטקסט"""
        # ניסיון למצוא שם ותפקיד בפורמט: "שם - תפקיד"
        match = re.match(r'^(.+?)\s*[-,]\s*(.+?)(?=(?:,|\s*(?:טלפון|נייד|טל|פקס|דוא"ל|אימייל|מייל|כתובת)\s*:|$))', text)
        if match:
            name = match.group(1).strip()
            role = match.group(2).strip()
            if self._is_likely_name_line(name):
                name = self._extract_name_from_line(name)
                if name:
                    self.logger.debug(f"נמצא שם ותפקיד: {name} - {role}")
                    return name, role
        
        # ניסיון למצוא רק שם
        if self._is_likely_name_line(text):
            name = self._extract_name_from_line(text)
            if name:
                self.logger.debug(f"נמצא שם: {name}")
                return name, None
        
        return None

    def _extract_phones(self, text: str) -> List[str]:
        """חולץ מספרי טלפון מטקסט"""
        phones = set()
        text = self._clean_text(text)
        
        # תבניות למספרי טלפון ישראליים
        patterns = [
            r'0[23489]-\d{7}',  # קווי
            r'05[0-9]-\d{7}',   # ניידים
            r'07[0-9]-\d{7}',   # מיוחדים
            r'\+972[23489]\d{8}',  # פורמט בינלאומי קווי
            r'\+972-?5[0-9]-?\d{7}',  # פורמט בינלאומי נייד
            r'0[23489]\d{7}',  # קווי ללא מקף
            r'05[0-9]\d{7}',   # ניידים ללא מקף
            r'07[0-9]\d{7}',   # מיוחדים ללא מקף
            r'972[23489]\d{8}',  # פורמט בינלאומי קווי ללא פלוס
            r'9725[0-9]\d{7}',  # פורמט בינלאומי נייד ללא פלוס
            r'0[23489]\s*\d{3}\s*\d{4}',  # קווי עם רווחים
            r'05[0-9]\s*\d{3}\s*\d{4}',   # ניידים עם רווחים
            r'07[0-9]\s*\d{3}\s*\d{4}',   # מיוחדים עם רווחים
            r'0[23489]\(\d{3}\)\d{4}',  # קווי עם סוגריים
            r'05[0-9]\(\d{3}\)\d{4}',   # ניידים עם סוגריים
            r'07[0-9]\(\d{3}\)\d{4}'    # מיוחדים עם סוגריים
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                phone = match.group()
                # ניקוי מספר הטלפון
                phone = re.sub(r'[\s\-\(\)]', '', phone)
                
                # המרה לפורמט סטנדרטי
                if phone.startswith('972'):
                    phone = '+' + phone
                elif phone.startswith('05'):
                    phone = '+972' + phone[1:]
                elif phone.startswith('0'):
                    phone = '+972' + phone[1:]
                
                # בדיקת תקינות
                if len(phone) >= 10:  # מספר מינימלי של ספרות
                    phones.add(phone)
                    self.logger.debug(f"נמצא מספר טלפון: {phone}")
                
        return sorted(list(phones))

    def _extract_emails(self, text: str) -> List[str]:
        """חולץ כתובות מייל מטקסט"""
        emails = set()
        text = self._clean_text(text)
        
        # תבניות לכתובות מייל
        patterns = [
            r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',  # תבנית בסיסית
            r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\.[a-zA-Z]{2,}',  # תתי-דומיינים
            r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\.[a-zA-Z]{2,}\.[a-zA-Z]{2,}'  # תתי-דומיינים מרובים
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                email = match.group().lower().strip()
                
                # בדיקת תקינות
                if self._is_valid_email(email):
                    emails.add(email)
                    self.logger.debug(f"נמצאה כתובת מייל: {email}")
                
        return sorted(list(emails))

    def _is_valid_email(self, email: str) -> bool:
        """בודק תקינות כתובת מייל"""
        # בדיקה בסיסית
        if not re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email):
            return False
            
        # בדיקת תבניות לא תקינות
        invalid_patterns = [
            r'example\.com$',
            r'test\.com$',
            r'domain\.com$',
            r'@.*@',
            r'\.{2,}',
            r'^[0-9]+@',
            r'\.$',
            r'\.@',
            r'@\.',
            r'\.@\.',
            r'\.com\.',
            r'\.net\.',
            r'\.org\.',
            r'\.edu\.',
            r'\.gov\.',
            r'\.mil\.',
            r'\.int\.',
            r'\.biz\.',
            r'\.info\.',
            r'\.name\.',
            r'\.pro\.',
            r'\.aero\.',
            r'\.coop\.',
            r'\.museum\.',
            r'\.jobs\.',
            r'\.mobi\.',
            r'\.tel\.',
            r'\.travel\.',
            r'\.cat\.',
            r'\.asia\.',
            r'\.post\.',
            r'\.xxx\.'
        ]
        
        for pattern in invalid_patterns:
            if re.search(pattern, email):
                return False
                
        # בדיקת אורך
        if len(email) < 5 or len(email) > 254:
            return False
            
        # בדיקת תווים מיוחדים
        if re.search(r'[<>()\[\]\\,;:{}]', email):
            return False
            
        # בדיקת תווים מיותרים
        if re.search(r'[^\x00-\x7F]', email):
            return False
            
        return True

    def _extract_addresses(self, text: str) -> List[str]:
        """חולץ כתובות מטקסט"""
        addresses = []
        text = self._clean_text(text)
        
        # תבניות לכתובות
        patterns = [
            # כתובת עם רחוב ומספר
            f"{self.street_prefixes}\\s+[\\u0590-\\u05FF\\s,]+\\d+",
            # כתובת עם עיר
            f"{self.street_prefixes}\\s+[\\u0590-\\u05FF\\s,]+\\d+\\s*,\\s*{self.city_names}",
            # כתובת עם ת.ד.
            r'ת\.?ד\.?\s*\d+',
            # כתובת עם מיקוד
            r'מיקוד\s*\d+',
            # כתובת עם בניין
            r'בניין\s+[0-9א-ת]+',
            # כתובת עם דירה
            r'דירה\s+[0-9א-ת]+',
            # כתובת עם קומה
            r'קומה\s+[0-9א-ת]+',
            # כתובת עם כניסה
            r'כניסה\s+[0-9א-ת]+',
            # כתובת עם מתחם
            r'מתחם\s+[0-9א-ת]+',
            # כתובת עם רחוב ומספר בית
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+',
            # כתובת עם רחוב ומספר דירה
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*/\s*\d+',
            # כתובת עם רחוב ומספר קומה
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*קומה\s*\d+',
            # כתובת עם רחוב ומספר כניסה
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*כניסה\s*\d+',
            # כתובת עם רחוב ומספר בניין
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*בניין\s*\d+',
            # כתובת עם רחוב ומספר דירה
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*דירה\s*\d+',
            # כתובת עם רחוב ומספר ת.ד.
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*ת\.?ד\.?\s*\d+',
            # כתובת עם רחוב ומספר מיקוד
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*מיקוד\s*\d+',
            # כתובת עם רחוב ומספר עיר
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*,\s*[\u0590-\u05FF\s]+',
            # כתובת עם רחוב ומספר עיר ומיקוד
            r'רחוב\s+[\u0590-\u05FF\s]+\s+\d+\s*,\s*[\u0590-\u05FF\s]+\s*מיקוד\s*\d+'
        ]
        
        # חיפוש כתובות
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                address = match.group().strip()
                if address and any(c.isalpha() for c in address):
                    addresses.append(address)
                    self.logger.debug(f"נמצאה כתובת: {address}")
        
        # הסרת כפילויות
        return list(set(addresses))

    def extract_from_xlsx(self, file_path: str) -> List[Contact]:
        """חולץ אנשי קשר מקובץ Excel"""
        try:
            contacts = []
            self.logger.info(f"מתחיל לעבד קובץ Excel: {file_path}")
            wb = load_workbook(file_path, data_only=True)
            
            for sheet_name in wb.sheetnames:
                try:
                    ws = wb[sheet_name]
                    self.logger.info(f"מעבד גיליון: {sheet_name}")
                    
                    # דילוג על גיליונות ריקים
                    if ws.max_row <= 1 or ws.max_column <= 1:
                        self.logger.debug(f"דילוג על גיליון ריק: {sheet_name}")
                        continue
                        
                    self.logger.info(f"גיליון {sheet_name} מכיל {ws.max_row} שורות ו-{ws.max_column} עמודות")
                    
                    # זיהוי עמודות
                    headers = []
                    header_row = None
                    
                    # חיפוש שורת כותרות בחמש השורות הראשונות
                    for row_idx in range(1, min(6, ws.max_row + 1)):
                        row_values = [str(cell.value).strip() if cell.value else "" for cell in ws[row_idx]]
                        self.logger.debug(f"בודק שורה {row_idx} לכותרות: {row_values}")
                        # בדיקה אם השורה מכילה מספיק מידע להיות שורת כותרות
                        if any(word in " ".join(row_values).lower() for word in ['שם', 'טלפון', 'מייל', 'כתובת', 'תפקיד']):
                            headers = [str(val).lower() for val in row_values]
                            header_row = row_idx
                            self.logger.info(f"נמצאה שורת כותרות בשורה {row_idx}: {headers}")
                            break
                    
                    if not header_row:
                        self.logger.warning(f"לא נמצאה שורת כותרות בגיליון {sheet_name}")
                        continue
                    
                    # מציאת עמודות רלוונטיות
                    name_cols = []
                    phone_cols = []
                    email_cols = []
                    address_cols = []
                    role_cols = []
                    
                    for idx, header in enumerate(headers):
                        header = str(header).lower()
                        # זיהוי עמודות שם
                        if any(word in header for word in ['שם', 'איש קשר', 'נציג', 'חברה', 'שם פרטי', 'שם משפחה', 'שם מלא']):
                            name_cols.append(idx)
                            self.logger.debug(f"נמצאה עמודת שם: {header} (עמודה {idx+1})")
                        # זיהוי עמודות טלפון
                        elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון', 'מס טלפון', 'טלפונים', 'סלולרי', 'נייח']):
                            phone_cols.append(idx)
                            self.logger.debug(f"נמצאה עמודת טלפון: {header} (עמודה {idx+1})")
                        # זיהוי עמודות מייל
                        elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל', 'כתובת מייל', 'אימיילים', '@']):
                            email_cols.append(idx)
                            self.logger.debug(f"נמצאה עמודת מייל: {header} (עמודה {idx+1})")
                        # זיהוי עמודות כתובת
                        elif any(word in header for word in ['כתובת', 'עיר', 'ישוב', 'רחוב', 'בית', 'כתובות', 'מיקוד']):
                            address_cols.append(idx)
                            self.logger.debug(f"נמצאה עמודת כתובת: {header} (עמודה {idx+1})")
                        # זיהוי עמודות תפקיד
                        elif any(word in header for word in ['תפקיד', 'תפקידים', 'תפקידים נוספים', 'משרה', 'תפקיד בחברה']):
                            role_cols.append(idx)
                            self.logger.debug(f"נמצאה עמודת תפקיד: {header} (עמודה {idx+1})")
                    
                    # עיבוד כל שורה
                    for row_idx in range(header_row + 1, ws.max_row + 1):
                        row = ws[row_idx]
                        row_values = [str(cell.value).strip() if cell.value else "" for cell in row]
                        
                        # דילוג על שורות ריקות
                        if not any(row_values):
                            continue
                            
                        contact = Contact(source_file=f"{file_path} - {sheet_name}")
                        self.logger.debug(f"מעבד שורה {row_idx}")
                        
                        # חילוץ שם
                        names = []
                        for col in name_cols:
                            if col < len(row_values) and row_values[col]:
                                names.append(row_values[col])
                        if names:
                            # אם יש כמה שמות, משתמשים בארוך ביותר
                            name = max(names, key=len)
                            if name and any(c.isalpha() for c in name):
                                contact.name = name
                                self.logger.debug(f"נמצא שם: {name}")
                        
                        # חילוץ טלפונים
                        for col in phone_cols:
                            if col < len(row_values) and row_values[col]:
                                phones = self._extract_phones(row_values[col])
                                for phone in phones:
                                    contact.add_phone(phone)
                        
                        # חילוץ אימיילים
                        for col in email_cols:
                            if col < len(row_values) and row_values[col]:
                                emails = self._extract_emails(row_values[col])
                                for email in emails:
                                    contact.add_email(email)
                        
                        # חילוץ כתובות
                        for col in address_cols:
                            if col < len(row_values) and row_values[col]:
                                addresses = self._extract_addresses(row_values[col])
                                for address in addresses:
                                    contact.add_address(address)
                        
                        # חילוץ תפקיד
                        roles = []
                        for col in role_cols:
                            if col < len(row_values) and row_values[col]:
                                roles.append(row_values[col])
                        if roles:
                            # אם יש כמה תפקידים, משתמשים בארוך ביותר
                            role = max(roles, key=len)
                            if role:
                                contact.add_role(role)
                        
                        # בדיקת כל הטקסט בשורה לחיפוש מידע נוסף
                        row_text = ' '.join(row_values)
                        if not contact.phones:
                            for phone in self._extract_phones(row_text):
                                contact.add_phone(phone)
                        if not contact.emails:
                            for email in self._extract_emails(row_text):
                                contact.add_email(email)
                        if not contact.addresses:
                            for address in self._extract_addresses(row_text):
                                contact.add_address(address)
                        
                        if contact.is_valid():
                            contacts.append(contact)
                            self.logger.debug(f"נוסף איש קשר: {contact.name}")
                        else:
                            self.logger.debug(f"נדחה איש קשר לא תקין בשורה {row_idx}")
                            
                except Exception as e:
                    self.logger.error(f"שגיאה בעיבוד גיליון {sheet_name}: {str(e)}")
                    continue
            
            self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בקובץ {file_path}")
            return contacts

        except Exception as e:
            self.logger.error(f"שגיאה בעיבוד קובץ Excel {file_path}: {str(e)}")
            return []

    def extract_from_doc(self, file_path: str) -> List[Contact]:
        """חולץ אנשי קשר מקובץ Word"""
        try:
            contacts = []
            
            # עבור קבצי DOCX
            if file_path.lower().endswith('.docx'):
                doc = Document(file_path)
                self.logger.info(f"מעבד קובץ Word: {file_path}")
                
                # עיבוד טבלאות
                for table_idx, table in enumerate(doc.tables):
                    try:
                        self.logger.debug(f"מעבד טבלה {table_idx + 1}")
                        
                        # קבלת כותרות מהשורה הראשונה
                        headers = []
                        header_row = None
                        
                        # חיפוש שורת כותרות בשלוש השורות הראשונות
                        for row_idx in range(min(3, len(table.rows))):
                            row_values = [cell.text.strip() for cell in table.rows[row_idx].cells]
                            # בדיקה אם השורה מכילה מספיק מידע להיות שורת כותרות
                            if any(word in " ".join(row_values).lower() for word in ['שם', 'טלפון', 'מייל', 'כתובת', 'תפקיד']):
                                headers = [str(val).lower() for val in row_values]
                                header_row = row_idx
                                self.logger.debug(f"נמצאה שורת כותרות בשורה {row_idx + 1}: {headers}")
                                break
                        
                        if not header_row:
                            self.logger.debug(f"לא נמצאה שורת כותרות בטבלה {table_idx + 1}, מנסה לחלץ מידע מכל השורות")
                            header_row = -1
                        
                        # מציאת עמודות רלוונטיות
                        name_cols = []
                        phone_cols = []
                        email_cols = []
                        address_cols = []
                        role_cols = []
                        
                        if header_row >= 0:
                            for idx, header in enumerate(headers):
                                header = str(header).lower()
                                # זיהוי עמודות שם
                                if any(word in header for word in ['שם', 'איש קשר', 'נציג', 'חברה', 'שם פרטי', 'שם משפחה', 'שם מלא']):
                                    name_cols.append(idx)
                                    self.logger.debug(f"נמצאה עמודת שם: {header} (עמודה {idx+1})")
                                # זיהוי עמודות טלפון
                                elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון', 'מס טלפון', 'טלפונים', 'סלולרי', 'נייח']):
                                    phone_cols.append(idx)
                                    self.logger.debug(f"נמצאה עמודת טלפון: {header} (עמודה {idx+1})")
                                # זיהוי עמודות מייל
                                elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל', 'כתובת מייל', 'אימיילים', '@']):
                                    email_cols.append(idx)
                                    self.logger.debug(f"נמצאה עמודת מייל: {header} (עמודה {idx+1})")
                                # זיהוי עמודות כתובת
                                elif any(word in header for word in ['כתובת', 'עיר', 'ישוב', 'רחוב', 'בית', 'כתובות', 'מיקוד']):
                                    address_cols.append(idx)
                                    self.logger.debug(f"נמצאה עמודת כתובת: {header} (עמודה {idx+1})")
                                # זיהוי עמודות תפקיד
                                elif any(word in header for word in ['תפקיד', 'תפקידים', 'תפקידים נוספים', 'משרה', 'תפקיד בחברה']):
                                    role_cols.append(idx)
                                    self.logger.debug(f"נמצאה עמודת תפקיד: {header} (עמודה {idx+1})")
                        
                        # עיבוד כל שורה
                        start_row = header_row + 1 if header_row >= 0 else 0
                        for row_idx in range(start_row, len(table.rows)):
                            row = table.rows[row_idx]
                            row_values = [cell.text.strip() for cell in row.cells]
                            
                            # דילוג על שורות ריקות
                            if not any(row_values):
                                continue
                                
                            contact = Contact(source_file=f"{file_path} - טבלה {table_idx + 1}")
                            self.logger.debug(f"מעבד שורה {row_idx + 1}")
                            
                            if header_row >= 0:
                                # חילוץ לפי עמודות מזוהות
                                # חילוץ שם
                                names = []
                                for col in name_cols:
                                    if col < len(row_values) and row_values[col]:
                                        names.append(row_values[col])
                                if names:
                                    # אם יש כמה שמות, משתמשים בארוך ביותר
                                    name = max(names, key=len)
                                    if name and any(c.isalpha() for c in name):
                                        contact.name = name
                                        self.logger.debug(f"נמצא שם: {name}")
                                
                                # חילוץ טלפונים
                                for col in phone_cols:
                                    if col < len(row_values) and row_values[col]:
                                        phones = self._extract_phones(row_values[col])
                                        for phone in phones:
                                            contact.add_phone(phone)
                                
                                # חילוץ אימיילים
                                for col in email_cols:
                                    if col < len(row_values) and row_values[col]:
                                        emails = self._extract_emails(row_values[col])
                                        for email in emails:
                                            contact.add_email(email)
                                
                                # חילוץ כתובות
                                for col in address_cols:
                                    if col < len(row_values) and row_values[col]:
                                        addresses = self._extract_addresses(row_values[col])
                                        for address in addresses:
                                            contact.add_address(address)
                                
                                # חילוץ תפקיד
                                roles = []
                                for col in role_cols:
                                    if col < len(row_values) and row_values[col]:
                                        roles.append(row_values[col])
                                if roles:
                                    # אם יש כמה תפקידים, משתמשים בארוך ביותר
                                    role = max(roles, key=len)
                                    if role:
                                        contact.add_role(role)
                            
                            # חילוץ מידע מכל התאים בשורה
                            row_text = ' '.join(row_values)
                            
                            # אם אין שם, מנסה למצוא בטקסט המלא
                            if not contact.name:
                                name_and_role = self._extract_name_and_role(row_text)
                                if name_and_role:
                                    name, role = name_and_role
                                    contact.name = name
                                    if role:
                                        contact.add_role(role)
                            
                            # חיפוש מידע נוסף בטקסט המלא
                            if not contact.phones:
                                for phone in self._extract_phones(row_text):
                                    contact.add_phone(phone)
                            if not contact.emails:
                                for email in self._extract_emails(row_text):
                                    contact.add_email(email)
                            if not contact.addresses:
                                for address in self._extract_addresses(row_text):
                                    contact.add_address(address)
                            
                            if contact.is_valid():
                                contacts.append(contact)
                                self.logger.debug(f"נוסף איש קשר: {contact.name}")
                            else:
                                self.logger.debug(f"נדחה איש קשר לא תקין בשורה {row_idx + 1}")
                    
                    except Exception as e:
                        self.logger.error(f"שגיאה בעיבוד טבלה {table_idx + 1}: {str(e)}")
                        continue
                
                # עיבוד פסקאות
                paragraphs_text = []
                current_paragraph = []
                
                for para in doc.paragraphs:
                    text = para.text.strip()
                    if text:
                        current_paragraph.append(text)
                    elif current_paragraph:
                        paragraphs_text.append(' '.join(current_paragraph))
                        current_paragraph = []
                
                if current_paragraph:
                    paragraphs_text.append(' '.join(current_paragraph))
                
                # עיבוד כל פסקה
                for para_text in paragraphs_text:
                    try:
                        # חיפוש אנשי קשר בפסקה
                        para_contacts = self._extract_contacts_from_text(para_text)
                        if para_contacts:
                            contacts.extend(para_contacts)
                            self.logger.debug(f"נמצאו {len(para_contacts)} אנשי קשר בפסקה")
                    except Exception as e:
                        self.logger.error(f"שגיאה בעיבוד פסקה: {str(e)}")
                        continue
            
            self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בקובץ {file_path}")
            return contacts

        except Exception as e:
            self.logger.error(f"שגיאה בחילוץ אנשי קשר מקובץ Word {file_path}: {str(e)}")
            return []

    def extract_from_pdf(self, file_path: str) -> List[Contact]:
        """חולץ אנשי קשר מקובץ PDF"""
        try:
            contacts = []
            self.logger.info(f"מעבד קובץ PDF: {file_path}")
            
            # שימוש ב-pdfplumber
            with pdfplumber.open(file_path) as pdf:
                # עיבוד כל עמוד
                for page_num, page in enumerate(pdf.pages, 1):
                    try:
                        self.logger.debug(f"מעבד עמוד {page_num}")
                        
                        # חילוץ טקסט
                        page_text = page.extract_text()
                        if page_text:
                            # חיפוש אנשי קשר בטקסט
                            text_contacts = self._extract_contacts_from_text(page_text)
                            if text_contacts:
                                contacts.extend(text_contacts)
                                self.logger.debug(f"נמצאו {len(text_contacts)} אנשי קשר בטקסט בעמוד {page_num}")
                        
                        # חילוץ טבלאות
                        tables = page.extract_tables()
                        for table_idx, table in enumerate(tables):
                            try:
                                self.logger.debug(f"מעבד טבלה {table_idx + 1} בעמוד {page_num}")
                                
                                if not table or not any(table):
                                    continue
                                
                                # קבלת כותרות מהשורה הראשונה
                                headers = []
                                header_row = None
                                
                                # חיפוש שורת כותרות בשלוש השורות הראשונות
                                for row_idx in range(min(3, len(table))):
                                    row_values = [str(cell).strip() if cell else "" for cell in table[row_idx]]
                                    # בדיקה אם השורה מכילה מספיק מידע להיות שורת כותרות
                                    if any(word in " ".join(row_values).lower() for word in ['שם', 'טלפון', 'מייל', 'כתובת', 'תפקיד']):
                                        headers = [str(val).lower() if val else "" for val in row_values]
                                        header_row = row_idx
                                        self.logger.debug(f"נמצאה שורת כותרות בשורה {row_idx + 1}: {headers}")
                                        break
                                
                                if not header_row:
                                    self.logger.debug(f"לא נמצאה שורת כותרות בטבלה {table_idx + 1}, מנסה לחלץ מידע מכל השורות")
                                    header_row = -1
                                
                                # מציאת עמודות רלוונטיות
                                name_cols = []
                                phone_cols = []
                                email_cols = []
                                address_cols = []
                                role_cols = []
                                
                                if header_row >= 0:
                                    for idx, header in enumerate(headers):
                                        header = str(header).lower()
                                        # זיהוי עמודות שם
                                        if any(word in header for word in ['שם', 'איש קשר', 'נציג', 'חברה', 'שם פרטי', 'שם משפחה', 'שם מלא']):
                                            name_cols.append(idx)
                                            self.logger.debug(f"נמצאה עמודת שם: {header} (עמודה {idx+1})")
                                        # זיהוי עמודות טלפון
                                        elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון', 'מס טלפון', 'טלפונים', 'סלולרי', 'נייח']):
                                            phone_cols.append(idx)
                                            self.logger.debug(f"נמצאה עמודת טלפון: {header} (עמודה {idx+1})")
                                        # זיהוי עמודות מייל
                                        elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל', 'כתובת מייל', 'אימיילים', '@']):
                                            email_cols.append(idx)
                                            self.logger.debug(f"נמצאה עמודת מייל: {header} (עמודה {idx+1})")
                                        # זיהוי עמודות כתובת
                                        elif any(word in header for word in ['כתובת', 'עיר', 'ישוב', 'רחוב', 'בית', 'כתובות', 'מיקוד']):
                                            address_cols.append(idx)
                                            self.logger.debug(f"נמצאה עמודת כתובת: {header} (עמודה {idx+1})")
                                        # זיהוי עמודות תפקיד
                                        elif any(word in header for word in ['תפקיד', 'תפקידים', 'תפקידים נוספים', 'משרה', 'תפקיד בחברה']):
                                            role_cols.append(idx)
                                            self.logger.debug(f"נמצאה עמודת תפקיד: {header} (עמודה {idx+1})")
                                
                                # עיבוד כל שורה
                                start_row = header_row + 1 if header_row >= 0 else 0
                                for row_idx in range(start_row, len(table)):
                                    row_values = [str(cell).strip() if cell else "" for cell in table[row_idx]]
                                    
                                    # דילוג על שורות ריקות
                                    if not any(row_values):
                                        continue
                                        
                                    contact = Contact(source_file=f"{file_path} - עמוד {page_num}, טבלה {table_idx + 1}")
                                    self.logger.debug(f"מעבד שורה {row_idx + 1}")
                                    
                                    if header_row >= 0:
                                        # חילוץ לפי עמודות מזוהות
                                        # חילוץ שם
                                        names = []
                                        for col in name_cols:
                                            if col < len(row_values) and row_values[col]:
                                                names.append(row_values[col])
                                        if names:
                                            # אם יש כמה שמות, משתמשים בארוך ביותר
                                            name = max(names, key=len)
                                            if name and any(c.isalpha() for c in name):
                                                contact.name = name
                                                self.logger.debug(f"נמצא שם: {name}")
                                        
                                        # חילוץ טלפונים
                                        for col in phone_cols:
                                            if col < len(row_values) and row_values[col]:
                                                phones = self._extract_phones(row_values[col])
                                                for phone in phones:
                                                    contact.add_phone(phone)
                                        
                                        # חילוץ אימיילים
                                        for col in email_cols:
                                            if col < len(row_values) and row_values[col]:
                                                emails = self._extract_emails(row_values[col])
                                                for email in emails:
                                                    contact.add_email(email)
                                        
                                        # חילוץ כתובות
                                        for col in address_cols:
                                            if col < len(row_values) and row_values[col]:
                                                addresses = self._extract_addresses(row_values[col])
                                                for address in addresses:
                                                    contact.add_address(address)
                                        
                                        # חילוץ תפקיד
                                        roles = []
                                        for col in role_cols:
                                            if col < len(row_values) and row_values[col]:
                                                roles.append(row_values[col])
                                        if roles:
                                            # אם יש כמה תפקידים, משתמשים בארוך ביותר
                                            role = max(roles, key=len)
                                            if role:
                                                contact.add_role(role)
                                    
                                    # חילוץ מידע מכל התאים בשורה
                                    row_text = ' '.join(row_values)
                                    
                                    # אם אין שם, מנסה למצוא בטקסט המלא
                                    if not contact.name:
                                        name_and_role = self._extract_name_and_role(row_text)
                                        if name_and_role:
                                            name, role = name_and_role
                                            contact.name = name
                                            if role:
                                                contact.add_role(role)
                                    
                                    # חיפוש מידע נוסף בטקסט המלא
                                    if not contact.phones:
                                        for phone in self._extract_phones(row_text):
                                            contact.add_phone(phone)
                                    if not contact.emails:
                                        for email in self._extract_emails(row_text):
                                            contact.add_email(email)
                                    if not contact.addresses:
                                        for address in self._extract_addresses(row_text):
                                            contact.add_address(address)
                                    
                                    if contact.is_valid():
                                        contacts.append(contact)
                                        self.logger.debug(f"נוסף איש קשר: {contact.name}")
                                    else:
                                        self.logger.debug(f"נדחה איש קשר לא תקין בשורה {row_idx + 1}")
                            
                            except Exception as e:
                                self.logger.error(f"שגיאה בעיבוד טבלה {table_idx + 1} בעמוד {page_num}: {str(e)}")
                                continue
                    
                    except Exception as e:
                        self.logger.error(f"שגיאה בעיבוד עמוד {page_num}: {str(e)}")
                        continue
            
            self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בקובץ {file_path}")
            return contacts

        except Exception as e:
            self.logger.error(f"שגיאה בחילוץ אנשי קשר מקובץ PDF {file_path}: {str(e)}")
            return []

    def _extract_contacts_from_text(self, text: str) -> List[Contact]:
        """חולץ אנשי קשר מטקסט גולמי"""
        contacts = []
        
        # פיצול הטקסט לשורות
        lines = text.split('\n')
        current_contact = None
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_contact and current_contact.is_valid():
                    contacts.append(current_contact)
                    self.logger.debug(f"נוסף איש קשר מהטקסט: {current_contact.name}")
                current_contact = None
                continue
            
            # ניסיון לחלץ שם ותפקיד
            name_and_role = self._extract_name_and_role(line)
            if name_and_role:
                if current_contact and current_contact.is_valid():
                    contacts.append(current_contact)
                    self.logger.debug(f"נוסף איש קשר מהטקסט: {current_contact.name}")
                name, role = name_and_role
                current_contact = Contact(name=name, source_file=None)
                if role:
                    current_contact.add_role(role)
                    self.logger.debug(f"נוסף תפקיד: {role}")
            
            # חילוץ מידע נוסף
            if current_contact:
                # חילוץ מספרי טלפון
                phones = self._extract_phones(line)
                for phone in phones:
                    current_contact.add_phone(phone)
                
                # חילוץ כתובות מייל
                emails = self._extract_emails(line)
                for email in emails:
                    current_contact.add_email(email)
                
                # חילוץ כתובות
                addresses = self._extract_addresses(line)
                for address in addresses:
                    current_contact.add_address(address)
        
        # הוספת איש הקשר האחרון אם תקין
        if current_contact and current_contact.is_valid():
            contacts.append(current_contact)
            self.logger.debug(f"נוסף איש קשר מהטקסט: {current_contact.name}")
        
        self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בטקסט")
        return contacts

    def load_contacts_from_excel(self, file_path):
        """טוען אנשי קשר מקובץ Excel קיים"""
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            contacts = []
            
            # קבלת כותרות
            headers = [str(cell.value).lower() for cell in ws[1]]
            self.logger.debug(f"כותרות בקובץ: {headers}")
            
            # מציאת עמודות רלוונטיות
            name_col = None
            phone_col = None
            email_col = None
            address_col = None
            role_col = None
            
            for idx, header in enumerate(headers, 1):
                if any(word in header for word in ['שם', 'איש קשר', 'נציג', 'חברה', 'שם פרטי', 'שם משפחה', 'שם מלא']):
                    name_col = idx
                elif any(word in header for word in ['טלפון', 'נייד', 'טל', 'פלאפון', 'מס טלפון', 'טלפונים']):
                    phone_col = idx
                elif any(word in header for word in ['מייל', 'אימייל', 'דוא"ל', 'דואל', 'כתובת מייל', 'אימיילים']):
                    email_col = idx
                elif any(word in header for word in ['כתובת', 'עיר', 'ישוב', 'רחוב', 'בית', 'כתובות']):
                    address_col = idx
                elif any(word in header for word in ['תפקיד', 'תפקידים', 'תפקידים נוספים']):
                    role_col = idx
            
            self.logger.debug(f"עמודות זוהו: שם={name_col}, טלפון={phone_col}, מייל={email_col}, כתובת={address_col}, תפקיד={role_col}")
            
            # עיבוד כל שורה
            for row in ws.iter_rows(min_row=2):
                contact = Contact(source_file=file_path)
                
                # חילוץ שם
                if name_col and row[name_col-1].value:
                    name = str(row[name_col-1].value).strip()
                    if name and any(c.isalpha() for c in name):
                        contact.name = name
                        self.logger.debug(f"נמצא שם: {name}")
                
                # חילוץ טלפון
                if phone_col and row[phone_col-1].value:
                    phones = str(row[phone_col-1].value).split('; ')
                    for phone in phones:
                        contact.add_phone(phone)
                
                # חילוץ מייל
                if email_col and row[email_col-1].value:
                    emails = str(row[email_col-1].value).split('; ')
                    for email in emails:
                        contact.add_email(email)
                
                # חילוץ כתובת
                if address_col and row[address_col-1].value:
                    addresses = str(row[address_col-1].value).split('; ')
                    for address in addresses:
                        contact.add_address(address)
                
                # חילוץ תפקיד
                if role_col and row[role_col-1].value:
                    role = str(row[role_col-1].value).strip()
                    if role:
                        contact.add_role(role)
                        self.logger.debug(f"נמצא תפקיד: {role}")
                
                if contact.is_valid():
                    contacts.append(contact)
                    self.logger.debug(f"נוסף איש קשר: {contact.name}")
            
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
            name_key = contact.name.lower().strip() if contact.name else ""
            phone_key = list(contact.phones)[0] if contact.phones else ""
            email_key = list(contact.emails)[0] if contact.emails else ""
            key = f"{name_key}_{phone_key}_{email_key}"
            
            if key in unique_contacts:
                # מיזוג אנשי קשר זהים
                unique_contacts[key].merge(contact)
                self.logger.debug(f"מיזוג איש קשר: {contact.name}")
            else:
                unique_contacts[key] = contact
                self.logger.debug(f"הוספת איש קשר חדש: {contact.name}")
        
        result = list(unique_contacts.values())
        self.logger.info(f"נמצאו {len(result)} אנשי קשר ייחודיים מתוך {len(contacts)} אנשי קשר")
        return result
        
    def extract_contacts(self, file_path):
        """מחלץ אנשי קשר מקובץ בודד"""
        extension = file_path.split('.')[-1].lower()
        
        if extension in ['txt']:
            return self._extract_from_txt(file_path)
        elif extension in ['docx']:
            return self._extract_from_docx(file_path)
        elif extension in ['pdf']:
            return self._extract_from_pdf(file_path)
        elif extension in ['xlsx', 'xls']:
            return self._extract_from_excel(file_path)
        else:
            raise ValueError(f"סוג הקובץ {extension} אינו נתמך")

    def scan_directory(self, directory_path):
        """סורק תיקייה ומחלץ אנשי קשר מכל הקבצים המתאימים"""
        all_contacts = []
        for root, _, files in os.walk(directory_path):
            for file in files:
                if file.split('.')[-1].lower() in ['txt', 'docx', 'pdf', 'xlsx', 'xls']:
                    file_path = os.path.join(root, file)
                    try:
                        contacts = self.extract_contacts(file_path)
                        all_contacts.extend(contacts)
                    except Exception as e:
                        print(f"שגיאה בקריאת הקובץ {file}: {str(e)}")
        return all_contacts

    def remove_duplicates(self, contacts):
        """מסיר כפילויות מרשימת אנשי הקשר"""
        df = pd.DataFrame(contacts)
        if not df.empty:
            df = df.drop_duplicates()
        return df.to_dict('records')

    def save_contacts_to_excel(self, contacts, output_path):
        """שומר את אנשי הקשר לקובץ אקסל"""
        df = pd.DataFrame(contacts)
        df.to_excel(output_path, index=False)

    def _extract_from_txt(self, file_path):
        """מחלץ אנשי קשר מקובץ טקסט"""
        contacts = []
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
            phones = self.phone_pattern.findall(text)
            emails = self.email_pattern.findall(text)
            
            for phone in phones:
                contacts.append({'phone': phone})
            for email in emails:
                contacts.append({'email': email})
        return contacts

    def _extract_from_docx(self, file_path):
        """מחלץ אנשי קשר מקובץ Word"""
        contacts = []
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            phones = self.phone_pattern.findall(paragraph.text)
            emails = self.email_pattern.findall(paragraph.text)
            
            for phone in phones:
                contacts.append({'phone': phone})
            for email in emails:
                contacts.append({'email': email})
        return contacts

    def _extract_from_pdf(self, file_path):
        try:
            if file_ext in ['.xlsx', '.xls']:
                self.logger.info(f"מחלץ מקובץ Excel: {file_path}")
                contacts = self.extract_from_xlsx(file_path)
            elif file_ext in ['.docx', '.doc']:
                self.logger.info(f"מחלץ מקובץ Word: {file_path}")
                contacts = self.extract_from_doc(file_path)
            elif file_ext in ['.pdf']:
                self.logger.info(f"מחלץ מקובץ PDF: {file_path}")
                contacts = self.extract_from_pdf(file_path)
            else:
                self.logger.warning(f"סוג קובץ לא נתמך: {file_ext} בקובץ {file_path}")
                # ניסיון לזהות את סוג הקובץ לפי התוכן
                contacts = []
                errors = []
                
                try:
                    # ניסיון לפתוח כקובץ Excel
                    contacts = self.extract_from_xlsx(file_path)
                    if contacts:
                        self.logger.info(f"זוהה כקובץ Excel למרות הסיומת: {file_path}")
                        return contacts
                except Exception as excel_error:
                    errors.append(f"שגיאת Excel: {str(excel_error)}")
                    
                try:
                    # ניסיון לפתוח כקובץ Word
                    contacts = self.extract_from_doc(file_path)
                    if contacts:
                        self.logger.info(f"זוהה כקובץ Word למרות הסיומת: {file_path}")
                        return contacts
                except Exception as word_error:
                    errors.append(f"שגיאת Word: {str(word_error)}")
                    
                try:
                    # ניסיון לפתוח כקובץ PDF
                    contacts = self.extract_from_pdf(file_path)
                    if contacts:
                        self.logger.info(f"זוהה כקובץ PDF למרות הסיומת: {file_path}")
                        return contacts
                except Exception as pdf_error:
                    errors.append(f"שגיאת PDF: {str(pdf_error)}")
                    
                self.logger.warning(f"סוג קובץ לא נתמך: {file_path}")
                self.logger.error(f"שגיאות בניסיון לזהות את סוג הקובץ: {', '.join(errors)}")
                return []
            
            self.logger.info(f"נמצאו {len(contacts)} אנשי קשר בקובץ {file_path}")
            return contacts
            
        except Exception as e:
            self.logger.error(f"שגיאה בחילוץ אנשי קשר מהקובץ {file_path}: {str(e)}")
            return []

    def save_contacts_to_excel(self, contacts: List[Contact], output_path: str) -> None:
        """שומר אנשי קשר לקובץ Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            
            # יצירת חוברת עבודה חדשה
            wb = Workbook()
            ws = wb.active
            ws.title = 'אנשי קשר'
            
            # הגדרת סגנונות
            header_font = Font(bold=True)
            header_fill = PatternFill(start_color='D7E4BC', end_color='D7E4BC', fill_type='solid')
            cell_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            cell_alignment = Alignment(wrap_text=True, vertical='top')
            
            # כתיבת כותרות
            headers = ['שם', 'תפקיד', 'טלפון', 'אימייל', 'כתובת', 'קובץ מקור']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = cell_border
                cell.alignment = cell_alignment
            
            # כתיבת נתונים
            for row, contact in enumerate(contacts, 2):
                ws.cell(row=row, column=1, value=contact.name)
                ws.cell(row=row, column=2, value=contact.role)
                ws.cell(row=row, column=3, value='; '.join(contact.phones))
                ws.cell(row=row, column=4, value='; '.join(contact.emails))
                ws.cell(row=row, column=5, value='; '.join(contact.addresses))
                ws.cell(row=row, column=6, value=contact.source_file)
                
                # החלת סגנונות על תאים
                for col in range(1, 7):
                    cell = ws.cell(row=row, column=col)
                    cell.border = cell_border
                    cell.alignment = cell_alignment
            
            # התאמת רוחב עמודות
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = min(adjusted_width, 100)
            
            # הקפאת שורת הכותרות
            ws.freeze_panes = 'A2'
            
            # שמירת הקובץ
            wb.save(output_path)
            self.logger.info(f"נשמרו {len(contacts)} אנשי קשר לקובץ {output_path}")
            
        except Exception as e:
            self.logger.error(f"שגיאה בשמירת אנשי קשר לקובץ Excel: {str(e)}")
            raise 