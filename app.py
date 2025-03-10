from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import tempfile
import logging
from contact_extractor import ContactExtractor
from config import Config
import socket
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.config.from_object(Config)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'

# הגדרת לוגר
logging.basicConfig(
    level=app.config['LOG_LEVEL'],
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(app.config['LOG_FILE']),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# וודא שתיקיות קיימות
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'doc', 'docx'}

def allowed_file(filename):
    """בודק אם הקובץ מורשה"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """דף הבית"""
    try:
        logger.info("מישהו ניגש לדף הבית")
        return render_template('index.html')
    except Exception as e:
        logger.error(f"שגיאה בטעינת דף הבית: {str(e)}")
        return f"שגיאה בטעינת הדף: {str(e)}", 500

@app.route('/extract', methods=['POST'])
def extract_contacts():
    if 'files' not in request.files:
        return jsonify({'success': False, 'error': 'No files uploaded'})
    
    files = request.files.getlist('files')
    all_contacts = []
    
    try:
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                # חילוץ אנשי קשר מהקובץ
                contacts = extract_contacts_from_file(filepath)
                all_contacts.extend(contacts)
                
                # מחיקת הקובץ לאחר העיבוד
                os.remove(filepath)
        
        # הסרת כפילויות
        unique_contacts = remove_duplicates(all_contacts)
        
        # יצירת קובץ Excel
        output_file = create_excel_output(unique_contacts)
        
        return jsonify({
            'success': True,
            'contacts': unique_contacts[:5],  # שליחת 5 אנשי הקשר הראשונים לתצוגה מקדימה
            'download_url': f'/download/{os.path.basename(output_file)}'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(
        os.path.join(app.config['DOWNLOAD_FOLDER'], filename),
        as_attachment=True,
        download_name=f'contacts_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )

def extract_contacts_from_file(filepath):
    """חילוץ אנשי קשר מקובץ"""
    contacts = []
    ext = filepath.split('.')[-1].lower()
    
    if ext in ['xlsx', 'xls']:
        df = pd.read_excel(filepath)
        # הוסף כאן את הלוגיקה לחילוץ אנשי קשר מ-Excel
    elif ext in ['doc', 'docx']:
        # הוסף כאן את הלוגיקה לחילוץ אנשי קשר מ-Word
        pass
    
    return contacts

def remove_duplicates(contacts):
    """הסרת כפילויות מרשימת אנשי הקשר"""
    # הוסף כאן את הלוגיקה להסרת כפילויות
    return contacts

def create_excel_output(contacts):
    """יצירת קובץ Excel מעוצב עם אנשי הקשר"""
    filename = f'contacts_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    
    df = pd.DataFrame(contacts)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Contacts')
    
    # עיצוב הקובץ
    workbook = writer.book
    worksheet = writer.sheets['Contacts']
    
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'bg_color': '#4F81BD',
        'font_color': 'white'
    })
    
    # הגדרת רוחב עמודות ועיצוב כותרות
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        worksheet.set_column(col_num, col_num, 15)
    
    writer.close()
    return output_path

def is_port_available(port):
    """בודק אם הפורט פנוי"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('0.0.0.0', port))
            return True
        except OSError:
            return False

if __name__ == '__main__':
    # יצירת תיקיות נדרשות
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)
    
    try:
        # מתחיל מפורט 5001
        port = 5001
        while port < 5100:
            if is_port_available(port):
                break
            port += 1
        
        if port >= 5100:
            print("לא נמצא פורט פנוי בטווח 5001-5099")
            exit(1)
            
        print(f"מפעיל את השרת בפורט {port}")
        print(f"גש לאפליקציה בכתובת: http://localhost:{port}")
        
        app.run(host='0.0.0.0', port=port, debug=True)
    except Exception as e:
        print(f"שגיאה בהפעלת השרת: {str(e)}") 