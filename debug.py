from docx import Document

doc = Document('./דוגמאות/אזור פתח תקווה copy.docx')
for table in doc.tables:
    print('\nTable:')
    for row in table.rows:
        print([cell.text for cell in row.cells]) 