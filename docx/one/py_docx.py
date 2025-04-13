from docx import Document

doc = Document("/Users/gavinzhang/workspaces/pdf_table_detector/docx/one/twotables_1.docx")
for idx, table in enumerate(doc.tables):
    table_xml = table._tbl.xml
    print(f"Table {idx} XML:")
    print(table_xml)