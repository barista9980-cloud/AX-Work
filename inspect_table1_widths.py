import docx

doc = docx.Document(r"C:\Users\User\OneDrive\바탕 화면\부동산_계약관리노트_402_403호.docx")

t1 = doc.tables[1]
print("Table 1 cols:", len(t1.columns))
for c_i, cell in enumerate(t1.rows[0].cells):
    print(f"Col {c_i}: width={cell.width.pt if cell.width else 'None'}, text='{cell.text}'")
