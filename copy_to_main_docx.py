import shutil

src = r"C:\Users\User\OneDrive\바탕 화면\부동산_계약관리노트_강남_도곡로1길23_정의양식.docx"
dst = r"C:\Users\User\OneDrive\바탕 화면\부동산_계약관리노트_강남_도곡로1길23.docx"

try:
    shutil.copyfile(src, dst)
    print("Successfully updated main docx file!")
except Exception as e:
    print(f"Could not overwrite main docx (file may be open): {e}")
