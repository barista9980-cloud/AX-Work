"""
Official Contract Note Generator Module for FoxConnect Real Estate Management
Standard Master Template: 99_템플릿_양식\부동산_계약노트_표준양식.docx
"""

import os
import shutil
import docx
from docx.oxml import OxmlElement

DEFAULT_MASTER_TEMPLATE = r"G:\내 드라이브\[FoxConnect]\[총무]업무\99_템플릿_양식\부동산_계약노트_표준양식.docx"

def generate_contract_note_from_master(output_path, master_template_path=DEFAULT_MASTER_TEMPLATE):
    """
    Clones official Standard Master Template 1:1 to guarantee 100% exact visual layout.
    """
    if not os.path.exists(master_template_path):
        raise FileNotFoundError(f"Master template not found: {master_template_path}")
        
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 1:1 File Copy from 99_템플릿_양식 Master Template
    shutil.copy2(master_template_path, output_path)
    
    # Open doc and enforce row-level cantSplit page boundaries
    doc = docx.Document(output_path)
    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(OxmlElement('w:cantSplit'))
            if r_idx == 0:
                trPr.append(OxmlElement('w:tblHeader'))
                
    doc.save(output_path)
    print(f"[ContractNoteGenerator] 1:1 Cloned from 99_템플릿_양식 Template -> {output_path}")
    return output_path