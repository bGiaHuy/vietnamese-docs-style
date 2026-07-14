import sys
import argparse
from docx import Document
from docx.shared import Cm

def validate_doc(filepath, profile="administrative"):
    print(f"Validating {filepath} for profile: {profile}...")
    try:
        doc = Document(filepath)
    except Exception as e:
        print(f"ERROR: Could not open document. {e}")
        return False

    errors = []

    # 1. Page Setup Validation
    for i, section in enumerate(doc.sections):
        # A4 size in EMUs: width = 7937500, height = 11226800. Allowing a small margin of error.
        width_cm = section.page_width.cm
        height_cm = section.page_height.cm
        
        if abs(width_cm - 21.0) > 0.5 or abs(height_cm - 29.7) > 0.5:
            errors.append(f"Section {i}: Page size is not A4 (Found {width_cm:.2f}x{height_cm:.2f} cm)")
            
        left_cm = section.left_margin.cm
        right_cm = section.right_margin.cm
        top_cm = section.top_margin.cm
        bottom_cm = section.bottom_margin.cm
        
        if not (2.9 <= left_cm <= 3.6):
            errors.append(f"Section {i}: Left margin {left_cm:.2f}cm is out of 3.0-3.5cm range.")
        if not (1.4 <= right_cm <= 2.1):
            errors.append(f"Section {i}: Right margin {right_cm:.2f}cm is out of 1.5-2.0cm range.")
        if not (1.9 <= top_cm <= 2.6):
            errors.append(f"Section {i}: Top margin {top_cm:.2f}cm is out of 2.0-2.5cm range.")
        if not (1.9 <= bottom_cm <= 2.6):
            errors.append(f"Section {i}: Bottom margin {bottom_cm:.2f}cm is out of 2.0-2.5cm range.")

    # 2. Text Validation
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        if "[CẦN BỔ SUNG" in text:
            errors.append(f"Paragraph {i}: Unresolved placeholder found: {text.strip()[:50]}...")
            
        if profile in ["administrative", "minutes-administrative"]:
            if "•" in text:
                errors.append(f"Paragraph {i}: Bullet dot (•) used in administrative profile.")
                
        # Check direct font formatting if present
        for r in p.runs:
            if r.font.color and r.font.color.rgb and str(r.font.color.rgb) != '000000':
                errors.append(f"Paragraph {i}: Non-black color found ({r.font.color.rgb}).")
            if r.font.name and r.font.name != 'Times New Roman':
                errors.append(f"Paragraph {i}: Font is {r.font.name} instead of Times New Roman.")

    # 3. Table Validation
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                for p_idx, p in enumerate(cell.paragraphs):
                    text = p.text
                    if "[CẦN BỔ SUNG" in text:
                        errors.append(f"Table {t_idx} Row {r_idx} Col {c_idx}: Unresolved placeholder.")
                    if profile in ["administrative", "minutes-administrative"]:
                        if "•" in text:
                            errors.append(f"Table {t_idx} Row {r_idx} Col {c_idx}: Bullet dot (•) used.")
                            
                # Check shading for administrative
                if profile in ["administrative", "minutes-administrative"]:
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    shading = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
                    if shading is not None:
                        val = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        fill = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
                        if (val != 'clear' and val is not None) or (fill != 'auto' and fill != '000000' and fill is not None):
                             errors.append(f"Table {t_idx} Row {r_idx} Col {c_idx}: Shading found in administrative document.")

    if errors:
        print(f"\n❌ Validation failed with {len(errors)} errors:")
        for err in errors:
            print(f"  - {err}")
        return False
        
    print("\n✅ Document passed structural validation.")
    return True

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Validate DOCX files against Vietnamese standards.")
    parser.add_argument("filepath", help="Path to the .docx file")
    parser.add_argument("--profile", default="administrative", help="Document profile (administrative, academic, etc.)")
    args = parser.parse_args()
    
    success = validate_doc(args.filepath, args.profile)
    sys.exit(0 if success else 1)
