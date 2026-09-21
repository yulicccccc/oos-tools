import fitz, os, sys, shutil

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
pdf_path = os.path.join(desktop_dir, "OOS-261186 EM SMO 116A Air 07MAY2026 - EM.pdf")

doc = fitz.open(pdf_path)

for page in doc:
    for w in page.widgets():
        name = w.field_name
        
        if name == "Text Field3":
            w.field_value = "Simin Mohammad\n(Weekly Active Air Sampling Analyst)\n\nSophia Santamaria\n(Weekly Active Air Sampling Plate Reader)"
            w.text_fontsize = 7.0
            w.update()
            print("Updated Text Field3")
            
        elif name == "Text Field8":
            w.field_value = "MICRO-SOP-2"
            w.update()
            print("Updated Text Field8")
            
        elif name == "Text Field9":
            w.field_value = "23-Jul-2026"
            w.update()
            print("Updated Text Field9")
            
        elif name == "Text Field10":
            w.field_value = "16"
            w.update()
            print("Updated Text Field10")
            
        elif name == "Text Field11":
            w.field_value = "Action level: >= 10 CFU/Plate"
            w.text_fontsize = 6.5
            w.update()
            print("Updated Text Field11")
            
        elif name == "Text Field13":
            w.field_value = "Yes, analysts Simin Mohammad and Sophia Santamaria were interviewed comprehensively."
            w.text_fontsize = 7.5
            w.update()
            print("Updated Text Field13")
            
        elif name == "Text Field43":
            w.field_value = "Incubator E001034 (Sensor E001501)\n\nIncubator E001031 (Sensor E001505)"
            w.text_fontsize = 7.0
            w.update()
            print("Updated Text Field43")
            
        elif name == "Text Field44":
            w.field_value = "Aug 2026 / Feb 2027\n\nAug 2026 / Feb 2027"
            w.text_fontsize = 7.0
            w.update()
            print("Updated Text Field44")
            
        elif name == "Text Field49":
            t49 = w.field_value
            t49 = t49.replace("document.The", "document. The")
            t49 = t49.replace("system.Active", "system. Active")
            t49 = t49.replace("Facility.The", "Facility. The")
            t49 = t49.replace("Sophia Sanatamaria", "Sophia Santamaria")
            t49 = t49.replace("clean room, were", "clean room were")
            t49 = t49.replace("(Gram (+) cocci)To observe", "(Gram (+) cocci). To observe")
            t49 = t49.replace("weekly monitoring environmental monitoring", "weekly environmental monitoring")
            w.field_value = t49
            w.text_fontsize = 8.0
            w.update()
            print("Updated Text Field49")
            
        elif name == "Text Field50":
            t50 = w.field_value
            t50 = t50.replace("MICRO-SOP--9", "MICRO-SOP-9")
            t50 = t50.replace("on 31-May-2026 It was documented", "on 31-May-2026. It was documented")
            t50 = t50.replace("suite 115", "Suite 116")
            t50 = t50.replace("Suite 115", "Suite 116")
            
            old_rc_substr = "may be attributed to a potential analyst error."
            new_rc_substr = (
                "is considered an isolated, transient environmental event. "
                "The analyst adhered to all aseptic protocols and standard operating procedures with no deviations noted during interview. "
                "Laboratory error is not considered the assignable cause."
            )
            if old_rc_substr in t50:
                t50 = t50.replace(old_rc_substr, new_rc_substr)
            
            w.field_value = t50
            w.text_fontsize = 8.0
            w.update()
            print("Updated Text Field50")

temp_out = os.path.join(desktop_dir, "temp_updated_261186.pdf")
doc.save(temp_out)
doc.close()

# Overwrite target PDF
shutil.move(temp_out, pdf_path)
print(f"Successfully updated all fields in: {pdf_path}")

# Also update desktop OOS-261186.pdf
shutil.copyfile(pdf_path, os.path.join(desktop_dir, "OOS-261186.pdf"))
print("Updated OOS-261186.pdf as well.")
