import pdfplumber
import pandas as pd
import os

# Pfad zu den entpackten PDF-Bausteinen
pdf_dir = "./Einzelne_PDF"  

rows = []

for filename in os.listdir(pdf_dir):
    if filename.endswith(".pdf"):
        filepath = os.path.join(pdf_dir, filename)
        print(f"Verarbeite: {filename}")
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for line in lines:
                    if ".A" in line:
                        parts = line.strip().split(" ", 1)
                        if len(parts) == 2:
                            anf_num, anf_text = parts
                            rows.append({
                                "PDF": filename,
                                "Anforderung": anf_num,
                                "Text": anf_text
                            })

df = pd.DataFrame(rows)
df.to_csv("bsi_anforderungen_2023.csv", index=False, encoding='utf-8')
print("Exportiert: bsi_anforderungen_2023.csv")
