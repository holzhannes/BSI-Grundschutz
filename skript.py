import pdfplumber
import pandas as pd
import os
import re

pdf_dir = "./Einzelne_PDF"

rows = []

for filename in os.listdir(pdf_dir):
    if filename.endswith(".pdf"):
        filepath = os.path.join(pdf_dir, filename)
        print(f"Verarbeite: {filename}")
        
        match = re.match(r"([A-Z]+)\.([0-9.]+)", filename)
        if not match:
            continue
        bereich = match.group(1)
        baustein = f"{bereich}.{match.group(2)}"
        
        bausteinbezeichnung = None

        with pdfplumber.open(filepath) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            if first_page_text:
                for l in first_page_text.split('\n'):
                    if l.strip().startswith("Baustein:"):
                        bausteinbezeichnung = l.replace("Baustein:", "").strip()
                        break

            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for i, line in enumerate(lines):
                    # determine requirements
                    match = re.match(r"([A-Z]+\.\d+\.\d+\.A\d+)\s+(.*)", line)
                    if match:
                        anforderungsbez = match.group(1)
                        titel_raw = match.group(2).strip()
                        anforderung = anforderungsbez.split('.')[-1]

                        # extract requirements
                        art_match = re.search(r"\((B|M|H)\)", titel_raw)
                        anforderungsart = art_match.group(1) if art_match else ""
                        titel = re.sub(r"\s*\([BMH]\)", "", titel_raw).strip()

                        # description from following cells
                        beschreibung = ""
                        c5_id = ""
                        for j in range(i + 1, min(i + 6, len(lines))):
                            next_line = lines[j].strip()
                            if next_line == "" or re.match(r"[A-Z]+\.\d+\.\d+\.A\d+", next_line):
                                break
                            beschreibung += " " + next_line
                            if "C5" in next_line or "-" in next_line:
                                c5_id = next_line

                        rows.append({
                            "PDF": filename,
                            "Bereich": bereich,
                            "Baustein": baustein,
                            "Bausteinbezeichnung": bausteinbezeichnung,
                            "Anforderung": anforderung,
                            "Anforderungsbezeichnung": anforderungsbez,
                            "Titel": titel,
                            "Anforderungsart": anforderungsart,
                            "Beschreibung": beschreibung.strip(),
                            "C5-ID": c5_id
                        })

# Dataframe to excel
df = pd.DataFrame(rows)
df.to_excel("zieltabelle_anforderungen.xlsx", index=False)
print(f"Exportiert: zieltabelle_anforderungen.xlsx mit {len(df)} Einträgen")
