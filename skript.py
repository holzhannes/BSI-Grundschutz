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
                    match = re.match(r"([A-Z]+\.\d+\.\d+\.A\d+)\s+(.*)", line)
                    if match:
                        anforderungsbez = match.group(1)
                        titel_raw = match.group(2).strip()
                        anforderung = anforderungsbez.split('.')[-1]

                        # Titel bereinigen (ohne Anforderungsart)
                        titel = re.sub(r"\s*\([BMHS]\)", "", titel_raw).strip()

                        # Anforderungsart aus Titel oder später aus Beschreibung
                        anforderungsart = ""
                        art_match = re.search(r"\((B|M|H|S)\)", titel_raw)
                        if art_match:
                            anforderungsart = art_match.group(1)

                        beschreibung = ""
                        c5_id = ""

                        for j in range(i + 1, len(lines)):
                            next_line = lines[j].strip()
                            if re.match(r"[A-Z]+\.\d+\.\d+\.A\d+", next_line):
                                break
                            if not next_line:
                                continue

                            # C5-ID wie SIM-01 oder ORG-02 erkennen
                            if not c5_id:
                                c5_match = re.search(r"\b([A-Z]{2,4}-\d{2,3})\b", next_line)
                                if c5_match:
                                    c5_id = c5_match.group(1)

                            # Anforderungsart aus Beschreibung erkennen, falls noch nicht vorhanden
                            if not anforderungsart:
                                art_desc_match = re.search(r"\((B|M|H|S)\)", next_line)
                                if art_desc_match:
                                    anforderungsart = art_desc_match.group(1)

                            beschreibung += " " + next_line

                        # Falls Anforderung entfallen ist
                        if "ENTFALLEN" in titel_raw.upper():
                            beschreibung = "Diese Anforderung ist entfallen."

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

# Als Excel speichern
df = pd.DataFrame(rows)
df.to_excel("zieltabelle_anforderungen.xlsx", index=False)
print(f"Exportiert: zieltabelle_anforderungen.xlsx mit {len(df)} Einträgen")
