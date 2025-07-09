import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import csv
import os
import pandas as pd
import re
import numpy as np, cv2

# Optional: Tesseract-Pfad
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# OCR-Korrekturen
def clean_ocr_line(line):
    replacements = {
        "Oltemperatur": "Öltemperatur",
        "rd": "[°C]",
        "cea": "[°C]",
        "Drehzahl": "Drehzahl [rpm]",
        ";": "",
        ".": "",
        "Wirkungsgrad": "Wirkungsgrad [eta]",
        "Öltemperatur Ike}":"Öltemperatur [°C]",
        "Öltemperatur rc":"Öltemperatur [°C]",
        '"Drehmoment [N]' : "Drehmoment [Nm]",
        '"Drehmoment IN]' : "Drehmoment [Nm]",
        "[Drehmoment IN]" : "Drehmoment [Nm]",
        "Öltemperatur ke}" : "Öltemperatur [°C]",
        "Öltemperatur Pc" : "Öltemperatur [°C]",
    }
    for wrong, right in replacements.items():
        line = line.replace(wrong, right)
    return line

# Komma-Ergänzungen für bestimmte Messwerte
def fix_missing_commas(line):
    if "Leckölvolumenstrom [l/min]" in line or "Wirkungsgrad [eta]" in line or "Leckölvolumenstrom" in line or "Wirkungsgrad" in line:
        parts = line.split()
        corrected = [parts[0]]
        for val in parts[1:]:
            if val.isdigit() and len(val) >= 2:
                corrected.append(val[:-1] + "," + val[-1])
            else:
                corrected.append(val)
        return corrected
    return line.split()

# === Vorschlag für deine preprocess_image(image)-Funktion:
def preprocess_image(image):


    image = image.convert("RGB")

    from PIL import ImageEnhance, ImageFilter
    import numpy as np, cv2

    image_cv = np.array(image)
    image_cv = cv2.cvtColor(image_cv, cv2.COLOR_RGB2GRAY)

    image = Image.fromarray(image_cv)
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)

    return image


def restructure_and_save_csv(rows, output_csv="ocr_gesamt_output_neu3.csv"):

    df = pd.DataFrame(rows)
    df.columns = ["Messgröße", "Wert1", "Wert2", "Wert3", "Wert4", "Wert5", "Wert6", "Seriennummer", "Datum_Zeit", "Test"]

    records = []

    grouped = df.groupby(["Seriennummer", "Datum_Zeit", "Test"])

    for (sn, dt, test), group in grouped:
        group_unique = group.drop_duplicates(subset="Messgröße", keep="first")
        messdaten = group_unique.set_index("Messgröße").iloc[:, 0:6].T
        messdaten.columns.name = None
        # Zeilen mit echten Hauptdruckwerten behalten
        if "Hauptdruck [bar]" in messdaten.columns:
            druckwerte = pd.to_numeric(messdaten["Hauptdruck [bar]"], errors="coerce")
            messdaten = messdaten[druckwerte.notna()]
            n = len(messdaten)
        else:
            n = len(messdaten)


        # === Druckstufenliste nach Länge und Inhalt bestimmen ===
        druckstufen = []
        print("Druckstufen: ", n)
        print("Messdaten: ", messdaten)
        
        if n == 6:
            # Prüfe 250 vs 300 anhand von Messwerten
            if "Hauptdruck [bar]" in messdaten.columns:
                try:
                    druckwerte = pd.to_numeric(messdaten["Hauptdruck [bar]"], errors="coerce")
                    hat_250 = ((druckwerte >= 245) & (druckwerte <= 255)).any()
                    hat_300 = ((druckwerte >= 295) & (druckwerte <= 305)).any()
                    if hat_250 and not hat_300:
                        druckstufen = ["30bar", "100bar", "200bar", "250bar", "", "30bar_2"]
                    else:
                        druckstufen = ["30bar", "100bar", "200bar", "300bar", "350bar", "30bar_2"]
                except:
                    druckstufen = ["30bar", "100bar", "200bar", "300bar", "350bar", "30bar_2"]
            else:
                druckstufen = ["30bar", "100bar", "200bar", "300bar", "350bar", "30bar_2"]

        elif n == 5:
            druckstufen = ["30bar", "100bar", "200bar", "250bar", "30bar_2"]

        else:
            # generisch, z. B. bei 4 oder 7 Zeilen
            druckstufen = [f"Stufe {i+1}" for i in range(n)]

        messdaten["Seriennummer"] = sn
        messdaten["Datum_Zeit"] = dt

        # Aufteilung von Datum und Uhrzeit
        if "_" in dt:
            datum_part, zeit_part = dt.split("_")
            zeit_formatiert = zeit_part.replace(".", ":")
        else:
            datum_part = dt
            zeit_formatiert = ""

        messdaten["Datum"] = datum_part
        messdaten["Uhrzeit"] = zeit_formatiert

        messdaten["Test"] = test
        messdaten["Druckstufen"] = druckstufen

        # Sortierreihenfolge der Druckstufen (für spätere Sortierung)
        druckstufen_map = {
            "30bar": 1,
            "100bar": 2,
            "200bar": 3,
            "250bar": 4,
            "300bar": 5,
            "350bar": 6,
            "30bar_2": 7
        }
        messdaten["Druckstufe_Nr"] = messdaten["Druckstufen"].map(druckstufen_map)


        records.append(messdaten)

    df_result = pd.concat(records, ignore_index=True)

    fixed_columns = ["Seriennummer", "Datum_Zeit", "Druckstufen"]
    messgroessen = [col for col in df_result.columns if col not in fixed_columns + ["Test"]]
    df_result = df_result[fixed_columns + messgroessen + ["Test"]]


    df_result.to_csv(output_csv, sep=";", index=False, encoding="utf-8-sig")
    print(f"📄 Umstrukturierte CSV gespeichert unter: {output_csv}")

# Hauptverarbeitung: alle Bilder im Ordner durchgehen

    print("\n📋 Gesammelte OCR-Tabelle:\n")
    #print(df.to_string(index=False, header=False))


def process_folder(folder_path, output_csv="ocr_gesamt_output3.csv"):
    all_rows = []
    filename_to_sn = {}
    sn_counter = {}

    # === Schritt 1: Alle Dateinamen analysieren und eindeutige Seriennummern erzeugen ===
    filenames = [fn for fn in os.listdir(folder_path) if fn.lower().endswith((".jpg", ".jpeg", ".png", ".tif"))]
    for filename in sorted(filenames):
        match = re.search(r"\[(\d{2}-[\w\d\s]+)\]\[(\d{2}\.\d{2}\.\d{4}_[\d\.]+)\]\[(.*?)\]", filename)
        if match:
            seriennummer, datum, pruefung = match.groups()
            count = sn_counter.get(seriennummer, 0)
            unique_sn = seriennummer if count == 0 else f"{seriennummer}_{count}"
            sn_counter[seriennummer] = count + 1
            filename_to_sn[filename] = (unique_sn, datum, pruefung)
        else:
            filename_to_sn[filename] = ("NA", "NA", "NA")

    # === Schritt 2: Bildverarbeitung + OCR ===
    for filename in filenames:
        image_path = os.path.join(folder_path, filename)
        print(f"🔍 Verarbeite: {image_path}")
        seriennummer, datum, pruefung = filename_to_sn.get(filename, ("NA", "NA", "NA"))

        try:
            image = Image.open(image_path)
            image = preprocess_image(image)
            ocr_result = pytesseract.image_to_string(image, lang='deu')

            lines = ocr_result.strip().split('\n')
            cleaned_lines = [clean_ocr_line(line.strip()) for line in lines if line.strip()]
            tabbed_rows = [fix_missing_commas(line) for line in cleaned_lines]

            for row in tabbed_rows:
                row2 = row[0] + " " + row[1]
                row.pop(0)
                row.pop(0)
                row.insert(0, row2)
                row += [seriennummer, datum, pruefung]
                if len(row) < 10:
                    row.insert(5, "")
                all_rows.append(row)

        except Exception as e:
            print(f"⚠️ Fehler bei {filename}: {e}")

    # === CSV-Speicherung ===
    with open(output_csv, "w", newline='', encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerows(all_rows)

    restructure_and_save_csv(all_rows, output_csv="ocr_gesamt_output_neu3.csv")
    print(f"\n✅ Gesamt-CSV gespeichert unter: {output_csv}")

    df = pd.DataFrame(all_rows)
    print("\n📋 Gesammelte OCR-Tabelle:\n")
    #print(df.to_string(index=False, header=False))

# === AUSFÜHREN ===
#ordner_pfad = r"C:\Users\RobTech\Desktop\Grundpumpe" #zuHause
ordner_pfad = r"C:\Users\rbuechner\Desktop\Grundpumpe" #Arbeit
process_folder(ordner_pfad)
