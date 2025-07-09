import os
import csv
import re
import numpy as np
import cv2
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox

# Optional: Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# --- OCR helpers from original script ---

def clean_ocr_line(line):
    replacements = {
        "Oltemperatur": "Öltemperatur",
        "rd": "[°C]",
        "cea": "[°C]",
        "Drehzahl": "Drehzahl [rpm]",
        ";": "",
        ".": "",
        "Wirkungsgrad": "Wirkungsgrad [eta]",
        "Öltemperatur Ike}": "Öltemperatur [°C]",
        "Öltemperatur rc": "Öltemperatur [°C]",
        '"Drehmoment [N]' : "Drehmoment [Nm]",
        '"Drehmoment IN]' : "Drehmoment [Nm]",
        "[Drehmoment IN]" : "Drehmoment [Nm]",
        "Öltemperatur ke}" : "Öltemperatur [°C]",
        "Öltemperatur Pc" : "Öltemperatur [°C]",
    }
    for wrong, right in replacements.items():
        line = line.replace(wrong, right)
    return line

def fix_missing_commas(line):
    if (
        "Leckölvolumenstrom [l/min]" in line
        or "Wirkungsgrad [eta]" in line
        or "Leckölvolumenstrom" in line
        or "Wirkungsgrad" in line
    ):
        parts = line.split()
        corrected = [parts[0]]
        for val in parts[1:]:
            if val.isdigit() and len(val) >= 2:
                corrected.append(val[:-1] + "," + val[-1])
            else:
                corrected.append(val)
        return corrected
    return line.split()

def preprocess_image(image):
    image = image.convert("RGB")
    image_cv = np.array(image)
    image_cv = cv2.cvtColor(image_cv, cv2.COLOR_RGB2GRAY)
    image = Image.fromarray(image_cv)
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    return image

def restructure_and_save_csv(rows, output_csv="ocr_gesamt_output_neu3.csv"):
    df = pd.DataFrame(rows)
    df.columns = [
        "Messgröße",
        "Wert1",
        "Wert2",
        "Wert3",
        "Wert4",
        "Wert5",
        "Wert6",
        "Seriennummer",
        "Datum_Zeit",
        "Test",
    ]
    records = []
    grouped = df.groupby(["Seriennummer", "Datum_Zeit", "Test"])
    for (sn, dt, test), group in grouped:
        group_unique = group.drop_duplicates(subset="Messgröße", keep="first")
        messdaten = group_unique.set_index("Messgröße").iloc[:, 0:6].T
        messdaten.columns.name = None
        if "Hauptdruck [bar]" in messdaten.columns:
            druckwerte = pd.to_numeric(messdaten["Hauptdruck [bar]"], errors="coerce")
            messdaten = messdaten[druckwerte.notna()]
            n = len(messdaten)
        else:
            n = len(messdaten)
        druckstufen = []
        if n == 6:
            if "Hauptdruck [bar]" in messdaten.columns:
                try:
                    druckwerte = pd.to_numeric(
                        messdaten["Hauptdruck [bar]"], errors="coerce"
                    )
                    hat_250 = ((druckwerte >= 245) & (druckwerte <= 255)).any()
                    hat_300 = ((druckwerte >= 295) & (druckwerte <= 305)).any()
                    if hat_250 and not hat_300:
                        druckstufen = [
                            "30bar",
                            "100bar",
                            "200bar",
                            "250bar",
                            "",
                            "30bar_2",
                        ]
                    else:
                        druckstufen = [
                            "30bar",
                            "100bar",
                            "200bar",
                            "300bar",
                            "350bar",
                            "30bar_2",
                        ]
                except Exception:
                    druckstufen = [
                        "30bar",
                        "100bar",
                        "200bar",
                        "300bar",
                        "350bar",
                        "30bar_2",
                    ]
            else:
                druckstufen = [
                    "30bar",
                    "100bar",
                    "200bar",
                    "300bar",
                    "350bar",
                    "30bar_2",
                ]
        elif n == 5:
            druckstufen = ["30bar", "100bar", "200bar", "250bar", "30bar_2"]
        else:
            druckstufen = [f"Stufe {i+1}" for i in range(n)]
        messdaten["Seriennummer"] = sn
        messdaten["Datum_Zeit"] = dt
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
        druckstufen_map = {
            "30bar": 1,
            "100bar": 2,
            "200bar": 3,
            "250bar": 4,
            "300bar": 5,
            "350bar": 6,
            "30bar_2": 7,
        }
        messdaten["Druckstufe_Nr"] = messdaten["Druckstufen"].map(druckstufen_map)
        records.append(messdaten)
    df_result = pd.concat(records, ignore_index=True)
    fixed_columns = ["Seriennummer", "Datum_Zeit", "Druckstufen"]
    messgroessen = [col for col in df_result.columns if col not in fixed_columns + ["Test"]]
    df_result = df_result[fixed_columns + messgroessen + ["Test"]]
    df_result.to_csv(output_csv, sep=";", index=False, encoding="utf-8-sig")
    print(f"📄 Umstrukturierte CSV gespeichert unter: {output_csv}")

def process_folder(folder_path, output_csv="ocr_gesamt_output3.csv"):
    all_rows = []
    filename_to_sn = {}
    sn_counter = {}
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
    for filename in filenames:
        image_path = os.path.join(folder_path, filename)
        print(f"🔍 Verarbeite: {image_path}")
        seriennummer, datum, pruefung = filename_to_sn.get(filename, ("NA", "NA", "NA"))
        try:
            image = Image.open(image_path)
            image = preprocess_image(image)
            ocr_result = pytesseract.image_to_string(image, lang="deu")
            lines = ocr_result.strip().split("\n")
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
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerows(all_rows)
    restructure_and_save_csv(all_rows, output_csv="ocr_gesamt_output_neu3.csv")
    print(f"\n✅ Gesamt-CSV gespeichert unter: {output_csv}")
    pd.DataFrame(all_rows)

# --- Simple GUI ---

def choose_input_dir(var):
    directory = filedialog.askdirectory(title="Ordner mit Grunddatenbildern wählen")
    if directory:
        var.set(directory)

def choose_output_dir(var):
    directory = filedialog.askdirectory(title="Ordner für Excel-Ausgabe wählen")
    if directory:
        var.set(directory)

def start_processing(inp_var, out_var):
    in_dir = inp_var.get()
    out_dir = out_var.get()
    if not in_dir:
        messagebox.showerror("Fehler", "Bitte Ordner mit Bilddateien auswählen")
        return
    if not out_dir:
        messagebox.showerror("Fehler", "Bitte Zielordner für Excel-Datei auswählen")
        return
    output_csv = os.path.join(out_dir, "ocr_gesamt_output3.csv")
    process_folder(in_dir, output_csv=output_csv)
    messagebox.showinfo("Fertig", f"Verarbeitung abgeschlossen. Datei gespeichert unter:\n{output_csv}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Projekt 1 - Daten auslesen")
    inp_var = tk.StringVar()
    out_var = tk.StringVar()

    tk.Label(root, text="Ordner mit Grunddatenbildern:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    tk.Entry(root, textvariable=inp_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Wählen", command=lambda: choose_input_dir(inp_var)).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Ordner für Excel-Datei:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    tk.Entry(root, textvariable=out_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Wählen", command=lambda: choose_output_dir(out_var)).grid(row=1, column=2, padx=5, pady=5)

    tk.Button(root, text="Start", command=lambda: start_processing(inp_var, out_var)).grid(row=2, column=0, columnspan=3, pady=10)

    root.mainloop()
