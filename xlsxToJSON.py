import pandas as pd
import json
import re 
import unicodedata

def remove_accents(text: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(c)
    )

def parse_niveles_sheet(path: str, sheet_name: str = "NIVELES") -> dict:
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    data = {}
    


    for col in range(1, df.shape[1], 5):  # block starts at col=1, skip every 5
        if col + 3 >= df.shape[1]:
            print(f"Skipping col={col} — out of bounds")
            continue

        company = df.iloc[0, col]
        level = df.iloc[1, col]

        if pd.isna(company) or pd.isna(level):
            print(f"Skipping col={col} because company or level is NaN")
            continue

        company = str(company).strip()
        level = str(level).strip()

        if company not in data:
            data[company] = {}
        if level not in data[company]:
            data[company][level] = {}

        print(f"🔍 Parsing block: company={company}, level={level}")

        for row in range(3, df.shape[0]):
            producto = df.iloc[row, col]
            emision = df.iloc[row, col + 1]
            renovacion = df.iloc[row, col + 2]
            cambio_br = df.iloc[row, col + 3]

            if (
                pd.isna(producto)
                or str(producto).strip().lower() in ["producto", ""]
                or (pd.isna(emision) and pd.isna(renovacion) and pd.isna(cambio_br))
            ):
                continue

            producto = remove_accents(str(df.iloc[row, col]).strip().upper())
            producto = re.sub(r"[^A-Z0-9 ]+", "", producto)
            data[company][level][producto] = {
                "Emision": str(emision).strip() if pd.notna(emision) else "",
                "Renovacion": str(renovacion).strip() if pd.notna(renovacion) else "",
                "Cambio BR": str(cambio_br).strip() if pd.notna(cambio_br) else ""
            }

    return data


# === Run and export to JSON ===
if __name__ == "__main__":
    niveles_data = parse_niveles_sheet("TABLA COMISIONES WORKER FINAL (2).xlsx")

    with open("niveles1.json", "w", encoding="utf-8") as f:
        json.dump(niveles_data, f, indent=2, ensure_ascii=False)

    print("✅ niveles1.json saved successfully.")