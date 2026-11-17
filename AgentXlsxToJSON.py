import pandas as pd
import json
import re
import unicodedata

excel_path = "TABLA COMISIONES WORKER FINAL (2).xlsx"
output_path = "niveles.json"

# Load Excel and get sheet names
xl = pd.ExcelFile(excel_path)
agent_data = {}

def remove_accents(text: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(c)
    )

for sheet in xl.sheet_names:
    df = xl.parse(sheet, header=None)
    
    # Find agent name (first non-empty, all-uppercase cell in first 5 rows)
    agent_name = None
    for i in range(min(5, len(df))):  # only loop over existing rows
        if i >= len(df):
            break
        for cell in df.iloc[i].dropna():
            match = re.match(r"^[A-ZÑÁÉÍÓÚÜ ]{3,}$", str(cell).strip())
            if match:
                agent_name = match.group(0).strip()
                break
        if agent_name:
            break

    if not agent_name:
        print(f"⚠️ Could not determine agent name in sheet '{sheet}'")
        continue

    print(f"📄 Parsing sheet: {sheet} — Agent: {agent_name}")
    agent_data[agent_name] = {}

    current_company = None
    skip_keywords = ['krediseguros nivel']

    for idx in range(len(df)):
        row = df.iloc[idx]

        # Check if this row has the 'PRODUCTO' header
        if row.astype(str).str.contains('PRODUCTO', case=False).any():
            # Backtrack to find the real company name (skipping "krediseguros nivel X")
            back_idx = idx - 1
            while back_idx >= 0:
                row_candidate = df.iloc[back_idx].dropna()
                if not row_candidate.empty:
                    candidate = str(row_candidate.iloc[0]).strip()
                    if not any(keyword in candidate.lower() for keyword in skip_keywords):
                        current_company = candidate
                        break
                back_idx -= 1
            if not current_company:
                current_company = f"Unknown_{idx}"

            # Get column indexes
            header_row = row.fillna("")
            col_map = {col.strip().lower(): i for i, col in enumerate(header_row)}
            try:
                prod_col = col_map.get("producto", 1)
                emi_col = col_map.get("emision", 2)
                ren_col = col_map.get("renovacion", 3)
                camb_col = col_map.get("cambio br", 4)
            except Exception as e:
                print(f"⚠️ Column mapping failed in {sheet}, row {idx}: {e}")
                continue

            # Read the data below this row
            data_rows = df.iloc[idx + 1:].reset_index(drop=True)
            company_dict = {}

            for _, prod_row in data_rows.iterrows():
                producto = remove_accents(str(prod_row[prod_col]).strip().upper())
                producto = re.sub(r"[^A-Z0-9 ]+", "", producto)
                if not producto or producto.lower().startswith('nivel'):
                    break  # End of current table

                emision = str(prod_row[emi_col]).strip() if emi_col in prod_row else ""
                renovacion = str(prod_row[ren_col]).strip() if ren_col in prod_row else ""
                cambio_br = str(prod_row[camb_col]).strip() if camb_col in prod_row else ""

                company_dict[producto] = {
                    "Emision": emision,
                    "Renovacion": renovacion,
                    "Cambio BR": cambio_br
                }

            # Save this company to the agent data
            if current_company not in agent_data[agent_name]:
                agent_data[agent_name][current_company] = {}

            agent_data[agent_name][current_company].update(company_dict)

# Save to JSON
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(agent_data, f, ensure_ascii=False, indent=2)

print("✅ niveles.json saved successfully.")