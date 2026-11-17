#logic: from xc take agent and search it in xnt, if is not there the go to xno
#       take the level of the agent from xc; search that level on xno (2 = krediseguros nivel 2)
#       take the product from xad and search it in that level
#       fill the colums Emisión, Renovación and Cambio BR with the values taken fromx xno
#       if the agent appear in xnt; take the product from xad and search it
#       fill the colums Emisión, Renovación and Cambio BR with the values taken from xno

import pandas as pd
import json
import re
import unicodedata
from rapidfuzz import process, fuzz

# ---------------------
# File paths
# ---------------------
file_xad = "reportes_krediseguros.xlsx"
file_xc = "xc_filled.xlsx"
niveles_custom_path = "niveles.json"
niveles_base_path = "niveles1.json"

# ---------------------
# Text Cleaning Functions
# ---------------------
def remove_accents(text: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(c)
    )

def clean_text(text: str) -> str:
    text = re.sub(r"\(.*?\)", "", str(text))  # remove anything in parentheses
    text = remove_accents(text.strip().upper())
    text = re.sub(r"[^A-Z0-9 ]+", "", text)
    text = re.sub(r"\s+", " ", text)
    return text

def drop_duplicate_contracts(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates(subset=["AGENTE", "PRODUCTO"], keep="first").reset_index(drop=True)

# ---------------------
# Load Data
# ---------------------
df_xad = pd.read_excel(file_xad, sheet_name='contratos_comisiones')
df_xc = pd.read_excel(file_xc, sheet_name='AGENTES')

# Normalize column names
df_xad.columns = df_xad.columns.str.strip().str.upper()
df_xc.columns = df_xc.columns.str.strip().str.upper()

# Cleaning function to the "AGENTE" column in df_xc
if "AGENTE" in df_xc.columns:
    df_xc["AGENTE"] = df_xc["AGENTE"].apply(clean_text)

# Rename if necessary
if "NOMBRE AGENTE" in df_xad.columns:
    df_xad.rename(columns={"NOMBRE AGENTE": "AGENTE"}, inplace=True)

# Clean AGENTE and PRODUCTO columns
for col in ["AGENTE", "PRODUCTO", "COMPAÑIA"]:
    if col in df_xad.columns:
        df_xad[col] = df_xad[col].apply(clean_text)

# Load JSON data
with open(niveles_custom_path, "r", encoding="utf-8") as f:
    xnt_data = json.load(f)
with open(niveles_base_path, "r", encoding="utf-8") as f:
    xno_data = json.load(f)

# Map level number to label
level_map = {
    "krediseguros nivel 1": "krediseguros nivel 1",
    "krediseguros nivel 2": "krediseguros nivel 2",
    "krediseguros nivel 3": "krediseguros nivel 3"
}

# Add empty columns
for col in ["Emisión", "Renovación", "Cambio BR"]:
    if col not in df_xad.columns:
        df_xad[col] = ""

# Build agent -> level map
agent_to_level = dict(zip(df_xc["AGENTE"], df_xc["NIVEL"].astype(str)))
print("🔍 agent_to_level size:", len(agent_to_level))
print("🔍 Sample keys:", list(agent_to_level.keys())[:5])


# Preprocess keys for matching
xnt_keys = [clean_text(k) for k in xnt_data.keys()]
agent_level_keys = [clean_text(k) for k in agent_to_level.keys()]
xnt_key_map = dict(zip(xnt_keys, xnt_data.keys()))
agent_level_key_map = dict(zip(agent_level_keys, agent_to_level.keys()))

# ---------------------
# Main logic
# ---------------------
for idx, row in df_xad.iterrows():
    agent = clean_text(row.get("AGENTE", ""))
    product = clean_text(row.get("PRODUCTO", ""))

    if not agent or not product:
        continue

    emision = renovacion = cambio_br = ""

    # --- Match agent to xnt (custom) ---
    agent_match, agent_score, _ = process.extractOne(agent, xnt_keys, scorer=fuzz.token_set_ratio)

    if agent_score >= 40:
        true_agent_key = xnt_key_map[agent_match]
        agent_dict = xnt_data[true_agent_key]
        for company_name, company_dict in agent_dict.items():
            product_keys = [clean_text(k) for k in company_dict.keys()]
            product_match, product_score, _ = process.extractOne(product, product_keys, scorer=fuzz.token_set_ratio)
            if product_score >= 80:
                true_product_key = list(company_dict.keys())[product_keys.index(product_match)]
                prod_data = company_dict[true_product_key]
                emision = prod_data.get("Emision", "")
                renovacion = prod_data.get("Renovacion", "")
                cambio_br = prod_data.get("Cambio BR", "")
                break

    else:
        # --- Match agent to level-based table (xno) ---
        matched_agent_xc, score_xc, _ = process.extractOne(agent, agent_level_keys, scorer=fuzz.token_set_ratio)
        if score_xc >= 40:
            true_agent_key = agent_level_key_map[matched_agent_xc]
            level = agent_to_level[true_agent_key]
            level_key = level_map.get(str(level), "")
            if level_key and level_key in xno_data:
                level_dict = xno_data[level_key]
                product_keys = [clean_text(k) for k in level_dict.keys()]
                product_match, product_score, _ = process.extractOne(product, product_keys, scorer=fuzz.token_set_ratio)
                if product_score >= 70:
                    true_product_key = list(level_dict.keys())[product_keys.index(product_match)]
                    prod_data = level_dict[true_product_key]
                    emision = prod_data.get("Emision", "")
                    renovacion = prod_data.get("Renovacion", "")
                    cambio_br = prod_data.get("Cambio BR", "")

    # Write values
    df_xad.at[idx, "Emisión"] = emision
    df_xad.at[idx, "Renovación"] = renovacion
    df_xad.at[idx, "Cambio BR"] = cambio_br

# ---------------------
# Retry logic for unfilled rows
# ---------------------
for idx, row in df_xad.iterrows():
    if all(str(row[col]).strip() == "" for col in ["Emisión", "Renovación", "Cambio BR"]):
        agent = clean_text(row.get("AGENTE", ""))
        product = clean_text(row.get("PRODUCTO", ""))

        if not agent or not product:
            continue

        matched_agent_xc, score_xc, _ = process.extractOne(agent, agent_level_keys, scorer=fuzz.token_set_ratio)
        if score_xc >= 20:
            true_agent_key = agent_level_key_map[matched_agent_xc]
            level = agent_to_level[true_agent_key]
            level_key = level_map.get(str(level), "")
            if level_key and level_key in xno_data:
                level_dict = xno_data[level_key]
                product_keys = [clean_text(k) for k in level_dict.keys()]
                product_match, product_score, _ = process.extractOne(product, product_keys, scorer=fuzz.token_set_ratio)
                if product_score >= 50:
                    true_product_key = list(level_dict.keys())[product_keys.index(product_match)]
                    prod_data = level_dict[true_product_key]
                    df_xad.at[idx, "Emisión"] = prod_data.get("Emision", "")
                    df_xad.at[idx, "Renovación"] = prod_data.get("Renovacion", "")
                    df_xad.at[idx, "Cambio BR"] = prod_data.get("Cambio BR", "")

# ---------------------
# Remove duplicates after all matching
# ---------------------
df_xad = drop_duplicate_contracts(df_xad)

# ---------------------
# Optional: Log unmatched rows
# ---------------------
unmatched = df_xad[
    df_xad[["Emisión", "Renovación", "Cambio BR"]]
      .apply(lambda v: all(str(x).strip() == "" for x in v), axis=1)
]
unmatched_count = len(unmatched)

print("⚠️ Unmatched AGENTES sample:", unmatched["AGENTE"].value_counts().head(10))
print("⚠️ Unmatched PRODUCTOS sample:", unmatched["PRODUCTO"].value_counts().head(10))
print("Count of unmatched", unmatched_count)
unmatched.to_excel("unmatched_debug.xlsx", index=False)


# ---------------------
# Export
# ---------------------
df_xad.to_excel("xad_filled.xlsx", index=False)
print("✅ Finished filling commissions.")
