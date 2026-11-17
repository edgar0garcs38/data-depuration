import pandas as pd
import json
import re
import unicodedata
from rapidfuzz import process, fuzz

file_xc = "CORREGIDO AGENTES REPORTE COMISIONES.xlsx"

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

df_xc = pd.read_excel(file_xc, sheet_name='AGENTES')

df_xc.columns = df_xc.columns.str.strip().str.upper()

if "AGENTE" in df_xc.columns:
    df_xc["AGENTE"] = df_xc["AGENTE"].apply(clean_text)
    
df_xc.to_excel("xc_filled.xlsx", index=False)
print("✅ Finished filling df_xc.")