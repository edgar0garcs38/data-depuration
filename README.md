# Commission Table Builder

Turns the brokerage's human-formatted commission tables (Excel sheets with merged headers and one block per insurer) into structured JSON. It then uses that JSON to automatically fill in the commission rates for every agent's contracts.

## What it does
- **`xlsxToJSON.py`:** parses the `NIVELES` sheet (commission rates by agent level) into JSON
- **`AgentXlsxToJSON.py`:** parses one sheet per *special* agent, detecting the agent name, each insurer block and its `PRODUCTO` table automatically
- **`AgentDepurationDf_Xc.py`:** cleans the agent list (accents, spacing, duplicates)
- **`DataDepuration.py`:** for each contract, looks up the agent (special table first, then level table) and the product, then fills the **Issuance**, **Renewal** and **Plan change** commission columns

## Tech stack
Python · pandas · JSON · Unicode normalization

## Run it
```bash
pip install -r requirements.txt
python xlsxToJSON.py
python AgentXlsxToJSON.py
python DataDepuration.py
```

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
