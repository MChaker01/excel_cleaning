import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import xlrd
from datetime import datetime
import os
import tempfile

# ==========================================
# HELPER FUNCTION: Convert .xls to .xlsx
# ==========================================
def convert_xls_to_xlsx(xls_path, xlsx_path):
    old_wb = xlrd.open_workbook(xls_path, formatting_info=False)
    old_sheet = old_wb.sheet_by_index(0)
    new_wb = Workbook()
    new_sheet = new_wb.active
    for row in range(old_sheet.nrows):
        for col in range(old_sheet.ncols):
            new_sheet.cell(row=row + 1, column=col + 1, value=old_sheet.cell_value(row, col))
    new_wb.save(xlsx_path)

# ==========================================
# SCRIPT 1: GCOM CLEANING
# ==========================================
def process_gcom(input_path, output_path, file_ext):
    if file_ext == ".xls":
        temp_xlsx = input_path + "_converted.xlsx"
        convert_xls_to_xlsx(input_path, temp_xlsx)
        workbook = load_workbook(temp_xlsx)
        os.remove(temp_xlsx)
    else:
        workbook = load_workbook(input_path)
        
    sheet = workbook.active
    sheet.delete_rows(1)

    COL_AGENCE  = 1
    COL_VENDEUR = 7
    COL_TYPE    = 10
    COL_ARTICLE = 19

    for row in range(sheet.max_row, 1, -1):
        article = sheet.cell(row=row, column=COL_ARTICLE).value
        article_str = str(article).strip().lower() if article else ""
        if "remis" in article_str:
            sheet.delete_rows(row)

    for row in range(2, sheet.max_row + 1):
        agence = sheet.cell(row=row, column=COL_AGENCE).value
        vendeur = sheet.cell(row=row, column=COL_VENDEUR).value
        agence_str = str(agence).strip().lower() if agence else ""
        vendeur_str = str(vendeur).strip().lower() if vendeur else ""

        if agence_str == "agence el jadida":
            sheet.cell(row=row, column=COL_TYPE).value = "Gros"
            
        if vendeur_str in ["bichlifin jamal", "lakhouil mohammed"]:
            sheet.cell(row=row, column=COL_TYPE).value = "Détail"
        elif vendeur_str in ["boujgue ahmed g", "idbikich oussam gros", "laghrib el hassane", "aouzal aziz g"]:
            sheet.cell(row=row, column=COL_TYPE).value = "Gros"    

    workbook.save(output_path)

# ==========================================
# SCRIPT 2: ACCENTURE CLEANING
# ==========================================
def normalize_header(header):
    if header is None: return ""
    return str(header).strip().replace("\n", " ").replace("  ", " ").lower()

def find_column_by_contains(sheet, search_text):
    search_text = search_text.lower()
    for col in range(1, sheet.max_column + 1):
        header_value = sheet.cell(row=1, column=col).value
        if search_text in normalize_header(header_value): return col
    return None

def convert_to_number(value):
    if value is None or value == "": return None
    value = str(value).strip().replace(" ", "")
    if "," in value and "." in value:
        return float(value.replace(",", ""))
    if "," in value:
        return float(value.replace(",", "."))
    return float(value)

def process_accenture(input_path, output_path):
    workbook = load_workbook(input_path)
    sheet = workbook.active

    sheet._images = []
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))

    max_col, max_row = sheet.max_column, sheet.max_row
    for col in range(max_col, 0, -1):
        is_blank = True
        for row in range(1, max_row + 1):
            if sheet.cell(row=row, column=col).value not in (None, ""):
                is_blank = False
                break
        if is_blank: sheet.delete_cols(col)

    sheet.delete_rows(1, 11)
    sheet.delete_cols(4)
    sheet.delete_cols(6)
    sheet.delete_cols(10)

    for col in [15, 16, 17]:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value not in (None, ""):
                try: cell.value = convert_to_number(cell.value)
                except: pass

    transaction_date_col = find_column_by_contains(sheet, "transaction date")
    if transaction_date_col:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=transaction_date_col)
            if cell.value not in (None, ""):
                try:
                    real_date = datetime.strptime(str(cell.value).strip()[:8], "%Y%m%d")
                    cell.value = real_date
                    cell.number_format = "dd/mm/yyyy"
                except: pass

    net_price_col = find_column_by_contains(sheet, "net price")
    realisation_col = find_column_by_contains(sheet, "realisation")
    if not realisation_col:
        realisation_col = sheet.max_column + 1
        sheet.cell(row=1, column=realisation_col, value="Realisation")

    for row in range(2, sheet.max_row + 1):
        cell_ht = sheet.cell(row=row, column=net_price_col)
        if isinstance(cell_ht.value, (int, float)):
            new_cell = sheet.cell(row=row, column=realisation_col, value=round(cell_ht.value * 1.2, 2))
            new_cell.number_format = '#,##0.00'

    data = sheet.values
    columns = next(data)
    df = pd.DataFrame(data, columns=columns)
    df.columns = [str(col).strip() if col is not None else "" for col in df.columns]

    salesman_col_name = next((c for c in df.columns if str(c).strip().lower() == "salesman"), None)
    realisation_col_name = next((c for c in df.columns if str(c).strip().lower() == "realisation"), None)
    transaction_no_col_name = next((c for c in df.columns if "transaction no" in str(c).strip().lower()), None)
    transaction_date_col_name = next((c for c in df.columns if "transaction date" in str(c).strip().lower()), None)

    if salesman_col_name and realisation_col_name and transaction_date_col_name:
        df = df[df[salesman_col_name].notna()]
        df[realisation_col_name] = pd.to_numeric(df[realisation_col_name], errors="coerce")
        df = df[df[transaction_date_col_name].notna()]
        df[salesman_col_name] = df[salesman_col_name].astype(str).str.strip().replace({
            "AGAV003-MOHAMED El MADI": "LMADI MOHAMED",
            "AGAV005-ALIAT RACHID": "ALIAT RACHID",
            "AGAV004-Aboubaker Zarbane": "boubaker zarban",
            "AGAV006-MOHAMED LAKHOUIL": "LAKHOUEL MOHAMED",
            "AGAV001-JAMAL BICHLIFEN": "JAMAL BICHLIFEN"
        })

        summary = df.groupby([transaction_date_col_name, salesman_col_name], dropna=False).agg(
            Total_Realisation=(realisation_col_name, "sum"),
            Nombre_BL=(transaction_no_col_name, lambda x: df.loc[x.index][df.loc[x.index, realisation_col_name] > 0][transaction_no_col_name].nunique())
        ).reset_index()

        summary["Total_Realisation"] = summary["Total_Realisation"].round(2)
        summary = summary.rename(columns={transaction_date_col_name: "Date", salesman_col_name: "Salesman"}).sort_values(by=["Date", "Total_Realisation"], ascending=[True, False])

        summary_sheet_name = "Pivot_Summary"
        if summary_sheet_name in workbook.sheetnames: del workbook[summary_sheet_name]
        summary_sheet = workbook.create_sheet(title=summary_sheet_name)

        for col_idx, col_name in enumerate(summary.columns, start=1):
            summary_sheet.cell(row=1, column=col_idx, value=col_name)

        for row_idx, row_data in enumerate(summary.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row_data, start=1):
                cell = summary_sheet.cell(row=row_idx, column=col_idx, value=value)
                if col_idx == 1 and isinstance(value, (datetime, pd.Timestamp)): cell.number_format = "dd/mm/yyyy"
                if col_idx == 3: cell.number_format = '#,##0.00'

        summary_sheet.auto_filter.ref = summary_sheet.dimensions

        for ws in [sheet, summary_sheet]:
            for column_cells in ws.columns:
                max_len = max([len(str(c.value)) for c in column_cells if c.value] + [0])
                ws.column_dimensions[column_cells[0].column_letter].width = min(max_len + 2, 25)

    workbook.save(output_path)

# ==========================================
# SCRIPT 3: ENCOUR-RD CLEANING
# ==========================================
def process_encour(input_path, output_path, file_ext):
    if file_ext == ".xls":
        temp_xlsx = input_path + "_converted.xlsx"
        convert_xls_to_xlsx(input_path, temp_xlsx)
        workbook = load_workbook(temp_xlsx)
        os.remove(temp_xlsx)
    else:
        workbook = load_workbook(input_path)

    sheet = workbook.active
    sheet._images = []
    
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))

    sheet.delete_rows(1, 10)
    max_col, max_row = sheet.max_column, sheet.max_row

    for row in range(1, max_row + 1):
        for col in range(5, max_col + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value not in (None, ""):
                try:
                    cell.value = float(str(cell.value).replace(',', '.'))
                    cell.number_format = '#,##0.00'
                except: pass           

    for row in range(sheet.max_row, 0, -1):
        if str(sheet.cell(row=row, column=2).value).strip().lower() == "total":
            sheet.delete_rows(row)
            break 

    for col in range(sheet.max_column, 0, -1):
        if not sheet.cell(row=1, column=col).value:
            sheet.delete_cols(col)

    for col in range(sheet.max_column, 0, -1):
        is_empty = True
        for row in range(1, sheet.max_row + 1):
            if sheet.cell(row=row, column=col).value not in (None, ""):
                is_empty = False
                break
        if is_empty: sheet.delete_cols(col)

    workbook.save(output_path)

# ==========================================
# SCRIPT 4: ASSABIL CLEANING
# ==========================================
def process_assabil(input_path, output_path):
    workbook = load_workbook(input_path)
    sheet = workbook.active

    sheet._images = []

    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))

    sheet.delete_rows(1, 7)

    max_col = sheet.max_column
    max_row = sheet.max_row

    for col in range(max_col, 0, -1):
        is_blank = True
        for row in range(1, max_row + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value not in (None, ""):
                is_blank = False
                break
        if is_blank:
            sheet.delete_cols(col)

    sheet.delete_cols(4)
    sheet.delete_cols(2)

    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=3).value
        if val == "PV_G_004":
            sheet.cell(row=row, column=3).value = "EL MOUHAMID MUSTAPHA"
        elif val == "CV_D_003":
            sheet.cell(row=row, column=3).value = "ALI AIT OUCHRAA"
        elif val == "CV_D_007":
            sheet.cell(row=row, column=3).value = "AJERRAR ABDELLAH"
        elif val == "CV_D_006":
            sheet.cell(row=row, column=3).value = "BOUBAKER ZARBAN"
        elif val == "CV_D_005":
            sheet.cell(row=row, column=3).value = "LACHGER YOUSSEF"
        elif val == "V003":
            sheet.cell(row=row, column=3).value = "BICHLIFIN JAMAL"

    sheet.insert_cols(1, 2)
    sheet.cell(row=1, column=1).value = "Genre"
    sheet.cell(row=1, column=2).value = "Agence"
    
    sheet.insert_cols(5, 3)
    sheet.cell(row=1, column=5).value = "N°BC / N°RC"
    sheet.cell(row=1, column=6).value = "N°FC"
    sheet.cell(row=1, column=7).value = "Code vendeur"
    
    sheet.insert_cols(11, 5)
    sheet.cell(row=1, column=11).value = "Type Client"
    sheet.cell(row=1, column=12).value = "Secteur"
    sheet.cell(row=1, column=13).value = "Adresse"
    sheet.cell(row=1, column=14).value = "Ville"
    sheet.cell(row=1, column=15).value = "Livreur"
    
    sheet.insert_cols(17, 2)
    sheet.cell(row=1, column=17).value = "Famille"
    sheet.cell(row=1, column=18).value = "Sous Famille"
    
    sheet.insert_cols(21, 2)
    sheet.cell(row=1, column=21).value = "Quantité"
    sheet.cell(row=1, column=22).value = "RTN"
    sheet.cell(row=1, column=25).value = "TX RSE"
    
    sheet.insert_cols(26, 2)
    sheet.cell(row=1, column=26).value = "Remise MT"
    sheet.cell(row=1, column=27).value = "Total Remises"
    
    sheet.insert_cols(29, 1)
    sheet.cell(row=1, column=29).value = " Objectif"

    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=1).value = "CONV"
        sheet.cell(row=row, column=2).value = "Agence Agadir"  
        sheet.cell(row=row, column=5).value = "3924"  
        sheet.cell(row=row, column=6).value = "3900"  
        sheet.cell(row=row, column=7).value = "VAD002"  
        sheet.cell(row=row, column=11).value = "Détail"
        sheet.cell(row=row, column=12).value = "agadir detail"
        sheet.cell(row=row, column=13).value = "centre bigra"
        sheet.cell(row=row, column=14).value = "Agadir"
        sheet.cell(row=row, column=15).value = "HICHAM LAGRAOUI"
        sheet.cell(row=row, column=17).value = sheet.cell(row=row, column=16).value
        sheet.cell(row=row, column=18).value = sheet.cell(row=row, column=16).value
        sheet.cell(row=row, column=21).value = sheet.cell(row=row, column=25).value
        sheet.cell(row=row, column=22).value = 0
        sheet.cell(row=row, column=25).value = 0
        sheet.cell(row=row, column=26).value = 0
        sheet.cell(row=row, column=27).value = 0
        sheet.cell(row=row, column=29).value = 0

    workbook.save(output_path)

# ==========================================
# SCRIPT 5: FERRERO — Accenture CR19 → GCOM
# ==========================================

# Accenture salesman code → GCOM vendor name
# Input format: "AGAV001-JAMAL BICHLIFEN"  →  take name after "-"  →  look up below
_VENDOR_NAME_MAP = {
    "JAMAL BICHLIFEN":   "BICHLIFIN JAMAL",
    "RACHID ALIAT":      "ALIAT RACHID",
    "MOHAMED EL MADI":   "EL MADI MOHAMED",
    "MOHAMMED LAKHOUIL": "LAKHOUIL MOHAMMED",
    "ABOUBAKER ZARBANE": "BOUBAKER ZARBAN",
}

def _resolve_vendor(raw: str) -> str:
    """'AGAV001-JAMAL BICHLIFEN'  →  'BICHLIFIN JAMAL'"""
    if not raw:
        return ""
    dash = raw.find("-")
    name_part = raw[dash + 1:].strip().upper() if dash >= 0 else raw.strip().upper()
    return _VENDOR_NAME_MAP.get(name_part, name_part)

def _parse_accenture_date(raw) -> datetime:
    """Parse Accenture YYYYMMDD integer/string → datetime."""
    s = str(raw).strip()
    if len(s) >= 8:
        return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
    raise ValueError(f"Cannot parse date: {raw!r}")

def _convert_qty(qty: int, uom: str, bl_tu: int, bl_su: int, bl_cu: int) -> int:
    """
    Convert Accenture quantity to GCOM units.
    
    Conversion rules (mirrors transformExcelController.ts):
      TU  →  qty × bl_tu           (carton/transport unit → individual units)
      SU  →  qty × bl_su           (secondary unit → individual units)
                                    if bl_su == 0: no SU level exists, pass qty through
      CU  →  qty                   (consumer unit = individual unit, no conversion)
    """
    uom = uom.strip().upper()
    if uom == "TU":
        return qty * bl_tu if bl_tu > 0 else qty
    if uom == "SU":
        return qty * bl_su if bl_su > 0 else qty
    return qty  # CU

def _load_ferrero_mapping(mapping_path: str) -> dict:
    """
    Load Ferrero_GCOM_ACCENTURE.xlsx into a lookup dict:
      { accenture_code_str → { gcom_code, name, bl_tu, bl_su, bl_cu } }
    """
    df = pd.read_excel(mapping_path, header=0)
    df.columns = ["name", "gcom_code", "accenture_code", "bl_tu", "bl_su", "bl_cu"]
    df["accenture_code"] = df["accenture_code"].astype(str).str.strip()
    df["gcom_code"] = df["gcom_code"].astype(str).str.strip()

    mapping = {}
    for _, row in df.iterrows():
        code = row["accenture_code"]
        if code and code.lower() != "nan":
            mapping[code] = {
                "gcom_code": row["gcom_code"],
                "name":      row["name"],
                "bl_tu":     int(row["bl_tu"]) if pd.notna(row["bl_tu"]) else None,
                "bl_su":     int(row["bl_su"]) if pd.notna(row["bl_su"]) else None,
                "bl_cu":     int(row["bl_cu"]) if pd.notna(row["bl_cu"]) else None,
            }
    return mapping

def process_ferrero_accenture(input_path: str, output_path: str, mapping_path: str) -> dict:
    """
    Transform a Ferrero CR19 Document Listing Excel export from Accenture
    into a GCOM-ready import file.

    Output workbook contains two sheets:
      • "Livraisons par POS"  — 27-column GCOM format (resolved rows only)
      • "Non Résolus"         — rows that could not be mapped, with reason

    Returns { total, resolved, unresolved } for the UI summary.
    """
    # ── 1. Load the Accenture→GCOM article mapping ─────────────────────
    mapping = _load_ferrero_mapping(mapping_path)

    # ── 2. Parse the Accenture Excel — locate header row dynamically ───
    #    CR19 format: rows 1-11 are report metadata, row 12 is the column
    #    header. We scan for the row containing "Product Code" so the code
    #    is resilient to Accenture changing the number of metadata rows.
    wb_in = load_workbook(input_path, read_only=True, data_only=True)
    ws_in = wb_in.active

    header_row_idx = None
    for row in ws_in.iter_rows():
        for cell in row:
            if str(cell.value or "").strip() == "Product Code":
                header_row_idx = cell.row
                break
        if header_row_idx:
            break

    if header_row_idx is None:
        wb_in.close()
        raise ValueError('Format invalide — colonne "Product Code" introuvable.')

    # Build column-name → 0-based-index map from the header row
    header_vals = list(
        ws_in.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=True)
    )[0]
    col_map = {str(v).strip(): i for i, v in enumerate(header_vals) if v is not None}

    required_cols = [
        "Customer Code", "Customer Name", "Transaction No.",
        "Product Code", "Transaction Date", "Qty", "UOM", "Salesman",
    ]
    for col in required_cols:
        if col not in col_map:
            wb_in.close()
            raise ValueError(f'Colonne manquante dans le fichier Accenture: "{col}"')

    # ── 3. Read data rows ───────────────────────────────────────────────
    max_col_needed = max(col_map[c] for c in required_cols)
    raw_rows = []

    for row_vals in ws_in.iter_rows(min_row=header_row_idx + 1, values_only=True):
        # Skip rows that are too short (empty/summary rows at end of file)
        if len(row_vals) <= max_col_needed:
            continue

        product_code = str(row_vals[col_map["Product Code"]] or "").strip()
        qty_raw      = row_vals[col_map["Qty"]]
        uom          = str(row_vals[col_map["UOM"]] or "").strip().upper()

        if not product_code or not qty_raw or not uom:
            continue

        try:
            qty = int(float(str(qty_raw)))
        except (ValueError, TypeError):
            continue

        raw_rows.append({
            "customer_code":    str(row_vals[col_map["Customer Code"]]   or "").strip(),
            "customer_name":    str(row_vals[col_map["Customer Name"]]   or "").strip(),
            "transaction_no":   str(row_vals[col_map["Transaction No."]] or "").strip(),
            "product_code":     product_code,
            "transaction_date": row_vals[col_map["Transaction Date"]],
            "qty":              qty,
            "uom":              uom,
            "salesman":         str(row_vals[col_map["Salesman"]] or "").strip(),
        })

    wb_in.close()

    if not raw_rows:
        raise ValueError("Aucune ligne de données trouvée dans le fichier.")

    # ── 4. Map each row to GCOM format ─────────────────────────────────
    gcom_rows  = []
    unresolved = []

    for row in raw_rows:
        art = mapping.get(row["product_code"])

        if art is None:
            unresolved.append({**row, "reason": "article_not_found"})
            continue

        if art["bl_tu"] is None or art["bl_su"] is None or art["bl_cu"] is None:
            unresolved.append({**row, "reason": "conversion_not_configured"})
            continue

        gcom_qty = _convert_qty(row["qty"], row["uom"], art["bl_tu"], art["bl_su"], art["bl_cu"])

        try:
            date = _parse_accenture_date(row["transaction_date"])
        except Exception:
            date = None

        gcom_rows.append({
            "produit":    art["gcom_code"],
            "date":       date,
            "bl":         row["transaction_no"],
            "client":     row["customer_code"],
            "nom_client": row["customer_name"],
            "qte":        gcom_qty,
            "vendor":     _resolve_vendor(row["salesman"]),
        })

    # ── 5. Build output workbook ────────────────────────────────────────
    wb_out = Workbook()

    # ── Sheet 1: "Livraisons par POS" — 27-column GCOM import format ───
    #    Column layout (0-based index, matches GCOM Livraisons_par_POS exactly):
    #      [0]  Produit    [1]  V        [2]  Date       [3]  BL
    #      [4]  Cd VD      [5]  V        [6]  Client     [7]  Nom Client
    #      [8]  Qte Carton [9]  v        [10] QTE        [11] Remise
    #      [12–23] V×12 (empty separator columns)
    #      [24] Secteur    [25] Livreur  [26] TIMBRE
    ws_gcom = wb_out.active
    ws_gcom.title = "Livraisons par POS"

    GCOM_HEADERS = [
        "Produit", "V", "Date", "BL", "Cd VD", "V",
        "Client", "Nom Client", "Qte Carton", "v", "QTE", "Remise",
        "V", "V", "V", "V", "V", "V", "V", "V", "V", "V", "V", "V",
        "Secteur", "Livreur", "TIMBRE",
    ]
    ws_gcom.append(GCOM_HEADERS)

    header_fill = PatternFill("solid", fgColor="D9E1F2")
    for cell in ws_gcom[1]:
        cell.fill      = header_fill
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for r in gcom_rows:
        out_row = [None] * 27
        out_row[0]  = r["produit"]      # A  Produit
        out_row[2]  = r["date"]         # C  Date
        out_row[3]  = r["bl"]           # D  BL
        out_row[6]  = r["client"]       # G  Client
        out_row[7]  = r["nom_client"]   # H  Nom Client
        out_row[8]  = 0                 # I  Qte Carton
        out_row[10] = r["qte"]          # K  QTE
        out_row[11] = 0                 # L  Remise
        out_row[24] = r["vendor"]       # Y  Secteur
        out_row[25] = r["vendor"]       # Z  Livreur
        out_row[26] = 0                 # AA TIMBRE
        ws_gcom.append(out_row)
        ws_gcom.cell(row=ws_gcom.max_row, column=3).number_format = "DD/MM/YYYY"

    ws_gcom.column_dimensions["A"].width = 16
    ws_gcom.column_dimensions["C"].width = 12
    ws_gcom.column_dimensions["D"].width = 14
    ws_gcom.column_dimensions["G"].width = 14
    ws_gcom.column_dimensions["H"].width = 32
    ws_gcom.column_dimensions["K"].width = 10

    # ── Sheet 2: "Non Résolus" — rows that couldn't be mapped ──────────
    if unresolved:
        ws_ur = wb_out.create_sheet(title="Non Résolus")
        ws_ur.append(["Code Accenture", "Client", "Nom Client", "BL", "Qté", "Unité", "Raison"])

        ur_fill = PatternFill("solid", fgColor="FDEBD0")
        for cell in ws_ur[1]:
            cell.fill      = ur_fill
            cell.font      = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        reason_labels = {
            "article_not_found":         "Article introuvable dans le catalogue",
            "conversion_not_configured": "Conversion non configurée (TU/SU/CU manquants)",
        }
        for u in unresolved:
            ws_ur.append([
                u["product_code"],
                u["customer_code"],
                u["customer_name"],
                u["transaction_no"],
                u["qty"],
                u["uom"],
                reason_labels.get(u["reason"], u["reason"]),
            ])

        for col in ws_ur.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws_ur.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    wb_out.save(output_path)

    return {
        "total":      len(raw_rows),
        "resolved":   len(gcom_rows),
        "unresolved": len(unresolved),
    }


# ==========================================
# STREAMLIT UI DASHBOARD
# ==========================================
st.set_page_config(page_title="ZiarStock Data Cleaner", layout="centered", page_icon="🧹")

st.title("🧹 Centre de Nettoyage Excel")
st.write("Choisissez le type de nettoyage, uploadez votre fichier, et téléchargez le résultat.")

st.sidebar.header("Outils de Nettoyage")
tool = st.sidebar.radio("Sélectionnez l'outil :", [
    "1. Nettoyage G-COM",
    "2. Nettoyage Accenture",
    "3. Nettoyage Encour-RD",
    "4. Nettoyage Assabil",
    "5. Ferrero — Accenture CR19 → GCOM",
])

# ── Tool 5 has a different layout (two uploads) ──────────────────────────────
if "Ferrero" in tool:
    st.info(
        "**Comment ça marche :** Exportez le rapport **CR19 Document Listing** "
        "depuis Accenture, puis uploadez-le avec le fichier de correspondance "
        "**Ferrero_GCOM_ACCENTURE.xlsx**. L'outil convertit automatiquement "
        "les codes articles et les quantités (TU/SU/CU → unités GCOM)."
    )

    col1, col2 = st.columns(2)
    with col1:
        cr19_file = st.file_uploader("📄 Fichier CR19 Accenture (.xlsx)", type=["xlsx", "xls"])
    with col2:
        mapping_file = st.file_uploader("📋 Table de correspondance (.xlsx)", type=["xlsx"])

    if cr19_file and mapping_file:
        st.info("Fichiers chargés. Cliquez sur le bouton pour lancer la transformation.")

        if st.button("🚀 Lancer la Transformation", type="primary"):
            with st.spinner("Transformation en cours..."):
                cr19_ext = os.path.splitext(cr19_file.name)[1].lower()

                with tempfile.NamedTemporaryFile(delete=False, suffix=cr19_ext) as f_cr19:
                    f_cr19.write(cr19_file.getvalue())
                    cr19_path = f_cr19.name

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f_map:
                    f_map.write(mapping_file.getvalue())
                    mapping_path = f_map.name

                output_path = cr19_path.replace(cr19_ext, "_GCOM.xlsx")

                try:
                    summary = process_ferrero_accenture(cr19_path, output_path, mapping_path)

                    with open(output_path, "rb") as f:
                        processed_data = f.read()

                    st.success("✅ Transformation terminée avec succès !")

                    # Summary metrics
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total lignes", summary["total"])
                    m2.metric("✅ Converties", summary["resolved"])
                    m3.metric("⚠️ Non résolues", summary["unresolved"])

                    if summary["unresolved"] > 0:
                        st.warning(
                            f"**{summary['unresolved']} lignes non résolues** ont été placées "
                            "dans l'onglet **\"Non Résolus\"** du fichier téléchargé. "
                            "Vérifiez la table de correspondance pour ces articles."
                        )

                    output_filename = cr19_file.name.replace(cr19_ext, "_GCOM.xlsx")
                    st.download_button(
                        label="📥 Télécharger le Fichier GCOM",
                        data=processed_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"❌ Erreur lors de la transformation : {e}")

                finally:
                    for p in [cr19_path, mapping_path, output_path]:
                        if os.path.exists(p):
                            os.remove(p)
    else:
        st.warning("⬆️ Uploadez les deux fichiers pour continuer.")

# ── Tools 1-4: original single-file layout ───────────────────────────────────
else:
    if "Accenture" in tool or "Assabil" in tool:
        allowed_types = ["xlsx"]
    else:
        allowed_types = ["xls", "xlsx"]

    uploaded_file = st.file_uploader(f"Uploadez le fichier pour : {tool}", type=allowed_types)

    if uploaded_file is not None:
        st.info("Fichier chargé. Cliquez sur le bouton pour lancer le traitement.")

        if st.button("Lancer le Traitement", type="primary"):
            with st.spinner("Traitement en cours... Merci de patienter."):

                file_ext = os.path.splitext(uploaded_file.name)[1].lower()

                with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    input_path = tmp_in.name

                output_path = input_path.replace(file_ext, "_cleaned.xlsx")

                try:
                    if "G-COM" in tool:
                        process_gcom(input_path, output_path, file_ext)
                    elif "Accenture" in tool:
                        process_accenture(input_path, output_path)
                    elif "Encour-RD" in tool:
                        process_encour(input_path, output_path, file_ext)
                    elif "Assabil" in tool:
                        process_assabil(input_path, output_path)

                    with open(output_path, "rb") as f:
                        processed_data = f.read()

                    st.success("✅ Nettoyage terminé avec succès !")

                    st.download_button(
                        label="📥 Télécharger le Fichier Nettoyé",
                        data=processed_data,
                        file_name=uploaded_file.name.replace(file_ext, "_cleaned.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"❌ Une erreur s'est produite lors du traitement : {e}")

                finally:
                    if os.path.exists(input_path): os.remove(input_path)
                    if os.path.exists(output_path): os.remove(output_path)
