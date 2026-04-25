import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
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

    # 1) Delete rows where Article contains "Remise"
    for row in range(sheet.max_row, 1, -1):
        article = sheet.cell(row=row, column=COL_ARTICLE).value
        article_str = str(article).strip().lower() if article else ""
        if "remis" in article_str:
            sheet.delete_rows(row)

    # 2) Update Type client based on Agence and Vendeur
    for row in range(2, sheet.max_row + 1):
        agence = sheet.cell(row=row, column=COL_AGENCE).value
        vendeur = sheet.cell(row=row, column=COL_VENDEUR).value
        agence_str = str(agence).strip().lower() if agence else ""
        vendeur_str = str(vendeur).strip().lower() if vendeur else ""

        if agence_str == "agence el jadida":
            sheet.cell(row=row, column=COL_TYPE).value = "Gros"
        if vendeur_str == "bichlifin jamal":
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

    sheet._images = [] # Delete images
    for merged_range in list(sheet.merged_cells.ranges): # Unmerge
        sheet.unmerge_cells(str(merged_range))

    # Delete blank columns
    max_col, max_row = sheet.max_column, sheet.max_row
    for col in range(max_col, 0, -1):
        is_blank = True
        for row in range(1, max_row + 1):
            if sheet.cell(row=row, column=col).value not in (None, ""):
                is_blank = False
                break
        if is_blank: sheet.delete_cols(col)

    # Delete unwanted rows/cols
    sheet.delete_rows(1, 11)
    sheet.delete_cols(4)
    sheet.delete_cols(6)
    sheet.delete_cols(10)

    # Convert numeric
    for col in [15, 16, 17]:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value not in (None, ""):
                try: cell.value = convert_to_number(cell.value)
                except: pass

    # Date conversion
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

    # Pandas Summary
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
# STREAMLIT UI DASHBOARD
# ==========================================
st.set_page_config(page_title="ZiarStock Data Cleaner", layout="centered", page_icon="🧹")

st.title("🧹 Centre de Nettoyage Excel")
st.write("Choisissez le type de nettoyage, uploadez votre fichier, et téléchargez le résultat.")

# Option Selector
st.sidebar.header("Outils de Nettoyage")
tool = st.sidebar.radio("Sélectionnez l'outil :", 
    ["1. Nettoyage G-COM", "2. Nettoyage Accenture", "3. Nettoyage Encour-RD"]
)

# Set accepted formats based on tool
allowed_types = ["xlsx"] if "Accenture" in tool else ["xls", "xlsx"]

uploaded_file = st.file_uploader(f"Uploadez le fichier pour : {tool}", type=allowed_types)

if uploaded_file is not None:
    st.info("Fichier chargé. Cliquez sur le bouton pour lancer le traitement.")
    
    if st.button("Lancer le Traitement", type="primary"):
        with st.spinner("Traitement en cours... Merci de patienter."):
            
            # Create temporary files on the server to safely process data
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getvalue())
                input_path = tmp_in.name
                
            output_path = input_path.replace(file_ext, "_cleaned.xlsx")

            try:
                # Route to the correct Python logic
                if "G-COM" in tool:
                    process_gcom(input_path, output_path, file_ext)
                elif "Accenture" in tool:
                    process_accenture(input_path, output_path)
                elif "Encour-RD" in tool:
                    process_encour(input_path, output_path, file_ext)
                
                # Read the processed file to serve as download
                with open(output_path, "rb") as f:
                    processed_data = f.read()
                
                st.success("✅ Nettoyage terminé avec succès !")
                
                # Provide the Download Button
                st.download_button(
                    label="📥 Télécharger le Fichier Nettoyé",
                    data=processed_data,
                    file_name=uploaded_file.name.replace(file_ext, "_cleaned.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Une erreur s'est produite lors du traitement : {e}")
                
            finally:
                # Cleanup: Delete the temp files from the server so it doesn't run out of memory
                if os.path.exists(input_path): os.remove(input_path)
                if os.path.exists(output_path): os.remove(output_path)
