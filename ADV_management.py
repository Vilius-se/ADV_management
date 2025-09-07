# app.py
import streamlit as st
import pandas as pd
import io, re
from collections import Counter
import os

# =====================
# Pagalbinės funkcijos
# =====================

def norm_name(x):
    """Normalizuoja tekstą: pašalina tarpus, paverčia į didžiąsias raides."""
    return ''.join(str(x).upper().split())

def parse_qty(x):
    """Konvertuoja kiekio reikšmes (pvz. 1.000,52 / 1,000.52) į float."""
    if pd.isna(x): 
        return 0.0
    if isinstance(x,(int,float)): 
        return float(x)
    s=str(x).strip().replace('\xa0','').replace(' ','')
    if ',' in s and '.' in s: 
        s=s.replace(',','')   # 1,234.56 -> 1234.56
    else: 
        s=s.replace('.','').replace(',','.')
    try: 
        return float(s)
    except: 
        return 0.0

def safe_filename(s):
    """Sukuria saugų failo pavadinimą (pašalina draudžiamus simbolius)."""
    s='' if s is None else str(s).strip()
    s=re.sub(r'[\\/:*?"<>|]+','',s)
    return s.replace(' ','_')

def build_kaunas_index(df_kaunas):
    """Sukuria indeksą pagal Kaunas sandėlio komponentus ir jų likučius."""
    idx={}
    for _, r in df_kaunas.sort_values(['Component','Bin Code']).iterrows():
        key=r['Norm']
        q=float(r['Quantity'])
        if q<=0: 
            continue
        idx.setdefault(key,[]).append([r['Bin Code'], q])
    return idx

def allocate_bins_for_table(df_orders, bins_index):
    """Paskirsto komponentų kiekius į sandėlio bin'us."""
    if df_orders is None or df_orders.empty:
        return pd.DataFrame(columns=list(df_orders.columns)+['Bin Code'])
    rows=[]
    for _, r in df_orders.iterrows():
        item=r['Item No.']
        key=norm_name(item)
        remaining=float(r['Quantity'])
        if remaining<=0: 
            continue
        if key not in bins_index:
            row=r.copy()
            row['Document No.']=str(row.get('Document No.',''))+'/NERA'
            row['Bin Code']=''
            rows.append(row)
            continue
        bins=bins_index[key]; i=0
        while remaining>0 and i<len(bins):
            bin_code, avail=bins[i]
            if avail<=0: 
                i+=1; continue
            take=min(remaining, avail)
            row=r.copy()
            row['Job Task No.']=row.get('Job Task No.',1144)
            row['Quantity']=round(take,2)
            row['Bin Code']=bin_code
            rows.append(row)
            remaining-=take
            bins[i][1]=round(avail-take,6)
            if bins[i][1]<=1e-9: 
                i+=1
        if remaining>1e-9:
            row=r.copy()
            row['Quantity']=round(remaining,2)
            row['Document No.']=str(row.get('Document No.',''))+'/NERA'
            row['Bin Code']=''
            rows.append(row)
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=list(df_orders.columns)+['Bin Code'])

def process_stock_usage_keep_zeros(df_target, stock_df):
    """Naudoja turimą stock'ą, bet palieka įrašus net jei 0."""
    used_rows, adjusted_rows = [], []
    stock_tmp=stock_df.copy()
    stock_tmp['Norm']=stock_tmp['Component'].astype(str).map(norm_name)
    stock_norm_qty=stock_tmp.groupby('Norm',as_index=True)['Quantity'].sum().to_dict()
    for _, row in df_target.iterrows():
        item=row['Item No.']
        key=norm_name(item)
        qty_needed=float(row['Quantity'])
        qty_used=0.0
        qty_in_stock=float(stock_norm_qty.get(key,0))
        if qty_in_stock>0:
            qty_used=min(qty_needed, qty_in_stock)
            stock_norm_qty[key]=qty_in_stock-qty_used
        if qty_used>0:
            used_rows.append({'Item No.': item, 'Used from stock': qty_used})
        qty_remaining=qty_needed-qty_used
        adjusted=row.copy()
        adjusted['Job Task No.']=adjusted.get('Job Task No.',1144)
        adjusted['Quantity']=round(qty_remaining,2)
        adjusted_rows.append(adjusted)
    return pd.DataFrame(adjusted_rows), pd.DataFrame(used_rows)

def move_no_second_item_last(df):
    """Pakeičia stulpelių tvarką: 'No.' po pirmo, 'Item No.' į galą."""
    if df is None or df.empty: 
        return df
    cols=list(df.columns)
    if 'No.' in cols: 
        cols.insert(1, cols.pop(cols.index('No.')))
    if 'Item No.' in cols: 
        cols.append(cols.pop(cols.index('Item No.')))
    return df[cols]

def finalize_bom_alloc_table(df, vendor_to_no=None):
    """Paruošia galutinę BOM lentelę su stulpelių tvarka ir 'Vendor Item Number'."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if 'Used from stock' in out.columns:
        out = out.drop(columns=['Used from stock'])
    if 'Item No.' in out.columns:
        out = out.rename(columns={'Item No.': 'Vendor Item Number'})
    if 'Cross-Reference No.' in out.columns:
        out = out.rename(columns={'Cross-Reference No.': 'Vendor Item Number'})

    if 'No.' not in out.columns and 'Vendor Item Number' in out.columns:
        if vendor_to_no is not None:
            out['No.'] = out['Vendor Item Number'].map(lambda v: vendor_to_no.get(norm_name(v)))
    order = ['Type','No.','Document No.','Job No.','Job Task No.','Quantity','Location Code','Bin Code','Vendor Item Number']
    out = out[[c for c in order if c in out.columns] + [c for c in out.columns if c not in order]]
    return move_no_second_item_last(out)

def reorder_supplier_after_discount(df):
    """Pakeičia Supplier/Supplier No. stulpelių poziciją po Discount."""
    if df is None or df.empty: 
        return df
    cols=list(df.columns)
    changed=False
    if 'Supplier' in cols and 'Discount' in cols:
        cols.remove('Supplier'); cols.insert(cols.index('Discount')+1,'Supplier'); changed=True
    if 'Supplier No.' in cols and 'Discount' in cols:
        cols.remove('Supplier No.'); cols.insert(cols.index('Discount')+1,'Supplier No.'); changed=True
    return df[cols] if changed else df

def get_main_switch_and_accessories(ms_df, selected):
    """Grąžina main switch komponentą ir jo aksesuarus."""
    if ms_df is None or ms_df.empty or not selected: 
        return []
    row=ms_df[ms_df.iloc[:,0].astype(str).str.strip().str.upper()==str(selected).strip().upper()]
    if row.empty: 
        return []
    vals=[str(selected).strip()] + [str(v).strip() for v in row.iloc[0,1:].tolist() if pd.notna(v) and str(v).strip()]
    seen=set(); out=[]
    for v in vals:
        key=norm_name(v)
        if key and key not in seen:
            seen.add(key); out.append(v)
    return out

def read_excel_any(file, sheet_name=0):
    name = getattr(file, "name", "")
    ext = os.path.splitext(name)[1].lower()
    if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        return pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")
    elif ext == ".xls":
        # Reikia xlrd==1.2.0
        return pd.read_excel(file, sheet_name=sheet_name, engine="xlrd")
    elif ext == ".csv":
        return pd.read_csv(file)
    else:
        # bandymas su default
        return pd.read_excel(file, sheet_name=sheet_name)


# =====================
# Streamlit UI
# =====================

st.set_page_config(page_title="Advansor Tool", layout="wide")
st.title("⚙️ Advansor Component Tool")

# Failų įkėlimas (drag & drop)
st.header("📂 Įkelk failus")
cubic_file = st.file_uploader("Įkelk CUBIC", type=["xls", "xlsx", "xlsm", "csv"])
bom_file   = st.file_uploader("Įkelk BOM",   type=["xls", "xlsx", "xlsm", "csv"])
data_file  = st.file_uploader("Įkelk Data",  type=["xls", "xlsx", "xlsm"])
ks_file    = st.file_uploader("Įkelk Kaunas Stock", type=["xls", "xlsx", "xlsm", "csv"])

# Projekto parametrai
st.header("📋 Projekto parametrai")
doc_number = st.text_input("📄 Document No.")
type_choice = st.selectbox(
    "📌 Tipas:", 
    ['A','B','B1','B2','C','C1','C2','C3','C4','C4.1','C5','C6','C7','C8',
     'F','F1','F2','F3','F4','F4.1','F5','F6','F7','G','G1','G2','G3','G4','G5','G6','G7','Custom']
)
main_switch = st.selectbox(
    "⚡ Kirtiklis:", 
    ["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111","31109","31107","C404400S","C634630S"]
)
earthing    = st.selectbox("🌍 Įžeminimas:", ['TT','TN-S','TN-C-S'])
swing       = st.checkbox("🔄 Ar bus Swing Frame?")
ups         = st.checkbox("🔋 Ar bus UPS?")

# Paleidimo mygtukas
generate = st.button("✅ Generuoti")
if generate:
    if not all([cubic_file, bom_file, data_file, ks_file, doc_number]):
        st.warning("⚠️ Įkelk visus failus ir įvesk Document No.")
    else:
        try:
            # =====================
            # Failų nuskaitymas
            # =====================
            df_cubic = read_excel_any(cubic_file)
            df_bom_raw = read_excel_any(cubic_file)
            df_data = pd.ExcelFile(data_file)
            df_ks      = read_excel_any(ks_file, sheet_name=0)
            df_kaunas = pd.read_excel(ks_file, sheet_name=0, header=None, skiprows=3, usecols=[1,3,10])
            df_kaunas.columns=['Bin Code','Quantity','Component']
            df_kaunas['Quantity']=df_kaunas['Quantity'].apply(parse_qty).fillna(0)
            df_kaunas['Norm']=df_kaunas['Component'].apply(norm_name)
            df_kaunas=df_kaunas[df_kaunas['Quantity']>0].reset_index(drop=True)
            kaunas_bins_index=build_kaunas_index(df_kaunas)

            # =====================
            # BOM apdorojimas (trumpinta versija, integruojama tavo logika)
            # =====================
            df_bom = df_bom_raw.copy()
            df_bom.columns = ['Original_Item_Ref','Type','Quantity','Manufacturer','Description']
            df_bom['Quantity'] = pd.to_numeric(df_bom['Quantity'], errors='coerce').fillna(0)

            # Pvz.: pridėti Document No., Job No., Task No.
            df_bom_fmt = pd.DataFrame({
                'Type': 'Item',
                'Document No.': doc_number,
                'Job No.': doc_number,
                'Job Task No.': 1144,
                'Cross-Reference No.': df_bom['Type'],
                'Quantity': df_bom['Quantity'],
                'Location Code': 'KAUNAS',
                'Manufacturer': df_bom['Manufacturer'],
                'Description': df_bom['Description']
            })

            # =====================
            # Specialūs pasirinkimai
            # =====================
            if ups:
                df_bom_fmt = pd.concat([df_bom_fmt, pd.DataFrame([{
                    'Type':'Item','Document No.':doc_number,'Job No.':doc_number,
                    'Job Task No.':1144,'Cross-Reference No.':'ADV UPS HOLDER V3',
                    'Quantity':1,'Location Code':'KAUNAS','Manufacturer':'','Description':'UPS Holder'
                }])], ignore_index=True)

            if swing:
                df_bom_fmt = pd.concat([df_bom_fmt, pd.DataFrame([
                    {'Type':'Item','Document No.':doc_number,'Job No.':doc_number,'Job Task No.':1144,
                     'Cross-Reference No.':'1055-1000','Quantity':2,'Location Code':'KAUNAS','Manufacturer':'','Description':'Swing accessory 1'},
                    {'Type':'Item','Document No.':doc_number,'Job No.':doc_number,'Job Task No.':1144,
                     'Cross-Reference No.':'1055-1001','Quantity':2,'Location Code':'KAUNAS','Manufacturer':'','Description':'Swing accessory 2'}
                ])], ignore_index=True)

            # =====================
            # Atvaizdavimas
            # =====================
            st.subheader("📊 Sugeneruotas BOM")
            st.dataframe(df_bom_fmt)

            # Čia galima pridėti daugiau tavo logikos (Main Switch accessories, stock usage ir t.t.)
            # Tam integruosim pilną tavo Colab pipeline kitame žingsnyje.

            st.success("✅ BOM sėkmingai sugeneruotas")

            # Rezultato eksportas į Excel atmintyje
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_bom_fmt.to_excel(writer, index=False, sheet_name="BOM")

            st.download_button(
                label="⬇️ Atsisiųsti rezultatą",
                data=output.getvalue(),
                file_name=f"{safe_filename(doc_number)}_{safe_filename(type_choice)}_{safe_filename(earthing)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Klaida apdorojant failus: {e}")
                        # =====================
            # Calculation sheet
            # =====================
            # Bandysim paimti kainas iš Data.xlsx (lapas Hours)
            try:
                df_hours = pd.read_excel(data_file, sheet_name="Hours", header=None)
                hourly_rate = float(df_hours.iloc[1, 4])  # valandos kaina (E2 langelis)
                proj_type = str(type_choice).strip().upper()
                earthing_type = str(earthing).strip().upper()
                row_match = df_hours[df_hours.iloc[:,0].astype(str).str.upper() == proj_type]
                hours_value = 0
                if not row_match.empty:
                    if earthing_type == "TT":
                        hours_value = float(row_match.iloc[0,1])
                    elif earthing_type == "TN-S":
                        hours_value = float(row_match.iloc[0,2])
                    elif earthing_type == "TN-C-S":
                        hours_value = float(row_match.iloc[0,3])
                hours_cost = hours_value * hourly_rate
            except Exception:
                hours_value = 0
                hours_cost = 0
                hourly_rate = 0

            # Pavyzdinės kainos (čia reikėtų imti iš BOM po konvertacijos, kaip Colab’e)
            parts_cost = (df_bom_fmt['Quantity'].sum()) * 10   # čia demo, realiai imama iš Unit Cost
            cubic_cost = 5000.0
            smart_supply_cost = 9750.0
            wire_set_cost = 2500.0

            # Calculation į Excel
            from openpyxl import load_workbook
            from openpyxl.styles import Font, Border, Side, Alignment

            output_calc = io.BytesIO()
            with pd.ExcelWriter(output_calc, engine="openpyxl") as writer:
                df_bom_fmt.to_excel(writer, index=False, sheet_name="BOM")
                writer.book.create_sheet("Calculation")
                ws = writer.book["Calculation"]

                # Stulpelių plotis
                ws.column_dimensions['A'].width = 20
                ws.column_dimensions['B'].width = 20
                ws.column_dimensions['C'].width = 20

                thin = Side(style='thin')
                border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
                bold = Font(bold=True)

                # Projekto info
                ws['A2'] = f"{doc_number} | {proj_type} | {earthing_type}"
                ws['A2'].font = bold

                rows = [
                    ("Parts:",         parts_cost, ""),
                    ("Cubic:",         cubic_cost, ""),
                    ("Hours cost:",    hours_cost, f"{int(hours_value)} Hours"),
                    ("Smart supply:",  smart_supply_cost, ""),
                    ("Wire set:",      wire_set_cost, ""),
                    ("Extra:",         None, ""),   # ranka įvedama
                    ("Total:",         None, ""),
                    ("Total+5%:",      None, ""),
                    ("Total+35%:",     None, ""),
                ]

                start_row = 4
                for r, (label, value, info) in enumerate(rows, start=start_row):
                    ws.cell(row=r, column=1, value=label)
                    if label.startswith("Total"):
                        ws.cell(row=r, column=1).font = bold
                    if label == "Extra:":
                        ws.cell(row=r, column=2, value=None)
                    elif value is not None:
                        ws.cell(row=r, column=2, value=float(value))
                    ws.cell(row=r, column=3, value=info if info else None)

                # Formulės
                B = lambda row: f"B{row}"
                total_row = start_row + 6
                ws[B(total_row)]     = f"=SUM(B{start_row}:B{start_row+5})"
                ws[B(total_row+1)]   = f"={B(total_row)}*1.05"
                ws[B(total_row+2)]   = f"={B(total_row)}*1.35"

                # Formatavimas DKK
                for r in range(start_row, start_row+9):
                    ws[f"B{r}"].number_format = '#,##0.00 "DKK"'
                    ws[f"A{r}"].border = border_all
                    ws[f"B{r}"].border = border_all
                    ws[f"C{r}"].border = border_all
                    ws[f"A{r}"].alignment = Alignment(vertical='center')
                    ws[f"B{r}"].alignment = Alignment(vertical='center')

            st.download_button(
                label="⬇️ Atsisiųsti su Calculation",
                data=output_calc.getvalue(),
                file_name=f"{safe_filename(doc_number)}_{safe_filename(type_choice)}_{safe_filename(earthing)}_calc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
