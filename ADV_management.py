import streamlit as st
import pandas as pd
import io, re, os
from collections import Counter
from openpyxl.styles import Font, Border, Side, Alignment

def norm_name(x):
    return ''.join(str(x).upper().split())

def parse_qty(x):
    if pd.isna(x): return 0.0
    if isinstance(x,(int,float)): return float(x)
    s=str(x).strip().replace('\xa0','').replace(' ','')
    if ',' in s and '.' in s: s=s.replace(',','')
    else: s=s.replace('.','').replace(',','.')
    try: return float(s)
    except: return 0.0

def safe_filename(s):
    s='' if s is None else str(s).strip()
    s=re.sub(r'[\\/:*?"<>|]+','',s)
    return s.replace(' ','_')

def build_kaunas_index(df_kaunas):
    idx={}
    for _, r in df_kaunas.sort_values(['Component','Bin Code']).iterrows():
        key=r['Norm']; q=float(r['Quantity'])
        if q<=0: continue
        idx.setdefault(key,[]).append([r['Bin Code'], q])
    return idx

def allocate_bins_for_table(df_orders, bins_index):
    if df_orders is None or df_orders.empty:
        return pd.DataFrame(columns=list(df_orders.columns)+['Bin Code'])
    rows=[]
    for _, r in df_orders.iterrows():
        item=r['Item No.']; key=norm_name(item); remaining=float(r['Quantity'])
        if remaining<=0: continue
        if key not in bins_index:
            row=r.copy(); row['Document No.']=str(row.get('Document No.',''))+'/NERA'; row['Bin Code']=''; rows.append(row); continue
        bins=bins_index[key]; i=0
        while remaining>0 and i<len(bins):
            bin_code, avail=bins[i]
            if avail<=0: i+=1; continue
            take=min(remaining, avail)
            row=r.copy(); row['Job Task No.']=row.get('Job Task No.',1144); row['Quantity']=round(take,2); row['Bin Code']=bin_code
            rows.append(row); remaining-=take; bins[i][1]=round(avail-take,6)
            if bins[i][1]<=1e-9: i+=1
        if remaining>1e-9:
            row=r.copy(); row['Quantity']=round(remaining,2); row['Document No.']=str(row.get('Document No.',''))+'/NERA'; row['Bin Code']=''; rows.append(row)
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=list(df_orders.columns)+['Bin Code'])

def process_stock_usage_keep_zeros(df_target, stock_df):
    used_rows, adjusted_rows = [], []
    stock_tmp=stock_df.copy(); stock_tmp['Norm']=stock_tmp['Component'].astype(str).map(norm_name)
    stock_norm_qty=stock_tmp.groupby('Norm',as_index=True)['Quantity'].sum().to_dict()
    for _, row in df_target.iterrows():
        item=row['Item No.']; key=norm_name(item); qty_needed=float(row['Quantity']); qty_used=0.0; qty_in_stock=float(stock_norm_qty.get(key,0))
        if qty_in_stock>0:
            qty_used=min(qty_needed, qty_in_stock); stock_norm_qty[key]=qty_in_stock-qty_used
        if qty_used>0: used_rows.append({'Item No.': item, 'Used from stock': qty_used})
        qty_remaining=qty_needed-qty_used; adjusted=row.copy(); adjusted['Job Task No.']=adjusted.get('Job Task No.',1144); adjusted['Quantity']=round(qty_remaining,2); adjusted_rows.append(adjusted)
    return pd.DataFrame(adjusted_rows), pd.DataFrame(used_rows)

def move_no_second_item_last(df):
    if df is None or df.empty: return df
    cols=list(df.columns)
    if 'No.' in cols: cols.insert(1, cols.pop(cols.index('No.')))
    if 'Item No.' in cols: cols.append(cols.pop(cols.index('Item No.')))
    return df[cols]

def finalize_bom_alloc_table(df, vendor_to_no=None):
    if df is None or df.empty: return df
    out = df.copy()
    if 'Used from stock' in out.columns: out = out.drop(columns=['Used from stock'])
    if 'Item No.' in out.columns: out = out.rename(columns={'Item No.': 'Vendor Item Number'})
    if 'Cross-Reference No.' in out.columns: out = out.rename(columns={'Cross-Reference No.': 'Vendor Item Number'})
    if 'No.' not in out.columns and 'Vendor Item Number' in out.columns:
        if vendor_to_no is not None: out['No.'] = out['Vendor Item Number'].map(lambda v: vendor_to_no.get(norm_name(v)))
    order = ['Type','No.','Document No.','Job No.','Job Task No.','Quantity','Location Code','Bin Code','Vendor Item Number']
    out = out[[c for c in order if c in out.columns] + [c for c in out.columns if c not in order]]
    return move_no_second_item_last(out)

def reorder_supplier_after_discount(df):
    if df is None or df.empty: return df
    cols=list(df.columns); changed=False
    if 'Supplier' in cols and 'Discount' in cols: cols.remove('Supplier'); cols.insert(cols.index('Discount')+1,'Supplier'); changed=True
    if 'Supplier No.' in cols and 'Discount' in cols: cols.remove('Supplier No.'); cols.insert(cols.index('Discount')+1,'Supplier No.'); changed=True
    return df[cols] if changed else df

def get_main_switch_and_accessories(ms_df, selected):
    if ms_df is None or ms_df.empty or not selected: return []
    row=ms_df[ms_df.iloc[:,0].astype(str).str.strip().str.upper()==str(selected).strip().upper()]
    if row.empty: return []
    vals=[str(selected).strip()] + [str(v).strip() for v in row.iloc[0,1:].tolist() if pd.notna(v) and str(v).strip()]
    seen=set(); out=[]
    for v in vals:
        key=norm_name(v)
        if key and key not in seen: seen.add(key); out.append(v)
    return out

def excel_file_any(file):
    if file is None: raise ValueError("No file provided to excel_file_any")
    try: file.seek(0)
    except Exception: pass
    name = getattr(file, "name", ""); ext = os.path.splitext(name)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"): return pd.ExcelFile(file, engine="openpyxl")
    elif ext == ".xls": return pd.ExcelFile(file, engine="xlrd")
    else: return pd.ExcelFile(file)

def process_bom(df_bom_raw, xf_data, df_part_no):
    df_bom = df_bom_raw.copy()
    df_bom.columns = ['Original_Item_Ref','Type','Quantity','Manufacturer','Description']
    df_bom['Quantity'] = pd.to_numeric(df_bom['Quantity'], errors='coerce').fillna(0)
    df_partcode_full = pd.read_excel(xf_data, sheet_name='Part_code', usecols=[0,1,2], header=0)
    df_partcode_full.columns = ['From','To','PC_Manufacturer']
    name_map = {norm_name(a): str(b).strip() for a, b in df_partcode_full[['From','To']].dropna(subset=['From']).values if pd.notna(b) and str(b).strip()}
    manuf_map = {}
    for _, r in df_partcode_full.iterrows():
        a, b, m = r['From'], r['To'], r['PC_Manufacturer']
        if pd.isna(a): continue
        if pd.notna(m) and str(m).strip():
            if pd.notna(b) and str(b).strip(): manuf_map[norm_name(str(b))] = str(m).strip()
            manuf_map.setdefault(norm_name(str(a)), str(m).strip())
    df_bom['Type'] = df_bom['Type'].apply(lambda s: name_map.get(norm_name(s), s) if (pd.notna(s) and str(s).strip()) else s)
    mask_missing_manuf = df_bom['Manufacturer'].isna() | (df_bom['Manufacturer'].astype(str).str.strip() == '')
    df_bom.loc[mask_missing_manuf, 'Manufacturer'] = df_bom.loc[mask_missing_manuf, 'Type'].apply(lambda s: manuf_map.get(norm_name(s)) if (pd.notna(s) and str(s).strip()) else None)
    exclude_manufacturers = ["BITZER KÜHLMACHINENBAU GMBH","BELDEN","KABELTEC","HELUKABEL","LAPP","GENERAL CAVI","EMERSON CLIMATE TECHNOLOGIES","ELFAC A/S","DEKA CONTROLS GMBH","TYPE SPECIFIED IN BOM","PRYSMIAN","CO4","BELIMO","PEPPERL + FUCHS","WAGO"]
    exclude_types = ["47KOHM","134F7613","134H7160","CABLE JZ","AKSF","ROUNDPACKART","XALK178E","LMBWLB32-180S","ACH580-01-026A-4","ACH580-01-033A-4","ACH580-01-073A-4","PSP650MT3-230U"]
    comments_to_exclude = {"Q1","NO NEED"}
    df_stock_comments = pd.read_excel(xf_data, sheet_name="Stock", usecols=[0,1,2], names=["Component","Quantity","Comment"], header=None, skiprows=1)
    df_stock_comments["Component_norm"] = df_stock_comments["Component"].astype(str).str.strip().str.upper()
    df_stock_comments["Comment_norm"] = df_stock_comments["Comment"].astype(str).str.strip().str.upper()
    exclude_by_comment = set(df_stock_comments.loc[df_stock_comments["Comment_norm"].isin(comments_to_exclude),"Component_norm"].unique())
    _dfm = df_bom.copy(); _dfm["Type_norm"] = _dfm["Type"].astype(str).str.strip().str.upper(); _dfm["Manufacturer_norm"] = _dfm["Manufacturer"].astype(str).str.strip().str.upper()
    mask_types = ~_dfm["Type_norm"].isin(set(t.upper() for t in exclude_types))
    mask_comment = ~_dfm["Type_norm"].isin(exclude_by_comment)
    mask_manuf = ~_dfm["Manufacturer_norm"].isin(set(m.upper() for m in exclude_manufacturers))
    df_bom_filtered = _dfm[mask_types & mask_comment & mask_manuf].reset_index(drop=True)
    merged_conv = df_bom_filtered.copy(); merged_conv['Norm_Ref'] = merged_conv['Type'].astype(str).map(norm_name)
    merged_conv = merged_conv.merge(df_part_no[['PartNo_A','PartName_B','Desc_C','Manufacturer_D','SupplierNo_E','UnitPrice_F','Norm_B']], left_on='Norm_Ref', right_on='Norm_B', how='left')
    def profit_by_manuf(x: str) -> int: s = ('' if pd.isna(x) else str(x)).upper(); return 10 if ('DANFOSS' in s or 'CAREL' in s) else 17
    merged_conv['ProfitVal'] = merged_conv['Manufacturer_D'].apply(profit_by_manuf)
    df_bom_po_konvertacijos = pd.DataFrame({'Type': 'item','No.': merged_conv['PartNo_A'],'Quantity': pd.to_numeric(merged_conv['Quantity'], errors='coerce'),'Supplier No.': merged_conv['SupplierNo_E'],'Supplier': merged_conv['Manufacturer_D'],'Profit': merged_conv['ProfitVal'],'Discount': 0,'Vendor Item Number': merged_conv['PartName_B'],'Description': merged_conv['Desc_C'],'Unit Cost': pd.to_numeric(merged_conv['UnitPrice_F'], errors='coerce')})
    df_bom_po_konvertacijos = move_no_second_item_last(df_bom_po_konvertacijos)
    df_bom_po_konvertacijos = reorder_supplier_after_discount(df_bom_po_konvertacijos)
    return df_bom_po_konvertacijos

st.set_page_config(page_title="Advansor Tool", layout="wide")
st.title("⚙️ Advansor Component Tool")

st.header("📂 Įkelk failus")
cubic_file = st.file_uploader("Įkelk CUBIC", type=["xls","xlsx","xlsm","csv"])
bom_file   = st.file_uploader("Įkelk BOM",   type=["xls","xlsx","xlsm","csv"])
data_file  = st.file_uploader("Įkelk Data",  type=["xls","xlsx","xlsm"])
ks_file    = st.file_uploader("Įkelk Kaunas Stock", type=["xls","xlsx","xlsm","csv"])

st.header("📋 Projekto parametrai")
doc_number = st.text_input("📄 Document No.")
type_choice = st.selectbox("📌 Tipas:", ['A','B','B1','B2','C','C1','C2','C3','C4','C4.1','C5','C6','C7','C8','F','F1','F2','F3','F4','F4.1','F5','F6','F7','G','G1','G2','G3','G4','G5','G6','G7','Custom'])
main_switch = st.selectbox("⚡ Kirtiklis:", ["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111","31109","31107","C404400S","C634630S"])
earthing    = st.selectbox("🌍 Įžeminimas:", ['TT','TN-S','TN-C-S'])
swing       = st.checkbox("🔄 Ar bus Swing Frame?")
ups         = st.checkbox("🔋 Ar bus UPS?")
generate = st.button("✅ Generuoti")

if generate:
    if not all([cubic_file, bom_file, data_file, ks_file, doc_number]):
        st.warning("⚠️ Įkelk visus failus ir įvesk Document No.")
    else:
        try:
            xf_data = excel_file_any(data_file)
            df_bom_raw = pd.read_excel(bom_file)
            df_part_no = pd.read_excel(xf_data, sheet_name='Part_no', usecols="A:F", header=0)
            df_part_no.columns = ['PartNo_A','PartName_B','Desc_C','Manufacturer_D','SupplierNo_E','UnitPrice_F']
            df_part_no['Norm_B'] = df_part_no['PartName_B'].astype(str).map(norm_name)
            df_bom_po_konvertacijos = process_bom(df_bom_raw, xf_data, df_part_no)
            st.subheader("📊 BOM po konvertacijos")
            st.dataframe(df_bom_po_konvertacijos.head())
            def safe_sum_cost(df):
                if df is None or len(df)==0: return 0.0
                q = pd.to_numeric(df.get('Quantity'), errors='coerce').fillna(0)
                c = pd.to_numeric(df.get('Unit Cost'), errors='coerce').fillna(0)
                return float((q * c).sum())
            try:
                df_hours = pd.read_excel(xf_data, sheet_name="Hours", header=None)
                hourly_rate = float(df_hours.iloc[1, 4])
                proj_type = str(type_choice).strip().upper(); earthing_type = str(earthing).strip().upper()
                row_match = df_hours[df_hours.iloc[:,0].astype(str).str.upper() == proj_type]
                hours_value = 0
                if not row_match.empty:
                    if earthing_type == "TT": hours_value = float(row_match.iloc[0,1])
                    elif earthing_type == "TN-S": hours_value = float(row_match.iloc[0,2])
                    elif earthing_type == "TN-C-S": hours_value = float(row_match.iloc[0,3])
                hours_cost = hours_value * hourly_rate
            except Exception:
                hours_value = 0; hours_cost = 0; hourly_rate = 0
            parts_cost = safe_sum_cost(df_bom_po_konvertacijos)
            cubic_cost = 0.0
            smart_supply_cost = 9750.0
            wire_set_cost = 2500.0
            output_calc = io.BytesIO()
            with pd.ExcelWriter(output_calc, engine="openpyxl") as writer:
                if df_bom_po_konvertacijos is not None and len(df_bom_po_konvertacijos):
                    df_bom_po_konvertacijos.to_excel(writer, index=False, sheet_name="BOM po konvertacijos")
                wb = writer.book; ws = wb.create_sheet("Calculation")
                ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 20; ws.column_dimensions['C'].width = 20
                thin = Side(style='thin'); border_all = Border(left=thin, right=thin, top=thin, bottom=thin); bold = Font(bold=True)
                ws['A2'] = f"{doc_number} | {str(type_choice).upper()} | {str(earthing).upper()}"; ws['A2'].font = bold
                rows = [("Parts:", parts_cost, ""),("Cubic:", cubic_cost, ""),("Hours cost:", hours_cost, f"{int(hours_value)} Hours"),("Smart supply:", smart_supply_cost, ""),("Wire set:", wire_set_cost, ""),("Extra:", None, ""),("Total:", None, ""),("Total+5%:", None, ""),("Total+35%:", None, "")]
                start_row = 4
                for r, (label, value, info) in enumerate(rows, start=start_row):
                    ws.cell(row=r, column=1, value=label)
                    if label.startswith("Total"): ws.cell(row=r, column=1).font = bold
                    if label == "Extra:": ws.cell(row=r, column=2, value=None)
                    elif value is not None: ws.cell(row=r, column=2, value=float(value))
                    ws.cell(row=r, column=3, value=info if info else None)
                B = lambda row: f"B{row}"; total_row = start_row + 6
                ws[B(total_row)]   = f"=SUM(B{start_row}:B{start_row+5})"
                ws[B(total_row+1)] = f"={B(total_row)}*1.05"; ws[B(total_row+2)] = f"={B(total_row)}*1.35"
                for r in range(start_row, start_row+9):
                    ws[f"B{r}"].number_format = '#,##0.00 "DKK"'
                    for c in ("A","B","C"): ws[f"{c}{r}"].border = border_all
            st.download_button(label="⬇️ Atsisiųsti su Calculation", data=output_calc.getvalue(), file_name=f"{safe_filename(doc_number)}_{safe_filename(type_choice)}_{safe_filename(earthing)}_calc.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Klaida apdorojant failus: {e}")
