import streamlit as st
import pandas as pd
import re
import io
import datetime
from openpyxl import Workbook

def add_extra_components(df, extras):
    if df is None: 
        df = pd.DataFrame()
    df_out = df.copy()
    for e in extras:
        extra_row = pd.DataFrame([{
            "Original Type": e["type"],
            "Type": e["type"],
            "Quantity": e.get("qty", 1),
            "Source": "Extra",
            "No.": e.get("force_no", e["type"])  # <- čia kritinis
        }])
        df_out = pd.concat([df_out, extra_row], ignore_index=True)
    return df_out


def build_nav_table_from_bom(df_bom: pd.DataFrame, df_part_no: pd.DataFrame, label: str = "Project BOM") -> pd.DataFrame:
    req = ["PartNo_A", "SupplierNo_E", "Manufacturer_D"]
    if df_part_no is None or df_part_no.empty or any(c not in df_part_no.columns for c in req):
        return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])
    supplier_map = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["Manufacturer_D"].astype(str)))
    tmp = df_bom.copy()
    if "Quantity" not in tmp.columns: tmp["Quantity"] = 0
    if "Description" not in tmp.columns: tmp["Description"] = ""
    if "No." not in tmp.columns: tmp["No."] = ""
    tmp["No."] = tmp["No."].astype(str)
    tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0)
    nav_rows = []
    for _, r in tmp.iterrows():
        part_no = str(r["No."]).strip()
        qty = float(r.get("Quantity", 0) or 0)
        manuf = manuf_map.get(part_no, "")
        profit = 10 if "DANFOSS" in str(manuf).upper() else 17
        supplier = supplier_map.get(part_no, 30093)
        nav_rows.append({
            "Type": "Item",
            "No.": part_no,
            "Quantity": qty,
            "Supplier": supplier,
            "Profit": profit,
            "Discount": 0,
            "Description": r.get("Description", "")
        })
    nav_table = pd.DataFrame(nav_rows, columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])
    return nav_table

def pipeline_1_1_norm_name(x): return ''.join(str(x).upper().split())
def pipeline_1_2_parse_qty(x):
    if pd.isna(x): return 0.0
    if isinstance(x,(int,float)): return float(x)
    s = str(x).strip().replace('\xa0','').replace(' ','')
    if ',' in s and '.' in s: s = s.replace(',','')
    else: s = s.replace('.','').replace(',','.')
    try: return float(s)
    except: return 0.0
def pipeline_1_3_safe_filename(s):
    s = '' if s is None else str(s).strip()
    s = re.sub(r'[\\/:*?"<>|]+','',s)
    return s.replace(' ','_')
def pipeline_1_4_normalize_no(x):
    try: return str(int(float(str(x).replace(",","." ).strip())))
    except: return str(x).strip()
def read_excel_any(file,**kwargs):
    try: return pd.read_excel(file,engine="openpyxl",**kwargs)
    except: return pd.read_excel(file,engine="xlrd",**kwargs)
def allocate_from_stock(no,qty_needed,stock_rows):
    allocations=[]
    qty_needed=int(round(pd.to_numeric(pd.Series([qty_needed]),errors="coerce").fillna(0).iloc[0]))
    remaining=qty_needed
    if stock_rows is not None and not stock_rows.empty:
        for _,srow in stock_rows.iterrows():
            if remaining<=0: break
            bin_code=str(srow.get("Bin Code","")).strip()
            stock_qty=pd.to_numeric(pd.Series([srow.get("Quantity",0)]),errors="coerce").fillna(0).iloc[0]
            if stock_qty<=0: continue
            if bin_code=="67-01-01-01": continue
            take=min(int(round(stock_qty)),remaining)
            if take>0:
                allocations.append({"No.":no,"Bin Code":bin_code,"Allocated Qty":take})
                remaining-=take
    if remaining>0: allocations.append({"No.":no,"Bin Code":"","Allocated Qty":remaining})
    return allocations
normalize_no = pipeline_1_4_normalize_no

def pipeline_2_1_user_inputs():
    st.subheader("Project Information")
    project_number = st.text_input("Project number (1234-567)")
    if project_number and not re.match(r"^\d{4}-\d{3}$", project_number):
        st.error("Invalid format (must be 1234-567)")
        return None
    panel_type = st.selectbox("Panel type", ['A','B','B1','B2','C','C1','C2','C3','C4','C4.1','C5','C6','C7','C8','F','F1','F2','F3','F4','F4.1','F5','F6','F7','G','G1','G2','G3','G4','G5','G6','G7','Custom'])
    grounding = st.selectbox("Grounding type", ["TT","TN-S","TN-C-S"])
    main_switch = st.selectbox("Main switch", ["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111","31109","31107","C404400S","C634630S"])
    swing_frame = st.checkbox("Swing frame?")
    ups = st.checkbox("UPS?")
    rittal = st.checkbox("Rittal?")
    return {"project_number": project_number,"panel_type": panel_type,"grounding": grounding,"main_switch": main_switch,"swing_frame": swing_frame,"ups": ups,"rittal": rittal}

def pipeline_2_2_file_uploads(rittal=False):
    st.subheader("Upload Required Files")
    dfs = {}
    if not rittal:
        cubic_bom = st.file_uploader("Insert CUBIC BOM", type=["xls","xlsx","xlsm"], key="cubic_bom")
        if cubic_bom:
            df_cubic = read_excel_any(cubic_bom, skiprows=13, usecols="B,E:F,G")
            df_cubic = df_cubic.rename(columns=lambda c:str(c).strip())
            if {"E","F","G"}.issubset(df_cubic.columns): df_cubic["Quantity"] = df_cubic[["E","F","G"]].bfill(axis=1).iloc[:,0]
            elif "Quantity" not in df_cubic.columns: df_cubic["Quantity"] = 0
            df_cubic["Quantity"] = pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0)
            df_cubic = df_cubic.rename(columns={"Item Id":"Type"})
            df_cubic["Original Type"] = df_cubic["Type"]
            df_cubic["No."] = df_cubic["Type"]
            dfs["cubic_bom"] = df_cubic
    bom = st.file_uploader("Insert BOM", type=["xls","xlsx","xlsm"], key="bom")
    if bom:
        df_bom = read_excel_any(bom)
        if df_bom.shape[1] >= 2:
            colA = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
            colB = df_bom.iloc[:,1].fillna("").astype(str).str.strip()
            df_bom["Original Article"] = colA
            df_bom["Original Type"] = colB.where(colB!="",colA)
        else:
            df_bom["Original Article"] = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
            df_bom["Original Type"] = df_bom["Original Article"]
        dfs["bom"] = df_bom
    data_file = st.file_uploader("Insert DATA", type=["xls","xlsx","xlsm"], key="data")
    if data_file: dfs["data"] = pd.read_excel(data_file, sheet_name=None)
    ks_file = st.file_uploader("Insert Kaunas Stock", type=["xls","xlsx","xlsm"], key="ks")
    if ks_file: dfs["ks"] = read_excel_any(ks_file)
    return dfs

def pipeline_2_3_get_sheet_safe(data_dict, names):
    if not isinstance(data_dict,dict): return None
    for key in data_dict.keys():
        if str(key).strip().upper().replace(" ","_") in [n.upper().replace(" ","_") for n in names]:
            return data_dict[key]
    return None

def pipeline_2_4_normalize_part_no(df_raw):
    if df_raw is None or df_raw.empty: return pd.DataFrame()
    df = df_raw.copy().rename(columns=lambda c:str(c).strip())
    col_map = {}
    if df.shape[1] >= 1: col_map[df.columns[0]] = "PartNo_A"
    if df.shape[1] >= 2: col_map[df.columns[1]] = "PartName_B"
    if df.shape[1] >= 3: col_map[df.columns[2]] = "Desc_C"
    if df.shape[1] >= 4: col_map[df.columns[3]] = "Manufacturer_D"
    if df.shape[1] >= 5: col_map[df.columns[4]] = "SupplierNo_E"
    if df.shape[1] >= 6: col_map[df.columns[5]] = "UnitPrice_F"
    return df.rename(columns=col_map)

def pipeline_3A_0_rename(df_bom, df_part_code, extras=None):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    df = df_bom.copy()
    if df_part_code is not None and not df_part_code.empty:
        rename_map = dict(zip(df_part_code.iloc[:, 0].astype(str).str.strip(), df_part_code.iloc[:, 1].astype(str).str.strip()))
        for col in ["Type", "Original Type"]:
            if col in df.columns: df[col] = df[col].astype(str).str.strip().replace(rename_map)
    if "Type" not in df.columns: df["Type"] = df.iloc[:, 0].astype(str)
    if "Original Type" not in df.columns: df["Original Type"] = df["Type"]
    if "Original Article" not in df.columns: df["Original Article"] = df.iloc[:, 0].astype(str)
    if extras: df = add_extra_components(df, [e for e in extras if e.get("target") == "bom"])
    return df

def pipeline_3A_1_filter(df_bom, df_stock):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    if df_stock is None or df_stock.empty: return df_bom.copy()
    cols = list(df_stock.columns)
    if len(cols) < 3: return df_bom.copy()
    df_stock = df_stock.rename(columns={cols[0]:"Component", cols[2]:"Comment"})
    excluded = df_stock[df_stock["Comment"].astype(str).str.lower().str.strip()=="no need"]["Component"].astype(str)
    excluded_norm = excluded.str.upper().str.replace(" ","").str.strip().unique()
    df = df_bom.copy()
    df["Norm_Type"] = df["Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    out = df[~df["Norm_Type"].isin(excluded_norm)].reset_index(drop=True)
    return out.drop(columns=["Norm_Type"])

def pipeline_3A_2_accessories(df_bom, df_acc):
    if df_acc is None or df_acc.empty: return df_bom
    df_out = df_bom.copy()
    for _,row in df_bom.iterrows():
        main_item = str(row["Type"]).strip()
        matches = df_acc[df_acc.iloc[:,0].astype(str).str.strip()==main_item]
        for _,acc_row in matches.iterrows():
            acc_vals = acc_row.values[1:]
            for i in range(0,len(acc_vals),3):
                if i+2 >= len(acc_vals) or pd.isna(acc_vals[i]): break
                acc_item = str(acc_vals[i]).strip()
                try: acc_qty = float(str(acc_vals[i+1]).replace(",",".")) 
                except: acc_qty = 1
                acc_manuf = str(acc_vals[i+2]).strip()
                df_out = pd.concat([df_out,pd.DataFrame([{"Type":acc_item,"Quantity":acc_qty,"Manufacturer":acc_manuf,"Source":"Accessory"}])],ignore_index=True)
    return df_out

def pipeline_3A_3_nav(df_bom, df_part_no):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    if df_part_no is None or df_part_no.empty:
        df_bom["No."] = ""
        return df_bom
    df_part = df_part_no.copy().reset_index(drop=True).rename(columns=lambda c:str(c).strip())
    if "PartName_B" not in df_part.columns or "PartNo_A" not in df_part.columns:
        df_bom["No."] = ""
        return df_bom
    df_part["Norm_B"] = df_part["PartName_B"].astype(str).str.upper().str.replace(" ","").str.strip()
    def norm_partno(x):
        try: return str(int(float(str(x).strip().replace(",","."))))
        except: return str(x).strip()
    df_part["PartNo_A"] = df_part["PartNo_A"].map(norm_partno).fillna("").astype(str)
    df_part = df_part.drop_duplicates(subset=["Norm_B"],keep="first").drop_duplicates(subset=["PartNo_A"],keep="first")
    df = df_bom.copy()
    df["Norm_Type"] = df["Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    map_by_type = dict(zip(df_part["Norm_B"], df_part["PartNo_A"]))
    df["No."] = df["Norm_Type"].map(map_by_type).fillna("").astype(str)
    merge_cols = [c for c in ["PartNo_A","Desc_C","Manufacturer_D","SupplierNo_E","UnitPrice_F","Norm_B"] if c in df_part.columns]
    if merge_cols:
        df = df.merge(df_part[merge_cols], left_on="No.", right_on="PartNo_A", how="left")
        df = df.rename(columns={"Desc_C":"Description","Manufacturer_D":"Supplier","SupplierNo_E":"Supplier No.","UnitPrice_F":"Unit Cost"})
        df = df.drop(columns=[c for c in ["Norm_Type","Norm_B","PartNo_A"] if c in df.columns], errors="ignore")
    else:
        df = df.drop(columns=["Norm_Type"], errors="ignore")
    return df

def pipeline_3A_4_stock(df_bom, ks_file):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    if isinstance(ks_file,pd.DataFrame): df_stock = ks_file.copy()
    else: df_stock = pd.read_excel(io.BytesIO(ks_file.getvalue()),engine="openpyxl")
    df_stock = df_stock.rename(columns=lambda c:str(c).strip())
    df_stock = df_stock[[df_stock.columns[2],df_stock.columns[1],df_stock.columns[3]]]
    df_stock.columns = ["No.","Bin Code","Quantity"]
    df_stock["No."] = df_stock["No."].apply(normalize_no)
    df_bom["No."]   = df_bom["No."].apply(normalize_no)
    stock_groups = {k:v for k,v in df_stock.groupby("No.")}
    df_bom["Stock Rows"] = df_bom["No."].map(stock_groups)
    return df_bom

def pipeline_3A_5_tables(df_bom, project_number, df_part_no):
    rows = []
    for _, row in df_bom.iterrows():
        no = row.get("No.")
        qty = float(row.get("Quantity", 0) or 0)
        stock_rows = row.get("Stock Rows")

        if not isinstance(stock_rows, pd.DataFrame) or stock_rows.empty:
            rows.append({
                "Type": "Item",
                "No.": no,
                "Document No.": f"{project_number}/N",
                "Job No.": project_number,
                "Job Task No.": 1144,
                "Quantity": int(qty),
                "Location Code": "KAUNAS",
                "Bin Code": "",
                "Description": row.get("Description", ""),
                "Original Type": row.get("Original Type", "")
            })
            continue

        allocations = allocate_from_stock(no, qty, stock_rows)
        for alloc in allocations:
            rows.append({
                "Type": "Item",
                "No.": no,
                "Document No.": project_number,
                "Job No.": project_number,
                "Job Task No.": 1144,
                "Quantity": alloc["Allocated Qty"],
                "Location Code": "KAUNAS" if alloc["Bin Code"] else "",
                "Bin Code": alloc["Bin Code"],
                "Description": row.get("Description", ""),
                "Original Type": row.get("Original Type", "")
            })

    job_journal = pd.DataFrame(rows)

    supplier_map = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["Manufacturer_D"].astype(str)))

    tmp = df_bom.copy()
    if "Quantity" not in tmp.columns: tmp["Quantity"] = 0
    if "Description" not in tmp.columns: tmp["Description"] = ""
    tmp["No."] = tmp["No."].astype(str)
    tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0)

    nav_rows = []
    for _, r in tmp.iterrows():
        part_no = str(r["No."])
        qty = float(r.get("Quantity", 0) or 0)
        manuf = manuf_map.get(part_no, "")
        profit = 10 if "DANFOSS" in str(manuf).upper() else 17
        supplier = supplier_map.get(part_no, 30093)
        nav_rows.append({
            "Type": "Item",
            "No.": part_no,
            "Quantity": qty,
            "Supplier": supplier,
            "Profit": profit,
            "Discount": 0,
            "Description": r.get("Description", "")
        })

    nav_table = pd.DataFrame(nav_rows, columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])
    return job_journal, nav_table, df_bom

def pipeline_3B_0_prepare_cubic(df_cubic, df_part_code, extras=None):
    if df_cubic is None or df_cubic.empty: return pd.DataFrame()
    df = df_cubic.copy().rename(columns=lambda c: str(c).strip())
    if any(col in df.columns for col in ["E", "F", "G"]): df["Quantity"] = df[["E", "F", "G"]].bfill(axis=1).iloc[:, 0]
    if "Quantity" not in df.columns: df["Quantity"] = 0
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    if "Item Id" in df.columns:
        df["Type"] = df["Item Id"].astype(str).str.strip()
        df["Original Type"] = df["Type"]
    if "Type" not in df.columns: df["Type"] = ""; df["Original Type"] = ""
    if "No." not in df.columns: df["No."] = df["Type"]
    if df_part_code is not None and not df_part_code.empty:
        rename_map = dict(zip(df_part_code.iloc[:, 0].astype(str).str.strip(), df_part_code.iloc[:, 1].astype(str).str.strip()))
        for col in ["Type", "Original Type"]:
            if col in df.columns: df[col] = df[col].astype(str).str.strip().replace(rename_map)
    if extras: df = add_extra_components(df, [e for e in extras if e.get("target") == "cubic"])
    return df

def pipeline_3B_1_filtering(df_cubic,df_stock):
    if df_cubic is None or df_cubic.empty: return pd.DataFrame(),pd.DataFrame()
    if df_stock is None or df_stock.empty: return df_cubic.copy(),df_cubic.copy()
    cols=list(df_stock.columns)
    if len(cols)<3: return df_cubic.copy(),df_cubic.copy()
    df_stock=df_stock.rename(columns={cols[0]:"Component",cols[2]:"Comment"})
    excluded_all=df_stock[df_stock["Comment"].astype(str).str.strip()!=""]["Component"].astype(str)
    excluded_no_need=df_stock[df_stock["Comment"].astype(str).str.lower().str.strip()=="no need"]["Component"].astype(str)
    excluded_all_norm=excluded_all.str.upper().str.replace(" ","").str.strip().unique()
    excluded_no_need_norm=excluded_no_need.str.upper().str.replace(" ","").str.strip().unique()
    df=df_cubic.copy()
    df["Norm_Type"]=df["Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    df_journal=df[~df["Norm_Type"].isin(excluded_all_norm)].reset_index(drop=True)
    df_nav=df[~df["Norm_Type"].isin(excluded_no_need_norm)].reset_index(drop=True)
    return df_journal.drop(columns=["Norm_Type"]),df_nav.drop(columns=["Norm_Type"])

def pipeline_3B_2_accessories(df,df_acc):
    if df_acc is None or df_acc.empty: return df
    df_out=df.copy()
    for _,row in df.iterrows():
        main_item=str(row["Type"]).strip()
        matches=df_acc[df_acc.iloc[:,0].astype(str).str.strip()==main_item]
        for _,acc_row in matches.iterrows():
            acc_vals=acc_row.values[1:]
            for i in range(0,len(acc_vals),3):
                if i+2>=len(acc_vals) or pd.isna(acc_vals[i]): break
                acc_item=str(acc_vals[i]).strip()
                try: acc_qty=float(str(acc_vals[i+1]).replace(",",".")) 
                except: acc_qty=1
                acc_manuf=str(acc_vals[i+2]).strip()
                df_out=pd.concat([df_out,pd.DataFrame([{"Type":acc_item,"Quantity":acc_qty,"Manufacturer":acc_manuf,"Source":"Accessory"}])],ignore_index=True)
    return df_out

def pipeline_3B_3_nav(df,df_part_no): return pipeline_3A_3_nav(df,df_part_no)
def pipeline_3B_4_stock(df_journal,ks_file): return pipeline_3A_4_stock(df_journal,ks_file)
def pipeline_3B_5_tables(df_journal, df_nav, project_number, df_part_no):
    # Job Journal su /N logika
    rows = []
    for _, row in df_journal.iterrows():
        no = row.get("No.")
        qty = float(row.get("Quantity", 0) or 0)
        stock_rows = row.get("Stock Rows")

        if not isinstance(stock_rows, pd.DataFrame) or stock_rows.empty:
            rows.append({
                "Type": "Item",
                "No.": no,
                "Document No.": f"{project_number}/N",
                "Job No.": project_number,
                "Job Task No.": 1144,
                "Quantity": int(qty),
                "Location Code": "KAUNAS",
                "Bin Code": "",
                "Description": row.get("Description", ""),
                "Original Type": row.get("Original Type", "")
            })
            continue

        allocations = allocate_from_stock(no, qty, stock_rows)
        for alloc in allocations:
            rows.append({
                "Type": "Item",
                "No.": no,
                "Document No.": project_number,
                "Job No.": project_number,
                "Job Task No.": 1144,
                "Quantity": alloc["Allocated Qty"],
                "Location Code": "KAUNAS" if alloc["Bin Code"] else "",
                "Bin Code": alloc["Bin Code"],
                "Description": row.get("Description", ""),
                "Original Type": row.get("Original Type", "")
            })

    job_journal = pd.DataFrame(rows)

    # NAV table iš Project BOM logikos
    _, nav_table, _ = pipeline_3A_5_tables(df_nav, project_number, df_part_no)
    return job_journal, nav_table, df_nav

def pipeline_4_1_calculation(df_bom, df_cubic, df_hours, panel_type, grounding, project_number, df_instr=None):
    if df_bom is None: df_bom = pd.DataFrame()
    if df_cubic is None: df_cubic = pd.DataFrame()
    if df_hours is None: df_hours = pd.DataFrame()
    if not df_bom.empty and "Quantity" in df_bom and "Unit Cost" in df_bom:
        parts_cost = (pd.to_numeric(df_bom["Quantity"], errors="coerce").fillna(0) * pd.to_numeric(df_bom["Unit Cost"], errors="coerce").fillna(0)).sum()
    else: parts_cost = 0
    if not df_cubic.empty and "Quantity" in df_cubic and "Unit Cost" in df_cubic:
        cubic_cost = (pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0) * pd.to_numeric(df_cubic["Unit Cost"], errors="coerce").fillna(0)).sum()
    else: cubic_cost = 0
    hours_cost = 0
    if not df_hours.empty and df_hours.shape[1] > 4:
        hourly_rate = pd.to_numeric(df_hours.iloc[1,4], errors="coerce")
        row = df_hours[df_hours.iloc[:,0].astype(str).str.upper() == str(panel_type).upper()]
        if not row.empty:
            if grounding == "TT": h = pd.to_numeric(row.iloc[0,1], errors="coerce")
            elif grounding == "TN-S": h = pd.to_numeric(row.iloc[0,2], errors="coerce")
            else: h = pd.to_numeric(row.iloc[0,3], errors="coerce")
            hours_cost = (h if pd.notna(h) else 0) * (hourly_rate if pd.notna(hourly_rate) else 0)
    smart_supply = 9750.0
    wire_set = 2500.0
    total = parts_cost + cubic_cost + hours_cost + smart_supply + wire_set
    project_size = ""
    pallet_size = ""
    if df_instr is not None and not df_instr.empty:
        row = df_instr[df_instr.iloc[:,0].astype(str).str.upper() == str(panel_type).upper()]
        if not row.empty:
            project_size = str(row.iloc[0,1])
            pallet_size = str(row.iloc[0,2])
    df_calc = pd.DataFrame([
        {"Label":"Parts","Value":parts_cost},
        {"Label":"Cubic","Value":cubic_cost},
        {"Label":"Hours cost","Value":hours_cost},
        {"Label":"Smart supply","Value":smart_supply},
        {"Label":"Wire set","Value":wire_set},
        {"Label":"Extra","Value":0},
        {"Label":"Total","Value":total},
        {"Label":"Total+5%","Value":total*1.05},
        {"Label":"Total+35%","Value":total*1.35},
        {"Label":"Project size","Value":project_size},
        {"Label":"Pallet size","Value":pallet_size},
    ])
    return df_calc

def pipeline_4_2_missing_nav(df, source):
    if df is None or df.empty or "No." not in df.columns: return pd.DataFrame()
    missing = df[df["No."].astype(str).str.strip()==""] if not df.empty else pd.DataFrame()
    if missing.empty: return pd.DataFrame()
    qty = pd.to_numeric(missing["Quantity"], errors="coerce").fillna(0).astype(int) if "Quantity" in missing else 0
    return pd.DataFrame({"Source": source,"Original Article": missing.get("Original Article",""),"Original Type": missing.get("Original Type",""),"Quantity": qty,"NAV No.": missing["No."]})

# =====================================================
# Render
# =====================================================
def render():
    st.header("Stage 3: BOM Management")

    inputs = pipeline_2_1_user_inputs()
    if not inputs:
        return

    files = pipeline_2_2_file_uploads(inputs["rittal"])
    if not files:
        return

    required_A = ["bom", "data", "ks"]
    required_B = ["cubic_bom", "data", "ks"] if not inputs["rittal"] else []
    miss_A = [k for k in required_A if k not in files]
    miss_B = [k for k in required_B if k not in files]

    st.subheader("📋 Required files")
    col1, col2 = st.columns(2)
    with col1:
        st.success("Project BOM: OK") if not miss_A else st.warning(f"Project BOM missing: {miss_A}")
    with col2:
        if not inputs["rittal"]:
            st.success("CUBIC BOM: OK") if not miss_B else st.warning(f"CUBIC BOM missing: {miss_B}")
        else:
            st.info("CUBIC BOM skipped (Rittal)")

    if st.button("🚀 Run Processing"):
        st.session_state["processing_started"] = True
        st.session_state["mech_confirmed"] = False
        st.session_state["df_mech"] = pd.DataFrame()
        st.session_state["df_remain"] = pd.DataFrame()

    if not st.session_state.get("processing_started", False):
        st.stop()

    data_book = files.get("data", {})
    df_stock   = pipeline_2_3_get_sheet_safe(data_book, ["Stock"])
    df_part_no = pipeline_2_4_normalize_part_no(
        pipeline_2_3_get_sheet_safe(data_book, ["Part_no", "Parts_no", "Part no"])
    )
    df_hours   = pipeline_2_3_get_sheet_safe(data_book, ["Hours"])
    df_acc     = pipeline_2_3_get_sheet_safe(data_book, ["Accessories"])
    df_code    = pipeline_2_3_get_sheet_safe(data_book, ["Part_code"])
    df_instr   = pipeline_2_3_get_sheet_safe(data_book, ["Instructions"])
    df_main_sw = pipeline_2_3_get_sheet_safe(data_book, ["main_switch"])

    extras = []
    if inputs["ups"]:
        extras.append({"type": "LI32111CT01", "qty": 1, "target": "bom", "force_no": "2214036"})
        extras.append({"type": "ADV UPS holder V3", "qty": 1, "target": "bom", "force_no": "2214035"})
        extras.append({"type": "268-2610", "qty": 1, "target": "bom", "force_no": "1865206"})
    if inputs["swing_frame"]:
        extras.append({"type": "9030+2970", "qty": 1, "target": "cubic", "force_no": "2185835"})
    if df_instr is not None and not df_instr.empty:
        row = df_instr[df_instr.iloc[:,0].astype(str).str.upper() == str(inputs["panel_type"]).upper()]
        if not row.empty:
            if inputs["panel_type"][0] not in ["F","G"]:
                try: qty_sdd = int(pd.to_numeric(row.iloc[0,4], errors="coerce").fillna(0))
                except: qty_sdd = 0
                if qty_sdd > 0:
                    st.info(f"🔹 According to Instructions: need {qty_sdd} × SDD07550")
                    extras.append({"type": "SDD07550","qty": qty_sdd,"target": "cubic","force_no": "SDD07550"})
            for col_idx in range(5,10):  # F..J stulpeliai
                if col_idx < row.shape[1]:
                    val = str(row.iloc[0,col_idx]).strip()
                    if val and val.lower() != "nan":
                        st.info(f"🔹 According to Instructions: need 1 × {val}")
                        extras.append({"type": val,"qty": 1,"target": "cubic"})

    job_A, nav_A, df_bom_proc = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    job_B, nav_B, df_cub_proc = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if not miss_A:
        df_bom = pipeline_3A_0_rename(files["bom"], df_code, extras)
        df_bom = pipeline_3A_1_filter(df_bom, df_stock)
        df_bom = pipeline_3A_2_accessories(df_bom, df_acc)
        df_bom = pipeline_3A_3_nav(df_bom, df_part_no)
        df_bom = pipeline_3A_4_stock(df_bom, files["ks"])
        job_A, nav_A, df_bom_proc = pipeline_3A_5_tables(df_bom, inputs["project_number"], df_part_no)

    if not inputs["rittal"] and not miss_B:
        df_cubic = pipeline_3B_0_prepare_cubic(files["cubic_bom"], df_code, extras)
        df_j, df_n = pipeline_3B_1_filtering(df_cubic, df_stock)
        df_j = pipeline_3B_2_accessories(df_j, df_acc)
        df_n = pipeline_3B_2_accessories(df_n, df_acc)
        df_j = pipeline_3B_3_nav(df_j, df_part_no)
        df_n = pipeline_3B_3_nav(df_n, df_part_no)
        df_j = pipeline_3B_4_stock(df_j, files["ks"])
        job_B, nav_B, df_cub_proc = pipeline_3B_5_tables(df_j, df_n, inputs["project_number"], df_part_no)

    # --- MAIN SWITCH accessories ---
    if df_main_sw is not None and not df_main_sw.empty:
        row = df_main_sw[df_main_sw.iloc[:,1].astype(str).str.strip().str.upper() == str(inputs["main_switch"]).upper()]
        if not row.empty:
            for col_idx in range(2, 12):  # C–L stulpeliai
                if col_idx < row.shape[1]:
                    val = str(row.iloc[0, col_idx]).strip()
                    if val and val.lower() != "nan":
                        norm_val = val.upper().replace(" ", "")
                        part_match = None
                        if not df_part_no.empty and "PartName_B" in df_part_no.columns:
                            df_part_no["Norm_B"] = df_part_no["PartName_B"].astype(str).str.upper().str.replace(" ","").str.strip()
                            part_match = df_part_no[df_part_no["Norm_B"] == norm_val]
                        if part_match is not None and not part_match.empty:
                            no_val = str(part_match.iloc[0]["PartNo_A"])
                            desc   = str(part_match.iloc[0].get("Desc_C",""))
                            supp   = str(part_match.iloc[0].get("SupplierNo_E",""))
                        else:
                            no_val, desc, supp = "","",""
                        stock_rows = {}
                        if "ks" in files:
                            df_stock2 = files["ks"].copy()
                            df_stock2 = df_stock2.rename(columns=lambda c:str(c).strip())
                            df_stock2 = df_stock2[[df_stock2.columns[2], df_stock2.columns[1], df_stock2.columns[3]]]
                            df_stock2.columns = ["No.","Bin Code","Quantity"]
                            df_stock2["No."] = df_stock2["No."].apply(normalize_no)
                            stock_groups = {k:v for k,v in df_stock2.groupby("No.")}
                            stock_rows = stock_groups.get(no_val, pd.DataFrame(columns=["Bin Code","Quantity"]))
                        allocations = allocate_from_stock(no_val, 1, stock_rows)
                        for alloc in allocations:
                            job_A = pd.concat([job_A, pd.DataFrame([{
                                "Type": val,"Original Type": val,"No.": no_val,
                                "Document No.": inputs["project_number"] if no_val else inputs["project_number"]+"/N",
                                "Job No.": inputs["project_number"],"Job Task No.": 1144,
                                "Quantity": alloc["Allocated Qty"],
                                "Location Code": "KAUNAS" if alloc["Bin Code"] else "",
                                "Bin Code": alloc["Bin Code"],
                                "Description": desc,"Source": "Main switch accessory"
                            }])], ignore_index=True)
                        nav_A = pd.concat([nav_A, pd.DataFrame([{
                            "Type": "Item","No.": no_val,"Quantity": 1,
                            "Supplier": supp if supp else 30093,"Profit": 17,"Discount": 0,"Description": desc
                        }])], ignore_index=True)

    if not st.session_state.get("mech_confirmed", False):
        if not job_B.empty:
            st.subheader("📑 Job Journal (CUBIC BOM → allocate to Mechanics)")
            editable = job_B.copy()
            editable["Available Qty"] = editable["Quantity"].astype(int)
            mech_inputs = []
            with st.form("mech_form", clear_on_submit=False):
                for idx, row in editable.iterrows():
                    cols = st.columns([2, 3, 4, 2, 2])
                    cols[0].write(str(row.get("No.", "")))
                    cols[1].write(str(row.get("Original Type", "")))
                    cols[2].write(str(row.get("Description", "")))
                    cols[3].write(int(row["Available Qty"]))
                    take = cols[4].number_input("",min_value=0,max_value=int(row["Available Qty"]),step=1,format="%d",key=f"take_{idx}")
                    mech_inputs.append((idx, take))
                confirm = st.form_submit_button("✅ Confirm Mechanics Allocation")
            if confirm:
                mech_rows, remain_rows = [], []
                for idx, take in mech_inputs:
                    avail = int(editable.loc[idx, "Available Qty"])
                    r = editable.loc[idx].to_dict()
                    if take > 0: mech_rows.append({**r, "Quantity": take})
                    remain_qty = avail - take
                    if remain_qty > 0 and str(r.get("No.", "")) != "2185835":
                        remain_rows.append({**r, "Quantity": remain_qty})
                st.session_state["df_mech"] = pd.DataFrame(mech_rows)
                st.session_state["df_remain"] = pd.DataFrame(remain_rows)
                st.session_state["mech_confirmed"] = True
                if inputs["swing_frame"]:
                    swing_row = pd.DataFrame([{
                        "Type": "9030+2970","Original Type": "9030+2970","No.": "2185835",
                        "Quantity": 1,"Document No.": inputs["project_number"],"Job No.": inputs["project_number"],
                        "Job Task No.": 1144,"Location Code": "KAUNAS","Bin Code": "",
                        "Description": "Swing frame component","Source": "Extra"
                    }])
                    st.session_state["df_mech"] = pd.concat([st.session_state["df_mech"], swing_row],ignore_index=True)
        st.stop()

    if "df_mech" in st.session_state and not st.session_state["df_mech"].empty:
        st.subheader("📑 Job Journal (CUBIC BOM TO MECH.)")
        st.dataframe(st.session_state["df_mech"], use_container_width=True)
    if "df_remain" in st.session_state and not st.session_state["df_remain"].empty:
        st.subheader("📑 Job Journal (CUBIC BOM REMAINING)")
        st.dataframe(st.session_state["df_remain"], use_container_width=True)
    if not job_A.empty:
        st.subheader("📑 Job Journal (Project BOM)")
        st.dataframe(job_A, use_container_width=True)
    if not nav_A.empty:
        st.subheader("🛒 NAV Table (Project BOM)")
        st.dataframe(nav_A, use_container_width=True)
    if not nav_B.empty:
        st.subheader("🛒 NAV Table (CUBIC BOM)")
        st.dataframe(nav_B, use_container_width=True)

    calc = pipeline_4_1_calculation(df_bom_proc,df_cub_proc,df_hours,inputs["panel_type"],inputs["grounding"],inputs["project_number"],df_instr)
    st.subheader("💰 Calculation")
    st.dataframe(calc, use_container_width=True)

    miss_nav_A = pipeline_4_2_missing_nav(df_bom_proc, "Project BOM")
    miss_nav_B = pipeline_4_2_missing_nav(df_cub_proc, "CUBIC BOM")
    if not miss_nav_A.empty or not miss_nav_B.empty:
        st.subheader("⚠️ Missing NAV Numbers")
        if not miss_nav_A.empty: st.dataframe(miss_nav_A, use_container_width=True)
        if not miss_nav_B.empty: st.dataframe(miss_nav_B, use_container_width=True)

# =====================================================
# Export to Excel
# =====================================================
if st.button("💾 Export Results to Excel"):
    ts = datetime.datetime.now().strftime("%Y%m%d%H%M")
    pallet_size = ""
    if "df_calc" in locals():
        try:
            pallet_size = str(calc[calc["Label"]=="Pallet size"]["Value"].iloc[0])
        except:
            pallet_size = ""

    filename = f"{inputs['project_number']}_{inputs['panel_type']}_{inputs['grounding']}_{pallet_size}_{ts}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Info"
    ws.append(["Project number", inputs["project_number"]])
    ws.append(["Panel type", inputs["panel_type"]])
    ws.append(["Grounding", inputs["grounding"]])
    ws.append(["Main switch", inputs["main_switch"]])
    ws.append(["Swing frame", inputs["swing_frame"]])
    ws.append(["UPS", inputs["ups"]])
    ws.append(["Rittal", inputs["rittal"]])

    def add_df_to_wb(df, title):
        if df is None or df.empty: 
            return
        ws = wb.create_sheet(title)
        ws.append(df.columns.tolist())
        for _, row in df.iterrows():
            ws.append(row.tolist())

    add_df_to_wb(job_A, "JobJournal_ProjectBOM")
    add_df_to_wb(nav_A, "NAV_ProjectBOM")
    add_df_to_wb(job_B, "JobJournal_CUBICBOM")
    add_df_to_wb(nav_B, "NAV_CUBICBOM")
    if "df_mech" in st.session_state: add_df_to_wb(st.session_state["df_mech"], "Mech")
    if "df_remain" in st.session_state: add_df_to_wb(st.session_state["df_remain"], "Remaining")
    add_df_to_wb(calc, "Calculation")

    # Missing NAV tables
    add_df_to_wb(miss_nav_A, "MissingNAV_ProjectBOM")
    add_df_to_wb(miss_nav_B, "MissingNAV_CUBICBOM")

    # Save and download
    save_path = f"/tmp/{filename}"
    wb.save(save_path)

    with open(save_path, "rb") as f:
        st.download_button(
            label="⬇️ Download Excel",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
