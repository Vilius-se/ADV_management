import streamlit as st
import pandas as pd
import re
import io

# =====================================================
# 1.x – Helpers
# =====================================================


def _dbg(df, label, debug=False):
    if not debug: 
        return
    st.markdown(f"### 🔎 Debug: {label}")
    if df is None or df.empty:
        st.info("Empty DataFrame")
        return
    st.text(f"Shape: {df.shape}")
    st.dataframe(df.head(60), use_container_width=True)

def add_extra_components(df, extras):
    if df is None: 
        df = pd.DataFrame()
    df_out = df.copy()

    for e in extras:
        extra_row = pd.DataFrame([{
            "Original Type": e["type"],
            "Type": e["type"],
            "Quantity": e.get("qty", 1),
            "Source": "Extra"
        }])
        df_out = pd.concat([df_out, extra_row], ignore_index=True)

    return df_out

def build_nav_table_from_bom(df_bom: pd.DataFrame, df_part_no: pd.DataFrame,
                             label: str = "Project BOM", debug: bool = False) -> pd.DataFrame:
    _dbg(debug, f"{label} → NAV: input df_bom", df_bom)

    req = ["PartNo_A", "SupplierNo_E", "Manufacturer_D"]
    if df_part_no is None or df_part_no.empty or any(c not in df_part_no.columns for c in req):
        _dbg(debug, f"{label} → NAV: Part_no trūksta reikiamų stulpelių {req}")
        return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])

    supplier_map = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["Manufacturer_D"].astype(str)))
    _dbg(debug, f"{label} → NAV: supplier_map (sample)", supplier_map)
    _dbg(debug, f"{label} → NAV: manuf_map (sample)", manuf_map)

    tmp = df_bom.copy()
    if "Quantity" not in tmp.columns: tmp["Quantity"] = 0
    if "Description" not in tmp.columns: tmp["Description"] = ""
    if "No." not in tmp.columns: tmp["No."] = ""
    tmp["No."] = tmp["No."].astype(str)
    tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0)
    _dbg(debug, f"{label} → NAV: normalized tmp", tmp)

    nav_rows = []
    m_total = len(tmp)
    m_profit10 = 0
    m_supplier_default = 0
    m_missing_no = 0

    for _, r in tmp.iterrows():
        part_no = str(r["No."]).strip()
        qty = float(r.get("Quantity", 0) or 0)
        manuf = manuf_map.get(part_no, "")

        if part_no == "" or part_no.lower() == "nan":
            m_missing_no += 1

        profit = 10 if "DANFOSS" in str(manuf).upper() else 17
        if profit == 10:
            m_profit10 += 1

        supplier = supplier_map.get(part_no, 30093)
        if part_no not in supplier_map:
            m_supplier_default += 1

        nav_rows.append({
            "Type": "Item",
            "No.": part_no,
            "Quantity": qty,
            "Supplier": supplier,
            "Profit": profit,
            "Discount": 0,
            "Description": r.get("Description", "")
        })

    _dbg(debug, f"{label} → NAV: metrics", {
        "rows_total": m_total,
        "profit10_cnt": m_profit10,
        "supplier_default_cnt": m_supplier_default,
        "missing_No_cnt": m_missing_no
    })

    nav_table = pd.DataFrame(
        nav_rows,
        columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"]
    )
    _dbg(debug, f"{label} → NAV: final table", nav_table)
    return nav_table




def pipeline_1_1_norm_name(x):
    return ''.join(str(x).upper().split())

def pipeline_1_2_parse_qty(x):
    if pd.isna(x): 
        return 0.0
    if isinstance(x,(int,float)): 
        return float(x)
    s = str(x).strip().replace('\xa0','').replace(' ','')
    if ',' in s and '.' in s: 
        s = s.replace(',','')
    else: 
        s = s.replace('.','').replace(',','.')
    try: 
        return float(s)
    except: 
        return 0.0
def pipeline_1_3_safe_filename(s):
    s = '' if s is None else str(s).strip()
    s = re.sub(r'[\\/:*?"<>|]+','',s)
    return s.replace(' ','_')
def pipeline_1_4_normalize_no(x):
    try: 
        return str(int(float(str(x).replace(",","." ).strip())))
    except: 
        return str(x).strip()
def read_excel_any(file,**kwargs):
    try: 
        return pd.read_excel(file,engine="openpyxl",**kwargs)
    except: 
        return pd.read_excel(file,engine="xlrd",**kwargs)
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
    if remaining>0:
        allocations.append({"No.":no,"Bin Code":"","Allocated Qty":remaining})
    return allocations
# --- Backward-compat alias ---
normalize_no = pipeline_1_4_normalize_no


# =====================================================
# 2.x – Inputs & File Uploads
# =====================================================

def pipeline_2_1_user_inputs():
    st.subheader("Project Information")
    project_number = st.text_input("Project number (1234-567)")
    if project_number and not re.match(r"^\d{4}-\d{3}$", project_number):
        st.error("Invalid format (must be 1234-567)")
        return None

    panel_type = st.selectbox(
        "Panel type",
        ['A','B','B1','B2','C','C1','C2','C3','C4','C4.1','C5','C6','C7','C8',
         'F','F1','F2','F3','F4','F4.1','F5','F6','F7',
         'G','G1','G2','G3','G4','G5','G6','G7','Custom']
    )
    grounding = st.selectbox("Grounding type", ["TT","TN-S","TN-C-S"])
    main_switch = st.selectbox(
        "Main switch",
        ["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111",
         "31109","31107","C404400S","C634630S"]
    )
    swing_frame = st.checkbox("Swing frame?")
    ups = st.checkbox("UPS?")
    rittal = st.checkbox("Rittal?")

    return {
        "project_number": project_number,
        "panel_type": panel_type,
        "grounding": grounding,
        "main_switch": main_switch,
        "swing_frame": swing_frame,
        "ups": ups,
        "rittal": rittal
    }


def pipeline_2_2_get_sheet_safe(data_dict, names):
    if not isinstance(data_dict, dict): 
        return None
    for key in data_dict.keys():
        if str(key).strip().upper().replace(" ","_") in [n.upper().replace(" ","_") for n in names]:
            return data_dict[key]
    return None


def pipeline_2_3_normalize_part_no(df_raw):
    if df_raw is None or df_raw.empty: 
        return pd.DataFrame()
    df = df_raw.copy().rename(columns=lambda c:str(c).strip())
    col_map = {}
    if df.shape[1]>=1: col_map[df.columns[0]]="PartNo_A"
    if df.shape[1]>=2: col_map[df.columns[1]]="PartName_B"
    if df.shape[1]>=3: col_map[df.columns[2]]="Desc_C"
    if df.shape[1]>=4: col_map[df.columns[3]]="Manufacturer_D"
    if df.shape[1]>=5: col_map[df.columns[4]]="SupplierNo_E"
    if df.shape[1]>=6: col_map[df.columns[5]]="UnitPrice_F"
    return df.rename(columns=col_map)


def pipeline_2_4_file_uploads(rittal=False):
    st.subheader("Upload Required Files")
    dfs = {}

    # --- CUBIC BOM ---
    if not rittal:
        cubic_bom = st.file_uploader("Insert CUBIC BOM", type=["xls","xlsx","xlsm"], key="cubic_bom")
        if cubic_bom:
            df_cubic = read_excel_any(cubic_bom, skiprows=13, usecols="B,E:F,G")
            df_cubic = df_cubic.rename(columns=lambda c:str(c).strip())
            if {"E","F","G"}.issubset(df_cubic.columns):
                df_cubic["Quantity"] = df_cubic[["E","F","G"]].bfill(axis=1).iloc[:,0]
            elif "Quantity" not in df_cubic.columns:
                df_cubic["Quantity"] = 0
            df_cubic["Quantity"] = pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0)
            df_cubic = df_cubic.rename(columns={"Item Id":"Type"})
            df_cubic["Original Type"] = df_cubic["Type"]
            df_cubic["No."] = df_cubic["Type"]
            dfs["cubic_bom"] = df_cubic

    # --- BOM ---
    bom = st.file_uploader("Insert BOM", type=["xls","xlsx","xlsm"], key="bom")
    if bom:
        df_bom = read_excel_any(bom)
        if df_bom.shape[1] >= 2:
            colA = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
            colB = df_bom.iloc[:,1].fillna("").astype(str).str.strip()
            df_bom["Original Article"] = colA
            df_bom["Original Type"] = colB.where(colB!="", colA)
        else:
            df_bom["Original Article"] = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
            df_bom["Original Type"] = df_bom["Original Article"]
        dfs["bom"] = df_bom

    # --- DATA ---
    data_file = st.file_uploader("Insert DATA", type=["xls","xlsx","xlsm"], key="data")
    if data_file:
        dfs["data"] = pd.read_excel(data_file, sheet_name=None)

    # --- Kaunas Stock ---
    ks_file = st.file_uploader("Insert Kaunas Stock", type=["xls","xlsx","xlsm"], key="ks")
    if ks_file:
        dfs["ks"] = read_excel_any(ks_file)

    return dfs


# =====================================================
# 3A – Project BOM (su debug)
# =====================================================

def pipeline_3A_0_rename(df_bom, df_part_code, extras=None, debug=False):
    if df_bom is None or df_bom.empty: 
        return pd.DataFrame()

    if df_part_code is not None and not df_part_code.empty:
        rename_map = dict(zip(
            df_part_code.iloc[:,0].astype(str).str.strip(),
            df_part_code.iloc[:,1].astype(str).str.strip()
        ))
        df_bom = df_bom.rename(columns=rename_map)

    if "Type" not in df_bom.columns: 
        df_bom["Type"] = df_bom.iloc[:,0].astype(str)
    if "Original Type" not in df_bom.columns: 
        df_bom["Original Type"] = df_bom["Type"]
    if "Original Article" not in df_bom.columns: 
        df_bom["Original Article"] = df_bom.iloc[:,0].astype(str)

    if extras:
        df_bom = add_extra_components(df_bom, [e for e in extras if e["target"]=="bom"])

    if debug: st.write("🔧 After 3A_0_rename + extras:", df_bom.head(10))
    return df_bom


def pipeline_3A_1_filter(df_bom, df_stock, debug=False):
    if df_bom is None or df_bom.empty: 
        return pd.DataFrame()
    if df_stock is None or df_stock.empty: 
        return df_bom.copy()
    cols = list(df_stock.columns)
    if len(cols) < 3: 
        return df_bom.copy()

    df_stock = df_stock.rename(columns={cols[0]:"Component", cols[2]:"Comment"})
    excluded = df_stock[df_stock["Comment"].astype(str).str.lower().str.strip()=="no need"]["Component"].astype(str)
    excluded_norm = excluded.str.upper().str.replace(" ","").str.strip().unique()

    df = df_bom.copy()
    df["Norm_Type"] = df["Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    out = df[~df["Norm_Type"].isin(excluded_norm)].reset_index(drop=True)

    _dbg(out, "3A_1 Filtered BOM", debug=debug)
    return out.drop(columns=["Norm_Type"])


def pipeline_3A_2_accessories(df_bom, df_acc, debug=False):
    if df_acc is None or df_acc.empty: 
        return df_bom
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
                df_out = pd.concat([df_out,pd.DataFrame([{
                    "Type":acc_item,"Quantity":acc_qty,"Manufacturer":acc_manuf,"Source":"Accessory"
                }])],ignore_index=True)

    _dbg(df_out, "3A_2 With Accessories", debug=debug)
    return df_out


def pipeline_3A_3_nav(df_bom, df_part_no, debug=False):
    if df_bom is None or df_bom.empty: 
        return pd.DataFrame()
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
        df = df.rename(columns={
            "Desc_C":"Description",
            "Manufacturer_D":"Supplier",
            "SupplierNo_E":"Supplier No.",
            "UnitPrice_F":"Unit Cost"
        })
        df = df.drop(columns=[c for c in ["Norm_Type","Norm_B","PartNo_A"] if c in df.columns], errors="ignore")
    else:
        df = df.drop(columns=["Norm_Type"], errors="ignore")

    _dbg(df, "3A_3 With NAV numbers", debug=debug)
    return df


def pipeline_3A_4_stock(df_bom, ks_file, debug=False):
    if df_bom is None or df_bom.empty: 
        return pd.DataFrame()
    if isinstance(ks_file,pd.DataFrame): 
        df_stock = ks_file.copy()
    else: 
        df_stock = pd.read_excel(io.BytesIO(ks_file.getvalue()),engine="openpyxl")

    df_stock = df_stock.rename(columns=lambda c:str(c).strip())
    df_stock = df_stock[[df_stock.columns[2],df_stock.columns[1],df_stock.columns[3]]]
    df_stock.columns = ["No.","Bin Code","Quantity"]
    df_stock["No."] = df_stock["No."].apply(normalize_no)
    df_bom["No."]   = df_bom["No."].apply(normalize_no)
    stock_groups = {k:v for k,v in df_stock.groupby("No.")}
    df_bom["Stock Rows"] = df_bom["No."].map(stock_groups)

    _dbg(df_bom, "3A_4 With Stock info", debug=debug)
    return df_bom


def pipeline_3A_5_tables(df_bom, project_number, df_part_no, debug=False):
    rows=[]
    for _,row in df_bom.iterrows():
        no=row.get("No.")
        qty=float(row.get("Quantity",0) or 0)
        stock_rows=row.get("Stock Rows")
        if not isinstance(stock_rows,pd.DataFrame):
            stock_rows=pd.DataFrame(columns=["Bin Code","Quantity"])
        allocations=allocate_from_stock(no,qty,stock_rows)
        for alloc in allocations:
            rows.append({
                "Type":"Item","No.":no,"Document No.":project_number,"Job No.":project_number,
                "Job Task No.":1144,"Quantity":alloc["Allocated Qty"],
                "Location Code":"KAUNAS" if alloc["Bin Code"] else "",
                "Bin Code":alloc["Bin Code"],
                "Description":row.get("Description",""),
                "Original Type":row.get("Original Type","")
            })
    job_journal=pd.DataFrame(rows)

    supplier_map=dict(zip(df_part_no["PartNo_A"].astype(str),df_part_no["SupplierNo_E"]))
    manuf_map=dict(zip(df_part_no["PartNo_A"].astype(str),df_part_no["Manufacturer_D"].astype(str)))
    tmp=df_bom.copy()
    if "Quantity" not in tmp.columns: tmp["Quantity"]=0
    if "Description" not in tmp.columns: tmp["Description"]=""
    tmp["No."]=tmp["No."].astype(str)
    tmp["Quantity"]=pd.to_numeric(tmp["Quantity"],errors="coerce").fillna(0)

    nav_rows=[]
    for _,r in tmp.iterrows():
        part_no=str(r["No."]); qty=float(r.get("Quantity",0) or 0)
        manuf=manuf_map.get(part_no,"")
        profit=10 if "DANFOSS" in str(manuf).upper() else 17
        supplier=supplier_map.get(part_no,30093)
        nav_rows.append({
            "Type":"Item","No.":part_no,"Quantity":qty,"Supplier":supplier,
            "Profit":profit,"Discount":0,"Description":r.get("Description","")
        })
    nav_table=pd.DataFrame(nav_rows,columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])

    _dbg(job_journal, "3A_5 Job Journal", debug=debug)
    _dbg(nav_table, "3A_5 NAV Table", debug=debug)
    return job_journal,nav_table,df_bom
    
# =====================================================
# 3B – CUBIC BOM
# =====================================================

def pipeline_3B_0_prepare_cubic(df_cubic, df_part_code, extras=None, debug=False):
    if df_cubic is None or df_cubic.empty: 
        return pd.DataFrame()

    df = df_cubic.copy().rename(columns=lambda c:str(c).strip())

    if any(col in df.columns for col in ["E","F","G"]):
        df["Quantity"] = df[["E","F","G"]].bfill(axis=1).iloc[:,0]
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    elif "Quantity" in df.columns:
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    else: 
        df["Quantity"] = 0

    if "Item Id" in df.columns:
        df["Type"] = df["Item Id"].astype(str).str.strip()
        df["Original Type"] = df["Type"]
    elif "Type" not in df.columns:
        df["Type"] = ""
        df["Original Type"] = ""

    if "No." not in df.columns: 
        df["No."] = df["Type"]

    if df_part_code is not None and not df_part_code.empty:
        rename_map = dict(zip(
            df_part_code.iloc[:,0].astype(str).str.strip(),
            df_part_code.iloc[:,1].astype(str).str.strip()
        ))
        df = df.rename(columns=rename_map)

    if extras:
        df = add_extra_components(df, [e for e in extras if e["target"]=="cubic"])

    if debug: st.write("🔧 After 3B_0_prepare_cubic + extras:", df.head(10))
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
                df_out=pd.concat([df_out,pd.DataFrame([{
                    "Type":acc_item,"Quantity":acc_qty,"Manufacturer":acc_manuf,"Source":"Accessory"
                }])],ignore_index=True)
    return df_out

def pipeline_3B_3_nav(df,df_part_no):
    return pipeline_3A_3_nav(df,df_part_no)

def pipeline_3B_4_stock(df_journal,ks_file):
    return pipeline_3A_4_stock(df_journal,ks_file)

def pipeline_3B_5_tables(df_journal,df_nav,project_number,df_part_no):
    job_journal,_,_=pipeline_3A_5_tables(df_journal,project_number,df_part_no)
    _,nav_table,_=pipeline_3A_5_tables(df_nav,project_number,df_part_no)
    return job_journal,nav_table,df_nav

# =====================================================
# 4.x – Calculation & Missing NAV
# =====================================================

def pipeline_4_1_calculation(df_bom, df_cubic, df_hours, panel_type, grounding, project_number):
    if df_bom is None: df_bom = pd.DataFrame()
    if df_cubic is None: df_cubic = pd.DataFrame()
    if df_hours is None: df_hours = pd.DataFrame()

    if not df_bom.empty and "Quantity" in df_bom and "Unit Cost" in df_bom:
        parts_cost = (pd.to_numeric(df_bom["Quantity"], errors="coerce").fillna(0) *
                      pd.to_numeric(df_bom["Unit Cost"], errors="coerce").fillna(0)).sum()
    else:
        parts_cost = 0

    if not df_cubic.empty and "Quantity" in df_cubic and "Unit Cost" in df_cubic:
        cubic_cost = (pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0) *
                      pd.to_numeric(df_cubic["Unit Cost"], errors="coerce").fillna(0)).sum()
    else:
        cubic_cost = 0

    hours_cost = 0
    if not df_hours.empty and df_hours.shape[1] > 4:
        hourly_rate = pd.to_numeric(df_hours.iloc[1,4], errors="coerce")
        row = df_hours[df_hours.iloc[:,0].astype(str).str.upper() == str(panel_type).upper()]
        if not row.empty:
            if grounding == "TT":
                h = pd.to_numeric(row.iloc[0,1], errors="coerce")
            elif grounding == "TN-S":
                h = pd.to_numeric(row.iloc[0,2], errors="coerce")
            else:
                h = pd.to_numeric(row.iloc[0,3], errors="coerce")
            hours_cost = (h if pd.notna(h) else 0) * (hourly_rate if pd.notna(hourly_rate) else 0)

    smart_supply = 9750.0
    wire_set = 2500.0
    total = parts_cost + cubic_cost + hours_cost + smart_supply + wire_set

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
    ])
    return df_calc

def pipeline_4_2_missing_nav(df, source):
    if df is None or df.empty or "No." not in df.columns: return pd.DataFrame()
    missing = df[df["No."].astype(str).str.strip()==""] if not df.empty else pd.DataFrame()
    if missing.empty: return pd.DataFrame()
    qty = pd.to_numeric(missing["Quantity"], errors="coerce").fillna(0).astype(int) if "Quantity" in missing else 0
    return pd.DataFrame({
        "Source": source,
        "Original Article": missing.get("Original Article",""),
        "Original Type": missing.get("Original Type",""),
        "Quantity": qty,
        "NAV No.": missing["No."]
    })


def render(debug_flag=False):
    st.header("Stage 3: BOM Management")

    inputs = user_inputs()
    if not inputs: 
        return

    files = file_uploads(inputs["rittal"])
    if not files: 
        return

    required_A = ["bom","data","ks"]
    required_B = ["cubic_bom","data","ks"] if not inputs["rittal"] else []

    miss_A = [k for k in required_A if k not in files]
    miss_B = [k for k in required_B if k not in files]

    st.subheader("📋 Required files")
    col1,col2 = st.columns(2)
    with col1:
        st.success("Project BOM: OK") if not miss_A else st.warning(f"Project BOM missing: {miss_A}")
    with col2:
        if not inputs["rittal"]:
            st.success("CUBIC BOM: OK") if not miss_B else st.warning(f"CUBIC BOM missing: {miss_B}")
        else:
            st.info("CUBIC BOM skipped (Rittal)")

    if "ks" in files and not files["ks"].empty:
        st.subheader("🔎 Kaunas Stock preview")
        st.dataframe(files["ks"].head(20), use_container_width=True)

    if st.button("🚀 Run Processing"):
        data_book = files.get("data",{})
        df_stock   = get_sheet_safe(data_book,["Stock"])
        df_part_no = normalize_part_no(get_sheet_safe(data_book,["Part_no","Parts_no","Part no"]))
        df_hours   = get_sheet_safe(data_book,["Hours"])
        df_acc     = get_sheet_safe(data_book,["Accessories"])
        df_code    = get_sheet_safe(data_book,["Part_code"])

        # Extras pagal vartotojo pasirinkimus
        extras = []
        if inputs["ups"]:
            extras.append({"type": "LI32111CT01", "qty": 1, "target": "bom"})
        if inputs["swing_frame"]:
            extras.append({"type": "9030+2970", "qty": 1, "target": "cubic"})

        job_A, nav_A, df_bom_proc = pd.DataFrame(),pd.DataFrame(),pd.DataFrame()
        job_B, nav_B, df_cub_proc = pd.DataFrame(),pd.DataFrame(),pd.DataFrame()

        # --- Project BOM ---
        if not miss_A:
            st.subheader("📦 Project BOM")
            df_bom = pipeline_3A_0_rename(files["bom"], df_code, extras, debug=debug_flag)
            df_bom = pipeline_3A_1_filter(df_bom, df_stock)
            df_bom = pipeline_3A_2_accessories(df_bom, df_acc)
            df_bom = pipeline_3A_3_nav(df_bom, df_part_no)
            df_bom = pipeline_3A_4_stock(df_bom, files["ks"])
            job_A, nav_A, df_bom_proc = pipeline_3A_5_tables(df_bom, inputs["project_number"], df_part_no)

        # --- CUBIC BOM ---
        if not inputs["rittal"] and not miss_B:
            st.subheader("📦 CUBIC BOM")
            df_cubic = pipeline_3B_0_prepare_cubic(files["cubic_bom"], df_code, extras, debug=debug_flag)
            df_j, df_n = pipeline_3B_1_filtering(df_cubic, df_stock)
            df_j = pipeline_3B_2_accessories(df_j, df_acc)
            df_n = pipeline_3B_2_accessories(df_n, df_acc)
            df_j = pipeline_3B_3_nav(df_j, df_part_no)
            df_n = pipeline_3B_3_nav(df_n, df_part_no)
            df_j = pipeline_3B_4_stock(df_j, files["ks"])
            job_B, nav_B, df_cub_proc = pipeline_3B_5_tables(df_j, df_n, inputs["project_number"], df_part_no)

        # --- Calculation ---
        calc = pipeline_4_1_calculation(df_bom_proc, df_cub_proc, df_hours,
                                        inputs["panel_type"], inputs["grounding"], inputs["project_number"])
        miss_nav_A = pipeline_4_2_missing_nav(df_bom_proc,"Project BOM")
        miss_nav_B = pipeline_4_2_missing_nav(df_cub_proc,"CUBIC BOM")

        st.success("✅ Processing complete!")

        if not job_A.empty:
            st.subheader("📑 Job Journal (Project BOM)")
            st.dataframe(job_A,use_container_width=True)
        if not job_B.empty:
            st.subheader("📑 Job Journal (CUBIC BOM)")
            st.dataframe(job_B,use_container_width=True)

            # --- Naujas mechanikos pasirinkimas ---
            st.subheader("⚙️ Allocate to Mechanics")
            editable = job_B.copy()
            editable["Take Qty"] = 0
            edited = st.data_editor(
                editable,
                num_rows="dynamic",
                use_container_width=True,
                key="mech_editor"
            )

            if st.button("✅ Confirm Mechanics Allocation"):
                mech_rows = []
                for _, r in edited.iterrows():
                    take = float(r.get("Take Qty", 0) or 0)
                    if take > 0:
                        mech_rows.append({
                            "Type": r.get("Type", "Item"),
                            "No.": r.get("No."),
                            "Document No.": r.get("Document No."),
                            "Job No.": r.get("Job No."),
                            "Job Task No.": r.get("Job Task No."),
                            "Quantity": take,
                            "Location Code": r.get("Location Code", ""),
                            "Bin Code": r.get("Bin Code", ""),
                            "Description": r.get("Description", ""),
                            "Original Type": r.get("Original Type", "")
                        })
                df_mech = pd.DataFrame(mech_rows)
                if not df_mech.empty:
                    st.subheader("📑 Job Journal (CUBIC BOM TO MECH.)")
                    st.dataframe(df_mech, use_container_width=True)

        if not nav_A.empty:
            st.subheader("🛒 NAV Table (Project BOM)")
            st.dataframe(nav_A,use_container_width=True)
        if not nav_B.empty:
            st.subheader("🛒 NAV Table (CUBIC BOM)")
            st.dataframe(nav_B,use_container_width=True)

        st.subheader("💰 Calculation")
        st.dataframe(calc,use_container_width=True)

        if not miss_nav_A.empty or not miss_nav_B.empty:
            st.subheader("⚠️ Missing NAV Numbers")
            if not miss_nav_A.empty: st.dataframe(miss_nav_A,use_container_width=True)
            if not miss_nav_B.empty: st.dataframe(miss_nav_B,use_container_width=True)

