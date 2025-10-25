import streamlit as st, pandas as pd, re, io, datetime, os, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
CURRENCY="EUR"; CURRENCY_FORMAT='#,##0.00 "EUR"'; PURCHASE_LOCATION_CODE="KAUNAS"; ALLOC_LOCATION_CODE="KAUNAS"
# ========== Helpers ==========
def get_app_version():
    try:
        cnt=subprocess.check_output(["git","rev-list","--count","HEAD"],stderr=subprocess.DEVNULL).decode().strip()
        sha=subprocess.check_output(["git","rev-parse","--short","HEAD"],stderr=subprocess.DEVNULL).decode().strip()
        return f"v{int(cnt):03d} ({sha})"
    except Exception:
        env=os.getenv("APP_VERSION") or os.getenv("COMMIT_SHA"); return env if env else "v000"
def safe_parse_qty(x):
    if pd.isna(x): return 0.0
    if isinstance(x,(int,float)): return float(x)
    s=str(x).strip()
    if s in {"-","–","—",""}: return 0.0
    s=s.replace("\\xa0","").replace(" ","")
    if "," in s and "." in s: s=s.replace(",","")
    else: s=s.replace(".","").replace(",",".")
    try: return float(s)
    except Exception: return 0.0
def read_excel_any(file,**kwargs):
    try: return pd.read_excel(file,engine="openpyxl",**kwargs)
    except Exception: return pd.read_excel(file,engine="xlrd",**kwargs)
def normalize_no(x):
    try: return str(int(float(str(x).replace(",","." ).strip())))
    except Exception: return str(x).strip()
def ensure_scalar_strings(df):
    import numpy as np
    if df is None or df.empty: return df
    def _one(v):
        if isinstance(v,pd.Series):
            v=v.dropna(); return "" if v.empty else _one(v.iloc[0])
        if isinstance(v,(list,tuple,set,np.ndarray,dict)): return str(v)
        return v
    return df.applymap(_one)
def allocate_from_stock(no,need,stock_rows):
    out=[]; remaining=float(pd.to_numeric(pd.Series([need]),errors="coerce").fillna(0).iloc[0])
    if isinstance(stock_rows,pd.DataFrame) and not stock_rows.empty:
        for _,srow in stock_rows.iterrows():
            if remaining<=0: break
            bin_code=str(srow.get("Bin Code","")).strip(); qty=float(pd.to_numeric(pd.Series([srow.get("Quantity",0)]),errors="coerce").fillna(0).iloc[0])
            if qty<=0 or bin_code=="67-01-01-01": continue
            take=min(qty,remaining)
            if take>0: out.append({"No.":no,"Bin Code":bin_code,"Allocated Qty":take}); remaining-=take
    if remaining>0: out.append({"No.":no,"Bin Code":"","Allocated Qty":remaining})
    return out
def add_extra_components(df,extras):
    if df is None: df=pd.DataFrame()
    out=df.copy()
    for e in extras or []:
        out=pd.concat([out,pd.DataFrame([{"Original Type":e.get("type",""),"Quantity":e.get("qty",1),"Source":"Extra","No.":e.get("force_no",e.get("type",""))}])],ignore_index=True)
    return out
# ========== Inputs / Uploads ==========
def pipeline_2_1_user_inputs():
    st.subheader("Project Information")
    pn=st.text_input("Project number (1234-567)")
    if pn and not re.match(r"^\\d{4}-\\d{3}$",pn): st.error("Invalid format (must be 1234-567)"); return None
    types=["A","B","B1","B2","C","C1","C2","C3","C4","C4.1","C5","C6","C7","C8","F","F1","F2","F3","F4","F4.1","F5","F6","F7","G","G1","G2","G3","G4","G5","G6","G7","Custom"]; switches=["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111","31109","31107","C404400S","C634630S"]
    return {"project_number":pn,"panel_type":st.selectbox("Panel type",types),"grounding":st.selectbox("Grounding type",["TT","TN-S","TN-C-S"]),"main_switch":st.selectbox("Main switch",switches),"swing_frame":st.checkbox("Swing frame?"),"ups":st.checkbox("UPS?"),"rittal":st.checkbox("Rittal?")}
def pipeline_2_2_file_uploads(rittal=False):
    st.subheader("Upload Required Files"); dfs={}
    if not rittal:
        cubic_bom=st.file_uploader("Insert CUBIC BOM",type=["xls","xlsx","xlsm"],key="cubic_bom")
        if cubic_bom:
            try: df=read_excel_any(cubic_bom,skiprows=15)
            except Exception: df=read_excel_any(cubic_bom)
            df=df.rename(columns=lambda c:str(c).strip())
            qty_cols=[c for c in df.columns if str(c).strip() in {"E","F","G"}]
            combo=[c for c in df.columns if re.sub(r"\\s+","",str(c)).upper() in {"E+F+G","E+F","F+G","E+G"} or (("E" in str(c).upper()) and ("F" in str(c).upper()) and ("G" in str(c).upper()))]
            if qty_cols: df["Quantity"]=df[qty_cols].bfill(axis=1).iloc[:,0]
            elif combo:
                cc=combo[0]; df["Quantity"]=df[cc].apply(lambda v: safe_parse_qty(re.search(r"([0-9]+[.,]?[0-9]*)",str(v)).group(1)) if (pd.notna(v) and re.search(r"([0-9]+[.,]?[0-9]*)",str(v))) else 0.0)
            else:
                if "Quantity" not in df.columns: df["Quantity"]=0
            df["Quantity"]=pd.to_numeric(df["Quantity"],errors="coerce").fillna(0)
            if "Item Id" in df.columns: df=df.rename(columns={"Item Id":"Original Type"})
            else: df["Original Type"]=df[df.columns[0]].astype(str)
            if "No." not in df.columns: df["No."]=df["Original Type"]
            dfs["cubic_bom"]=df
    bom=st.file_uploader("Insert BOM",type=["xls","xlsx","xlsm"],key="bom")
    if bom:
        df=read_excel_any(bom)
        if df.shape[1]>=2:
            colA=df.iloc[:,0].fillna("").astype(str).str.strip(); colB=df.iloc[:,1].fillna("").astype(str).str.strip()
            df["Original Article"]=colA; df["Original Type"]=colB.where(colB!="",colA)
        else:
            df["Original Article"]=df.iloc[:,0].fillna("").astype(str).str.strip(); df["Original Type"]=df["Original Article"]
        dfs["bom"]=df
    data_file=st.file_uploader("Insert DATA",type=["xls","xlsx","xlsm"],key="data")
    if data_file: dfs["data"]=pd.read_excel(data_file,sheet_name=None)
    ks=st.file_uploader("Insert Kaunas Stock",type=["xls","xlsx","xlsm"],key="ks")
    if ks: dfs["ks"]=read_excel_any(ks)
    return dfs
def pipeline_2_3_get_sheet_safe(data_dict,names):
    if not isinstance(data_dict,dict): return None
    targets=[n.upper().replace(" ","_") for n in names]
    for key in data_dict.keys():
        if str(key).strip().upper().replace(" ","_") in targets: return data_dict[key]
    return None
def pipeline_2_4_normalize_part_no(df_raw):
    if df_raw is None or df_raw.empty: return pd.DataFrame()
    df=df_raw.copy().rename(columns=lambda c:str(c).strip()); m={}
    if df.shape[1]>=1: m[df.columns[0]]="PartNo_A"
    if df.shape[1]>=2: m[df.columns[1]]="PartName_B"
    if df.shape[1]>=3: m[df.columns[2]]="Desc_C"
    if df.shape[1]>=4: m[df.columns[3]]="Manufacturer_D"
    if df.shape[1]>=5: m[df.columns[4]]="SupplierNo_E"
    if df.shape[1]>=6: m[df.columns[5]]="UnitPrice_F"
    return df.rename(columns=m)
# ========== Project BOM (3A) ==========
def pipeline_3A_0_rename(df_bom,df_part_code,extras=None):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    df=df_bom.copy()
    if df_part_code is not None and not df_part_code.empty:
        rename_map=dict(zip(df_part_code.iloc[:,0].astype(str).str.strip(),df_part_code.iloc[:,1].astype(str).str.strip()))
        if "Original Type" in df.columns: df["Original Type"]=df["Original Type"].astype(str).str.strip().replace(rename_map)
    if "Original Article" not in df.columns: df["Original Article"]=df.iloc[:,0].astype(str)
    if extras: df=add_extra_components(df,[e for e in extras if e.get("target")=="bom"])
    return df
def pipeline_3A_1_filter(df_bom,df_stock):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    if df_stock is None or df_stock.empty: return df_bom.copy()
    cols=list(df_stock.columns)
    if len(cols)<3: return df_bom.copy()
    s=df_stock.rename(columns={cols[0]:"Component",cols[2]:"Comment"})
    excluded=s[s["Comment"].astype(str).str.lower().str.strip()=="no need"]["Component"].astype(str).str.upper().str.replace(" ","").str.strip().unique()
    df=df_bom.copy(); df["Norm_Type"]=df["Original Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    return df[~df["Norm_Type"].isin(excluded)].drop(columns=["Norm_Type"]).reset_index(drop=True)
def pipeline_3A_2_accessories(df_bom,df_acc):
    if df_acc is None or df_acc.empty: return df_bom
    out=df_bom.copy()
    for _,row in df_bom.iterrows():
        main=str(row["Original Type"]).strip(); m=df_acc[df_acc.iloc[:,0].astype(str).str.strip()==main]
        for _,acc_row in m.iterrows():
            vals=acc_row.values[1:]
            for i in range(0,len(vals),3):
                if i+2>=len(vals) or pd.isna(vals[i]): break
                item=str(vals[i]).strip(); q=safe_parse_qty(str(vals[i+1]).strip()); manuf=str(vals[i+2]).strip()
                out=pd.concat([out,pd.DataFrame([{"Original Type":item,"Quantity":q,"Manufacturer":manuf,"Source":"Accessory"}])],ignore_index=True)
    return out
def pipeline_3A_3_nav(df_bom,df_part_no):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    if df_part_no is None or df_part_no.empty: df=df_bom.copy(); df["No."]=""; return df
    part=df_part_no.copy().reset_index(drop=True).rename(columns=lambda c:str(c).strip())
    if "PartName_B" not in part.columns or "PartNo_A" not in part.columns: df=df_bom.copy(); df["No."]=""; return df
    part["Norm_B"]=part["PartName_B"].astype(str).str.upper().str.replace(" ","").str.strip()
    def nn(x):
        try: return str(int(float(str(x).strip().replace(",","."))))
        except Exception: return str(x).strip()
    part["PartNo_A"]=part["PartNo_A"].map(nn).fillna("").astype(str)
    part=part.drop_duplicates(subset=["Norm_B"],keep="first").drop_duplicates(subset=["PartNo_A"],keep="first")
    df=df_bom.copy(); df["Norm_Type"]=df["Original Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    df["No."]=df["Norm_Type"].map(dict(zip(part["Norm_B"],part["PartNo_A"]))).fillna("").astype(str)
    m=[c for c in ["PartNo_A","Desc_C","Manufacturer_D","SupplierNo_E","UnitPrice_F","Norm_B"] if c in part.columns]
    if m:
        df=df.merge(part[m],left_on="No.",right_on="PartNo_A",how="left").rename(columns={"Desc_C":"Description","Manufacturer_D":"Supplier","SupplierNo_E":"Supplier No.","UnitPrice_F":"Unit Cost"}).drop(columns=[c for c in ["Norm_Type","Norm_B","PartNo_A"] if c in df.columns],errors="ignore")
    else: df=df.drop(columns=["Norm_Type"],errors="ignore")
    return ensure_scalar_strings(df)
def _read_stock_df(ks_file):
    if isinstance(ks_file,pd.DataFrame): stock=ks_file.copy()
    else: stock=pd.read_excel(io.BytesIO(ks_file.getvalue()),engine="openpyxl")
    stock=stock.rename(columns=lambda c:str(c).strip())
    cand_no=[c for c in stock.columns if c.lower() in ["no.","no","item no.","item no"]]
    cand_bin=[c for c in stock.columns if c.lower() in ["bin code","bin","bin_code"]]
    cand_qty=[c for c in stock.columns if c.lower() in ["quantity","qty","q"]]
    if cand_no and cand_bin and cand_qty: cols=[cand_no[0],cand_bin[0],cand_qty[0]]; stock=stock[cols]; stock.columns=["No.","Bin Code","Quantity"]
    else:
        cols=list(stock.columns)
        if len(cols)>=4: stock=stock[[cols[2],cols[1],cols[3]]]; stock.columns=["No.","Bin Code","Quantity"]
        else: return pd.DataFrame(columns=["No.","Bin Code","Quantity"])
    stock["No."]=stock["No."].apply(normalize_no); stock["Quantity"]=pd.to_numeric(stock["Quantity"],errors="coerce").fillna(0.0); stock["Bin Code"]=stock["Bin Code"].astype(str).str.strip(); return stock
def pipeline_3A_4_stock(df_bom,ks_file):
    if df_bom is None or df_bom.empty: return pd.DataFrame()
    stock=_read_stock_df(ks_file); df=df_bom.copy(); df["No."]=df["No."].apply(normalize_no); groups={k:v for k,v in stock.groupby("No.")}; df["Stock Rows"]=df["No."].map(groups); return df
def pipeline_3A_5_tables(df_bom,project_number,df_part_no):
    rows=[]
    for _,row in df_bom.iterrows():
        no=row.get("No."); qty=safe_parse_qty(row.get("Quantity",0)); stock_rows=row.get("Stock Rows")
        if not isinstance(stock_rows,pd.DataFrame) or stock_rows.empty:
            rows.append({"Entry Type":"Item","No.":no,"Document No.":f"{project_number}/N","Job No.":project_number,"Job Task No.":1144,"Quantity":qty,"Location Code":PURCHASE_LOCATION_CODE,"Bin Code":"","Description":row.get("Description",""),"Original Type":row.get("Original Type","")}); continue
        for alloc in allocate_from_stock(no,qty,stock_rows):
            rows.append({"Entry Type":"Item","No.":no,"Document No.":project_number,"Job No.":project_number,"Job Task No.":1144,"Quantity":alloc["Allocated Qty"],"Location Code":ALLOC_LOCATION_CODE if alloc["Bin Code"] else PURCHASE_LOCATION_CODE,"Bin Code":alloc["Bin Code"],"Description":row.get("Description",""),"Original Type":row.get("Original Type","")})
    job=pd.DataFrame(rows); supplier_map=manuf_map={}
    if df_part_no is not None and not df_part_no.empty:
        if {"PartNo_A","SupplierNo_E"}.issubset(df_part_no.columns): supplier_map=dict(zip(df_part_no["PartNo_A"].astype(str),df_part_no["SupplierNo_E"]))
        if {"PartNo_A","Manufacturer_D"}.issubset(df_part_no.columns): manuf_map=dict(zip(df_part_no["PartNo_A"].astype(str),df_part_no["Manufacturer_D"].astype(str)))
    tmp=df_bom.copy(); 
    if "Quantity" not in tmp: tmp["Quantity"]=0
    if "Description" not in tmp: tmp["Description"]=""
    tmp["No."]=tmp["No."].astype(str); tmp["Quantity"]=pd.to_numeric(tmp["Quantity"],errors="coerce").fillna(0); rows2=[]
    for _,r in tmp.iterrows():
        part_no=str(r["No."]); qty=float(r.get("Quantity",0) or 0); manuf=(manuf_map or {}).get(part_no,""); profit=10 if "DANFOSS" in str(manuf).upper() else 17; supplier=(supplier_map or {}).get(part_no,30093)
        rows2.append({"Entry Type":"Item","No.":part_no,"Quantity":qty,"Supplier":supplier,"Profit":profit,"Discount":0,"Description":r.get("Description","")})
    nav=pd.DataFrame(rows2,columns=["Entry Type","No.","Quantity","Supplier","Profit","Discount","Description"]); return job,nav,df_bom
# ========== CUBIC BOM (3B) ==========
def pipeline_3B_0_prepare_cubic(df_cubic,df_part_code,extras=None):
    if df_cubic is None or df_cubic.empty: return pd.DataFrame()
    df=df_cubic.copy().rename(columns=lambda c:str(c).strip())
    qty_cols=[c for c in df.columns if str(c).strip() in {"E","F","G"}]
    combo=[c for c in df.columns if re.sub(r"\\s+","",str(c)).upper() in {"E+F+G","E+F","F+G","E+G"} or (("E" in str(c).upper()) and ("F" in str(c).upper()) and ("G" in str(c).upper()))]
    if qty_cols: df["Quantity"]=df[qty_cols].bfill(axis=1).iloc[:,0]
    elif combo:
        cc=combo[0]; df["Quantity"]=df[cc].apply(lambda v: safe_parse_qty(re.search(r"([0-9]+[.,]?[0-9]*)",str(v)).group(1)) if (pd.notna(v) and re.search(r"([0-9]+[.,]?[0-9]*)",str(v))) else 0.0)
    else:
        if "Quantity" not in df.columns: df["Quantity"]=0
    df["Quantity"]=pd.to_numeric(df["Quantity"],errors="coerce").fillna(0)
    if "Item Id" in df.columns: df["Original Type"]=df["Item Id"].astype(str).str.strip()
    elif "Original Type" not in df.columns: df["Original Type"]=df[df.columns[0]].astype(str)
    if "No." not in df.columns: df["No."]=df["Original Type"]
    if df_part_code is not None and not df_part_code.empty:
        rename_map=dict(zip(df_part_code.iloc[:,0].astype(str).str.strip(),df_part_code.iloc[:,1].astype(str).str.strip()))
        df["Original Type"]=df["Original Type"].astype(str).str.strip().replace(rename_map)
    if extras: df=add_extra_components(df,[e for e in extras if e.get("target")=="cubic"])
    return df
def pipeline_3B_1_filtering(df_cubic,df_stock):
    if df_cubic is None or df_cubic.empty: return pd.DataFrame(),pd.DataFrame()
    if df_stock is None or df_stock.empty: return df_cubic.copy(),df_cubic.copy()
    cols=list(df_stock.columns)
    if len(cols)<3: return df_cubic.copy(),df_cubic.copy()
    s=df_stock.rename(columns={cols[0]:"Component",cols[2]:"Comment"})
    bad=s[s["Comment"].astype(str).str.strip().str.lower().isin(["no need","q1"])]["Component"].astype(str)
    banned=bad.str.upper().str.replace(" ","").str.strip().unique()
    df=df_cubic.copy(); df["Norm_Type"]=df["Original Type"].astype(str).str.upper().str.replace(" ","").str.strip()
    keep=df[~df["Norm_Type"].isin(banned)].reset_index(drop=True).drop(columns=["Norm_Type"])
    return keep.copy(),keep.copy()
def pipeline_3B_2_accessories(df,df_acc):
    if df_acc is None or df_acc.empty: return df
    out=df.copy()
    for _,row in df.iterrows():
        main=str(row["Original Type"]).strip(); m=df_acc[df_acc.iloc[:,0].astype(str).str.strip()==main]
        for _,acc_row in m.iterrows():
            v=acc_row.values[1:]
            for i in range(0,len(v),3):
                if i+2>=len(v) or pd.isna(v[i]): break
                item=str(v[i]).strip(); qty=safe_parse_qty(str(v[i+1]).strip()); manuf=str(v[i+2]).strip()
                out=pd.concat([out,pd.DataFrame([{"Original Type":item,"Quantity":qty,"Manufacturer":manuf,"Source":"Accessory"}])],ignore_index=True)
    return out
def pipeline_3B_3_nav(df,df_part_no): return pipeline_3A_3_nav(df,df_part_no)
def pipeline_3B_4_stock(df_journal,ks_file): return pipeline_3A_4_stock(df_journal,ks_file)
def pipeline_3B_5_tables(df_journal,df_nav,project_number,df_part_no):
    rows=[]
    for _,row in df_journal.iterrows():
        no=row.get("No."); qty=safe_parse_qty(row.get("Quantity",0)); stock_rows=row.get("Stock Rows")
        if not isinstance(stock_rows,pd.DataFrame) or stock_rows.empty:
            rows.append({"Entry Type":"Item","No.":no,"Document No.":f"{project_number}/N","Job No.":project_number,"Job Task No.":1144,"Quantity":qty,"Location Code":PURCHASE_LOCATION_CODE,"Bin Code":"","Description":row.get("Description",""),"Original Type":row.get("Original Type","")}); continue
        for alloc in allocate_from_stock(no,qty,stock_rows):
            rows.append({"Entry Type":"Item","No.":no,"Document No.":project_number,"Job No.":project_number,"Job Task No.":1144,"Quantity":alloc["Allocated Qty"],"Location Code":ALLOC_LOCATION_CODE if alloc["Bin Code"] else PURCHASE_LOCATION_CODE,"Bin Code":alloc["Bin Code"],"Description":row.get("Description",""),"Original Type":row.get("Original Type","")})
    job=pd.DataFrame(rows); _,nav,_=pipeline_3A_5_tables(df_nav,project_number,df_part_no); return job,nav,df_nav
# ========== Calculation & Missing ==========
def pipeline_4_1_calculation(df_bom,df_cubic,df_hours,panel_type,grounding,project_number,df_instr=None):
    if df_bom is None: df_bom=pd.DataFrame()
    if df_cubic is None: df_cubic=pd.DataFrame()
    if df_hours is None: df_hours=pd.DataFrame()
    if not df_bom.empty and {"Quantity","Unit Cost"}.issubset(df_bom.columns): parts=(pd.to_numeric(df_bom["Quantity"],errors="coerce").fillna(0)*pd.to_numeric(df_bom["Unit Cost"],errors="coerce").fillna(0)).sum()
    else: parts=0
    if not df_cubic.empty and {"Quantity","Unit Cost"}.issubset(df_cubic.columns): cubic=(pd.to_numeric(df_cubic["Quantity"],errors="coerce").fillna(0)*pd.to_numeric(df_cubic["Unit Cost"],errors="coerce").fillna(0)).sum()
    else: cubic=0
    hours=0
    if not df_hours.empty and df_hours.shape[1]>4:
        rate=pd.to_numeric(df_hours.iloc[1,4],errors="coerce"); row=df_hours[df_hours.iloc[:,0].astype(str).str.upper()==str(panel_type).upper()]
        if not row.empty:
            if grounding=="TT": h=pd.to_numeric(row.iloc[0,1],errors="coerce")
            elif grounding=="TN-S": h=pd.to_numeric(row.iloc[0,2],errors="coerce")
            else: h=pd.to_numeric(row.iloc[0,3],errors="coerce")
            hours=(h if pd.notna(h) else 0)*(rate if pd.notna(rate) else 0)
    smart=9750.0; wire=2500.0; total=parts+cubic+hours+smart+wire
    proj=""; pallet=""
    if df_instr is not None and not df_instr.empty:
        r=df_instr[df_instr.iloc[:,0].astype(str).str.upper()==str(panel_type).upper()]
        if not r.empty: proj=str(r.iloc[0,1]) if r.shape[1]>1 else ""; pallet=str(r.iloc[0,2]) if r.shape[1]>2 else ""
    return pd.DataFrame([{"Label":"Parts","Value":parts},{"Label":"Cubic","Value":cubic},{"Label":"Hours cost","Value":hours},{"Label":"Smart supply","Value":smart},{"Label":"Wire set","Value":wire},{"Label":"Extra","Value":0},{"Label":"Total","Value":total},{"Label":"Total+5%","Value":total*1.05},{"Label":"Total+35%","Value":total*1.35},{"Label":"Project size","Value":proj},{"Label":"Pallet size","Value":pallet}])
def pipeline_4_2_missing_nav(df,source):
    if df is None or df.empty or "No." not in df.columns: return pd.DataFrame()
    m=df[df["No."].astype(str).str.strip()=="" ] if not df.empty else pd.DataFrame()
    if m.empty: return pd.DataFrame()
    qty=pd.to_numeric(m.get("Quantity",0),errors="coerce").fillna(0).astype(float) if "Quantity" in m else 0
    return pd.DataFrame({"Source":source,"Original Article":m.get("Original Article",""),"Original Type":m.get("Original Type",""),"Quantity":qty,"NAV No.":m["No."]})
# ========== Render (optimized +/-) ==========
BTN_CSS="""<style>
.mech *{color:#fff!important;font-family:system-ui,Segoe UI,Arial,sans-serif!important}
.mech .row{border-bottom:1px solid rgba(255,255,255,.25);padding:6px 0;margin:2px 0}
.mech .label{margin:0;line-height:1.2;font-weight:600}
.mech .qty{display:flex;align-items:center;gap:10px;justify-content:center}
.mech .pill{min-width:72px;text-align:center;font-weight:800;font-size:22px;padding:4px 12px;border:1px solid rgba(255,255,255,.35);border-radius:10px}
.mech .stButton>button{background:#0b6b39!important;color:#fff!important;font-weight:800!important;border-radius:10px!important;padding:6px 0!important}
</style>"""
def _norm_type(s): return str(s).upper().replace(" ","").strip()
def _norm_no(x):
    try: return str(int(float(str(x).replace(",","." ).strip())))
    except Exception: return str(x).strip()
def _excluded_sets(df_stock):
    if df_stock is None or df_stock.empty or df_stock.shape[1]<3: return set(),set()
    cols=list(df_stock.columns); s=df_stock.rename(columns={cols[0]:"Component",cols[2]:"Comment"})
    m=s["Comment"].astype(str).str.strip().str.lower().isin(["no need","q1"]); comp=s.loc[m,"Component"].astype(str)
    return set(comp.map(_norm_type)),set(comp.map(_norm_no))
def _apply_excl(df,ex_t,ex_n):
    if df is None or df.empty: return df
    t=df.copy(); t["_T"]=t["Original Type"].map(_norm_type); t["_N"]=t["No."].map(_norm_no) if "No." in t.columns else ""
    t=t[~t["_T"].isin(ex_t) & ~t["_N"].isin(ex_n)].drop(columns=["_T","_N"],errors="ignore"); return t
def _process_all(files,inputs):
    book=files.get("data",{})
    df_stock=pipeline_2_3_get_sheet_safe(book,["Stock"])
    df_part=pipeline_2_4_normalize_part_no(pipeline_2_3_get_sheet_safe(book,["Part_no","Parts_no","Part no"]))
    df_hours=pipeline_2_3_get_sheet_safe(book,["Hours"]); df_acc=pipeline_2_3_get_sheet_safe(book,["Accessories"]); df_code=pipeline_2_3_get_sheet_safe(book,["Part_code"]); df_instr=pipeline_2_3_get_sheet_safe(book,["Instructions"])
    extras=[]
    if inputs["ups"]: extras.extend([{"type":"LI32111CT01","qty":1,"target":"bom","force_no":"2214036"},{"type":"ADV UPS holder V3","qty":1,"target":"bom","force_no":"2214035"},{"type":"268-2610","qty":1,"target":"bom","force_no":"1865206"}])
    if inputs["swing_frame"]: extras.append({"type":"9030+2970","qty":1,"target":"cubic","force_no":"2185835"})
    if df_instr is not None and not df_instr.empty:
        row=df_instr[df_instr.iloc[:,0].astype(str).str.upper()==str(inputs["panel_type"]).upper()]
        if not row.empty:
            if inputs["panel_type"][0] not in ["F","G"]:
                try: q=int(pd.to_numeric(row.iloc[0,4],errors="coerce").fillna(0))
                except Exception: q=0
                if q>0: extras.append({"type":"SDD07550","qty":q,"target":"cubic","force_no":"SDD07550"})
            for cidx in range(5,10):
                if cidx<row.shape[1]:
                    v=str(row.iloc[0,cidx]).strip()
                    if v and v.lower()!="nan": extras.append({"type":v,"qty":1,"target":"cubic"})
    job_A=nav_A=df_bom_proc=pd.DataFrame()
    if {"bom","ks"}.issubset(files.keys()):
        df_bom=pipeline_3A_0_rename(files["bom"],df_code,extras); df_bom=pipeline_3A_1_filter(df_bom,df_stock); df_bom=pipeline_3A_2_accessories(df_bom,df_acc); df_bom=pipeline_3A_3_nav(df_bom,df_part); df_bom=pipeline_3A_4_stock(df_bom,files["ks"]); job_A,nav_A,df_bom_proc=pipeline_3A_5_tables(df_bom,inputs["project_number"],df_part)
    job_B=nav_B=df_cub_proc=pd.DataFrame()
    if not inputs["rittal"] and {"cubic_bom","ks"}.issubset(files.keys()):
        df_cubic=pipeline_3B_0_prepare_cubic(files["cubic_bom"],df_code,extras); df_j,df_n=pipeline_3B_1_filtering(df_cubic,df_stock); df_j=pipeline_3B_2_accessories(df_j,df_acc); df_n=pipeline_3B_2_accessories(df_n,df_acc); df_j=pipeline_3B_3_nav(df_j,df_part); df_n=pipeline_3B_3_nav(df_n,df_part); df_j=pipeline_3B_4_stock(df_j,files["ks"]); job_B,nav_B,df_cub_proc=pipeline_3B_5_tables(df_j,df_n,inputs["project_number"],df_part)
    ex_t,ex_n=_excluded_sets(df_stock)
    return {"hours":df_hours,"instr":df_instr,"job_A":job_A,"nav_A":nav_A,"bom_proc":df_bom_proc,"job_B":job_B,"nav_B":nav_B,"cub_proc":df_cub_proc,"ex_t":ex_t,"ex_n":ex_n}
def render():
    st.header(f"BOM Management · {get_app_version()}")
    inputs=pipeline_2_1_user_inputs()
    if not inputs: return
    st.session_state["inputs"]=inputs
    files=pipeline_2_2_file_uploads(inputs["rittal"])
    if not files: return
    reqA=["bom","data","ks"]; reqB=["cubic_bom","data","ks"] if not inputs["rittal"] else []
    missA=[k for k in reqA if k not in files]; missB=[k for k in reqB if k not in files]
    st.subheader("📋 Required files"); c1,c2=st.columns(2)
    with c1: st.success("Project BOM: OK") if not missA else st.warning(f"Project BOM missing: {missA}")
    with c2: st.info("CUBIC BOM skipped (Rittal)") if inputs["rittal"] else (st.success("CUBIC BOM: OK") if not missB else st.warning(f"CUBIC BOM missing: {missB}"))
    if st.button("🚀 Run Processing"):
        st.session_state["processing_started"]=True; st.session_state["bundle"]=_process_all(files,inputs); st.session_state["mech_confirmed"]=False; st.session_state["mech_take"]={}; st.session_state["df_mech"]=pd.DataFrame(); st.session_state["df_remain"]=pd.DataFrame()
    if not st.session_state.get("processing_started",False): st.stop()
    b=st.session_state.get("bundle",{}); job_B=b.get("job_B",pd.DataFrame())
    BTN_CSS="""<style>.mech *{color:#fff!important;font-family:system-ui,Segoe UI,Arial,sans-serif!important}.mech .row{border-bottom:1px solid rgba(255,255,255,.25);padding:6px 0;margin:2px 0}.mech .label{margin:0;line-height:1.2;font-weight:600}.mech .qty{display:flex;align-items:center;gap:10px;justify-content:center}.mech .pill{min-width:72px;text-align:center;font-weight:800;font-size:22px;padding:4px 12px;border:1px solid rgba(255,255,255,.35);border-radius:10px}.mech .stButton>button{background:#0b6b39!important;color:#fff!important;font-weight:800!important;border-radius:10px!important;padding:6px 0!important}</style>"""
    if not st.session_state.get("mech_confirmed",False) and not job_B.empty:
        st.subheader("📑 Job Journal (CUBIC BOM → allocate to Mechanics)"); st.markdown(BTN_CSS,unsafe_allow_html=True); st.session_state.setdefault("mech_take",{})
        editable=_apply_excl(job_B,b.get("ex_t",set()),b.get("ex_n",set())).copy(); editable["Available Qty"]=editable["Quantity"].astype(float)
        if editable.empty: st.info("No selectable items (filtered by Stock: No need/Q1)."); st.session_state["mech_confirmed"]=True; st.stop()
        head=st.columns([4,4,4,3]); head[0].markdown("**No.**"); head[1].markdown("**Original Type**"); head[2].markdown("**Description**"); head[3].markdown("**Allocate**")
        st.markdown("<div class='mech'>",unsafe_allow_html=True)
        for idx,row in editable.iterrows():
            cols=st.columns([4,4,4,3])
            with cols[0]: st.markdown(f"<div class='row'><p class='label'>{str(row.get('No.',''))}</p></div>",unsafe_allow_html=True)
            with cols[1]: st.markdown(f"<div class='row'><p class='label'>{str(row.get('Original Type',''))}</p></div>",unsafe_allow_html=True)
            with cols[2]: st.markdown(f"<div class='row'><p class='label'>{str(row.get('Description',''))}</p></div>",unsafe_allow_html=True)
            with cols[3]:
                key=f"take_{idx}"; mx=float(row["Available Qty"]); cur=float(st.session_state["mech_take"].get(key,0.0))
                a,bn,cp=st.columns([1,2,1])
                with a: st.button("−",key=f"minus_{idx}",disabled=(cur<=0),use_container_width=True,on_click=lambda k=key: st.session_state["mech_take"].update({k:max(st.session_state["mech_take"].get(k,0.0)-1,0.0)}))
                with bn: st.markdown(f"<div class='row qty'><div class='pill'>{cur:.0f}</div></div>",unsafe_allow_html=True)
                with cp: st.button("+",key=f"plus_{idx}",disabled=(cur>=mx),use_container_width=True,on_click=lambda k=key,m=mx: st.session_state["mech_take"].update({k:min(st.session_state["mech_take"].get(k,0.0)+1,m)}))
        st.markdown("</div>",unsafe_allow_html=True)
        if st.button("✅ Confirm Mechanics Allocation"):
            mech,remain=[],[]
            for idx,row in editable.iterrows():
                key=f"take_{idx}"; take=float(st.session_state["mech_take"].get(key,0.0)); avail=float(row["Available Qty"]); r=row.to_dict()
                if take>0: mech.append({**r,"Quantity":take})
                rem=max(avail-take,0.0)
                if rem>0 and str(r.get("No.",""))!="2185835": remain.append({**r,"Quantity":rem})
            st.session_state["df_mech"]=pd.DataFrame(mech); st.session_state["df_remain"]=pd.DataFrame(remain); st.session_state["mech_confirmed"]=True
            if inputs["swing_frame"]:
                swing=pd.DataFrame([{"Entry Type":"Item","Original Type":"9030+2970","No.":"2185835","Quantity":1,"Document No.":inputs["project_number"],"Job No.":inputs["project_number"],"Job Task No.":1144,"Location Code":PURCHASE_LOCATION_CODE,"Bin Code":"","Description":"Swing frame component","Source":"Extra"}])
                st.session_state["df_mech"]=pd.concat([st.session_state["df_mech"],swing],ignore_index=True)
        st.stop()
    def _show(df,title):
        if df is not None and not df.empty: st.subheader(title); st.data_editor(df,use_container_width=True,hide_index=True,height=300)
    _show(st.session_state.get("df_mech"),"📑 Job Journal (CUBIC BOM TO MECH.)"); _show(st.session_state.get("df_remain"),"📑 Job Journal (CUBIC BOM REMAINING)")
    _show(b.get("job_A"),"📑 Job Journal (Project BOM)"); _show(b.get("nav_A"),"🛒 NAV Table (Project BOM)"); _show(b.get("nav_B"),"🛒 NAV Table (CUBIC BOM)")
    calc=pipeline_4_1_calculation(b.get("bom_proc"),b.get("cub_proc"),b.get("hours"),inputs["panel_type"],inputs["grounding"],inputs["project_number"],b.get("instr")); _show(calc,"💰 Calculation")
    missA=pipeline_4_2_missing_nav(b.get("bom_proc"),"Project BOM"); missB=pipeline_4_2_missing_nav(b.get("cub_proc"),"CUBIC BOM"); _show(missA,"⚠️ Missing NAV Numbers (Project BOM)"); _show(missB,"⚠️ Missing NAV Numbers (CUBIC BOM)")
    st.subheader("💾 Export")
    if st.button("💾 Export Results to Excel"):
        ts=datetime.datetime.now().strftime("%Y%m%d%H%M")
        try: psize=str(calc[calc["Label"]=="Project size"]["Value"].iloc[0]); pl=str(calc[calc["Label"]=="Pallet size"]["Value"].iloc[0])
        except Exception: psize=pl=""
        fname=f"{inputs['project_number']}_{inputs['panel_type']}_{inputs['grounding']}_{pl}_{ts}.xlsx"
        wb=Workbook(); ws=wb.active; ws.title="Info"
        info=[["Project number",inputs["project_number"]],["Panel type",inputs["panel_type"]],["Grounding",inputs["grounding"]],["Main switch",inputs["main_switch"]],["Swing frame",inputs["swing_frame"]],["UPS",inputs["ups"]],["Rittal",inputs["rittal"]],["Project size",psize],["Pallet size",pl]]
        for r in info: ws.append(r)
        ws.column_dimensions["A"].width=20; ws.column_dimensions["B"].width=20
        bold=Font(bold=True); grey=PatternFill(start_color="DDDDDD",end_color="DDDDDD",fill_type="solid"); thin=Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
        for r in ws["A1":"A9"]:
            for c in r: c.font=bold; c.fill=grey; c.border=thin
        for r in ws["B1":"B9"]:
            for c in r: c.border=thin
        def add_df(df,title,colw=None,nav=False,calcSheet=False):
            if df is None or df.empty: return
            df=ensure_scalar_strings(df); w=wb.create_sheet(title); w.append(df.columns.tolist())
            for _,row in df.iterrows(): w.append(list(row.values))
            if colw:
                for col,wid in colw.items(): w.column_dimensions[col].width=wid
            mr,mc=w.max_row,w.max_column
            for rr in w.iter_rows(min_row=1,max_row=mr,min_col=1,max_col=mc):
                for cc in rr: cc.border=thin
            if nav:
                for rr in w["A1":"G1"]:
                    for cc in rr: cc.font=bold; cc.fill=grey
            if calcSheet:
                for rr in w["A1":"A10"]:
                    for cc in rr: cc.font=bold; cc.fill=grey
                for rr in w["B2":"B10"]:
                    for cc in rr: cc.number_format=CURRENCY_FORMAT
        job_w={"A":8,"B":10,"C":12,"D":12,"E":12,"F":12,"G":13,"H":12,"I":40,"J":25}
        add_df(st.session_state.get("df_mech"),"JobJournal_Mech",job_w); add_df(st.session_state.get("df_remain"),"JobJournal_Remaining",job_w); add_df(b.get("job_A"),"JobJournal_ProjectBOM",job_w); add_df(b.get("job_B"),"JobJournal_CUBICBOM",job_w)
        nav_w={"A":8,"B":10,"C":9,"D":9,"E":9,"F":9,"G":50}
        add_df(b.get("nav_B"),"NAV_CUBICBOM",nav_w,nav=True); add_df(b.get("nav_A"),"NAV_ProjectBOM",nav_w,nav=True); add_df(calc,"Calculation",{"A":12,"B":18},calcSheet=True); add_df(missA,"MissingNAV_ProjectBOM"); add_df(missB,"MissingNAV_CUBICBOM")
        p=f"/mnt/data/{fname}"; wb.save(p); st.download_button("⬇️ Download Excel",data=open(p,"rb"),file_name=fname,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if __name__=="__main__": render()