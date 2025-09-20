import streamlit as st
import pandas as pd
import io
import re

# =====================================================
# Pipeline 1.x – Helpers
# =====================================================

def pipeline_1_1_norm_name(x):
    """
    Normalize name: make uppercase, remove spaces.
    Pvz.: 'abc 123' → 'ABC123'
    """
    return ''.join(str(x).upper().split())

def pipeline_1_2_parse_qty(x):
    """
    Parse numeric quantities from string or mixed format.
    Tvarko kablelius, taškus, tarpelius.
    Pvz.: '1,5' → 1.5, '2.000,50' → 2000.5
    """
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
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
    """
    Format filename safe for Windows/SharePoint.
    Pašalina draudžiamus simbolius, tarpus pakeičia į '_'.
    """
    s = '' if s is None else str(s).strip()
    s = re.sub(r'[\\/:*?"<>|]+','',s)
    return s.replace(' ','_')

# =====================================================
# Pipeline 2.x – vartotojo įvestis ir failai
# =====================================================

def pipeline_2_1_user_inputs():
    """
    Surenka vartotojo įvestis: projekto numerį, panelės tipą,
    įžeminimo tipą, pagrindinį jungiklį ir pasirinktus checkbox’us.
    """
    st.subheader("🔢 Project Information")

    project_number = st.text_input("Project number (format: 1234-567)")
    if project_number and not re.match(r"^\d{4}-\d{3}$", project_number):
        st.error("⚠️ Invalid format (must be 1234-567)")
        return None

    panel_type = st.selectbox(
        "Panel type", 
        options=[
            'A','B','B1','B2','C','C1','C2','C3','C4','C4.1','C5','C6','C7','C8',
            'F','F1','F2','F3','F4','F4.1','F5','F6','F7',
            'G','G1','G2','G3','G4','G5','G6','G7',
            'Custom'
        ]
    )

    grounding   = st.selectbox("Grounding type", ["TT", "TN-S", "TN-C-S"])
    main_switch = st.selectbox("Main switch", ["C160S4FM","C125S4FM","C080S4FM","31115","31113","31111","31109","31107","C404400S","C634630S"])

    swing_frame = st.checkbox("Swing frame?")
    ups         = st.checkbox("UPS?")
    rittal      = st.checkbox("Rittal?")

    return {
        "project_number": project_number,
        "panel_type": panel_type,
        "grounding": grounding,
        "main_switch": main_switch,
        "swing_frame": swing_frame,
        "ups": ups,
        "rittal": rittal,
    }

def get_sheet_safe(data_dict, names):
    """
    Grąžina pirmą sutampantį lapą iš data_dict pagal galimus pavadinimus.
    names: sąrašas galimų variantų
    """
    for key in data_dict.keys():
        if str(key).strip().upper().replace(" ", "_") in [n.upper().replace(" ", "_") for n in names]:
            return data_dict[key]
    return None

# ---- Helper: universalus Excel reader (.xls + .xlsx) ----
def read_excel_any(file, **kwargs):
    try:
        return pd.read_excel(file, engine="openpyxl", **kwargs)
    except Exception:
        return pd.read_excel(file, engine="xlrd", **kwargs)

# Universal Excel reader (.xls / .xlsx / .xlsm)
def read_excel_any(file, **kwargs):
    try:
        return pd.read_excel(file, engine="openpyxl", **kwargs)
    except Exception:
        return pd.read_excel(file, engine="xlrd", **kwargs)

# ---- Pipeline 2.2: File uploads (be stulpelių validacijos) ----
def pipeline_2_2_file_uploads(rittal=False):
    st.subheader("📂 Upload Required Files")

    dfs = {}

    # --- CUBIC BOM (tik jei ne Rittal) ---
    if not rittal:
        st.markdown("<h3 style='color:#0ea5e9; font-weight:700;'>📂 Insert CUBIC BOM</h3>", unsafe_allow_html=True)
        cubic_bom = st.file_uploader("", type=["xls", "xlsx", "xlsm"], key="cubic_bom")
        if cubic_bom:
            try:
                dfs["cubic_bom"] = read_excel_any(cubic_bom, skiprows=13, usecols="B:F")
            except Exception as e:
                st.error(f"⚠️ Cannot open CUBIC BOM: {e}")

    # --- BOM ---
    st.markdown("<h3 style='color:#0ea5e9; font-weight:700;'>📂 Insert BOM</h3>", unsafe_allow_html=True)
    bom = st.file_uploader("", type=["xls", "xlsx", "xlsm"], key="bom")
    if bom:
        try:
            dfs["bom"] = read_excel_any(bom)
        except Exception as e:
            st.error(f"⚠️ Cannot open BOM: {e}")

    # --- DATA ---
    st.markdown("<h3 style='color:#0ea5e9; font-weight:700;'>📂 Insert DATA</h3>", unsafe_allow_html=True)
    data_file = st.file_uploader("", type=["xls", "xlsx", "xlsm"], key="data")
    if data_file:
        try:
            dfs["data"] = pd.read_excel(data_file, sheet_name=None)  # <-- VISI LAPAI
        except Exception as e:
            st.error(f"⚠️ Cannot open DATA: {e}")

    # --- Kaunas Stock ---
    st.markdown("<h3 style='color:#0ea5e9; font-weight:700;'>📂 Insert Kaunas Stock</h3>", unsafe_allow_html=True)
    ks_file = st.file_uploader("", type=["xls", "xlsx", "xlsm"], key="ks")
    if ks_file:
        try:
            dfs["ks"] = read_excel_any(ks_file)
        except Exception as e:
            st.error(f"⚠️ Cannot open Kaunas Stock: {e}")

    return dfs

# =====================================================
# Pipeline 3.x – Duomenų apdorojimas
# =====================================================

def pipeline_3_1_filtering(df_bom: pd.DataFrame, df_stock: pd.DataFrame) -> pd.DataFrame:
    """
    Pašalina iš BOM visus komponentus, kurie turi Comment reikšmę DATA.xlsx → Stock lape.
    Pagal nutylėjimą laikom: 
      1-asis stulpelis = Component,
      3-iasis stulpelis = Comment.
    """
    st.info("🚦 Filtering BOM according to DATA.xlsx Stock (Comment)...")

    # Pervadinam pirmą ir trečią stulpelį
    cols = list(df_stock.columns)
    if len(cols) >= 3:
        rename_map = {cols[0]: "Component", cols[2]: "Comment"}
        df_stock = df_stock.rename(columns=rename_map)
    else:
        st.error("❌ Stock sheet must have bent 3 stulpelius (Component, ..., Comment)")
        return df_bom

    # Atrenkam komponentus su komentarais
    excluded_components = df_stock[df_stock["Comment"].notna()]["Component"].dropna().astype(str)

    # Normalizuojam pavadinimus
    excluded_norm = excluded_components.str.upper().str.replace(" ", "").unique()
    df_bom = df_bom.copy()
    df_bom["Norm_Type"] = df_bom["Type"].astype(str).str.upper().str.replace(" ", "")

    # Filtravimas
    filtered = df_bom[~df_bom["Norm_Type"].isin(excluded_norm)].reset_index(drop=True)

    st.success(f"✅ BOM filtered: {len(df_bom)} → {len(filtered)} rows "
               f"(removed {len(df_bom) - len(filtered)} items with comments)")
    return filtered.drop(columns=["Norm_Type"])

def pipeline_3_2_add_accessories(df_bom: pd.DataFrame, df_accessories: pd.DataFrame) -> pd.DataFrame:
    """
    Prideda accessories pagal DATA.xlsx → Accessories lapą.
    Logika: jei BOM’e yra pagrindinis komponentas, įtraukiami jo accessories.
    """
    st.info("➕ Adding accessories...")

    if df_accessories is None or df_accessories.empty:
        st.warning("⚠️ Accessories sheet not found, skipping")
        return df_bom

    df_out = df_bom.copy()
    added = []

    for _, row in df_bom.iterrows():
        main_item = str(row["Type"]).strip()
        matches = df_accessories[df_accessories.iloc[:,0].astype(str).str.strip() == main_item]
        for _, acc_row in matches.iterrows():
            acc_values = acc_row.values[1:]
            for i in range(0, len(acc_values), 3):
                if i+2 >= len(acc_values) or pd.isna(acc_values[i]):
                    break
                acc_item = str(acc_values[i]).strip()
                try:
                    acc_qty = float(str(acc_values[i+1]).replace(",","."))
                except:
                    acc_qty = 1
                acc_manuf = str(acc_values[i+2]).strip()
                df_out = pd.concat([df_out, pd.DataFrame([{
                    "Type": "item",
                    "Cross-Reference No.": acc_item,
                    "Quantity": acc_qty,
                    "Manufacturer": acc_manuf
                }])], ignore_index=True)
                added.append(acc_item)

    st.success(f"✅ Added {len(added)} accessories")
    return df_out


def pipeline_3_3_add_nav_numbers(df_bom, df_part_no_raw):
    """
    Prideda NAV numerius į BOM pagal Part_no lapą iš DATA.xlsx.
    Jei nepavyko priskirti – No. paliekamas tuščias (None).
    Sukuriama papildoma lentelė su nerastais komponentais.
    """
    df_part_no = df_part_no_raw.copy()
    df_part_no.columns = [
        'PartNo_A',       # "Item no."
        'PartName_B',     # "Type no."
        'Desc_C',         # "Description"
        'Manufacturer_D', # "Supplier"
        'SupplierNo_E',   # "Supplier No."
        'UnitPrice_F'     # "Cost price / Unit cost DKK"
    ]

    # Normalizavimas
    df_part_no['Norm_B'] = df_part_no['PartName_B'].astype(str).str.upper().str.replace(" ", "")
    part_map = dict(zip(df_part_no['Norm_B'], df_part_no['PartNo_A']))

    # BOM papildymas
    df_bom = df_bom.copy()
    df_bom['Norm_Type'] = df_bom['Type'].astype(str).str.upper().str.replace(" ", "")
    df_bom['No.'] = df_bom['Norm_Type'].map(part_map)  # <- jei nerasta → None

    # Merge papildomos info
    df_bom = df_bom.merge(
        df_part_no[['PartNo_A', 'Desc_C', 'Manufacturer_D', 'SupplierNo_E', 'UnitPrice_F', 'Norm_B']],
        left_on='Norm_Type', right_on='Norm_B', how='left'
    )

    df_bom = df_bom.drop(columns=['Norm_Type', 'Norm_B'])
    df_bom = df_bom.rename(columns={
        'PartNo_A': 'NAV_No',
        'Desc_C': 'Description',
        'Manufacturer_D': 'Supplier',
        'SupplierNo_E': 'Supplier No.',
        'UnitPrice_F': 'Unit Cost'
    })

    # Lentelė su nerastais
    missing_df = df_bom[df_bom['No.'].isna()][['Type', 'Description']].copy()
    missing_df['Status'] = "NERASTA"

    st.session_state["part_no"] = df_part_no
    st.session_state["missing_parts"] = missing_df

    return df_bom


def pipeline_3_4_check_stock(df_bom, ks_file):
    df_out = df_bom.copy()

    # Pašalinam dublikatus stulpelių pavadinimuose
    df_out = df_out.loc[:, ~df_out.columns.duplicated()].copy()

    # Jei failas jau DataFrame
    if isinstance(ks_file, pd.DataFrame):
        df_kaunas = ks_file.copy()
    else:
        import io
        content = ks_file.getvalue()
        df_kaunas = pd.read_excel(io.BytesIO(content), engine="openpyxl")

    df_kaunas.columns = [str(c).strip() for c in df_kaunas.columns]

    # Pervadinimas
    rename_map = {}
    for col in df_kaunas.columns:
        col_up = col.strip().upper()
        if "COMP" in col_up:
            rename_map[col] = "Component"
        elif "BIN" in col_up:
            rename_map[col] = "Bin Code"
        elif "QTY" in col_up or "QUANTITY" in col_up:
            rename_map[col] = "Quantity"
    df_kaunas = df_kaunas.rename(columns=rename_map)

    for req in ["Component", "Bin Code", "Quantity"]:
        if req not in df_kaunas.columns:
            df_kaunas[req] = ""

    stock_map = dict(zip(
        df_kaunas["Component"].astype(str),
        df_kaunas["Bin Code"].astype(str)
    ))

    if "No." in df_out.columns:
        key_col = "No."
    elif "Item No." in df_out.columns:
        key_col = "Item No."
    elif "Type" in df_out.columns:
        key_col = "Type"
    else:
        raise ValueError("❌ BOM file has no valid key column (expected 'No.' or 'Item No.')")

    key_series = df_out[key_col]
    if isinstance(key_series, pd.DataFrame):
        key_series = key_series.iloc[:, 0]

    keys = key_series.fillna("").astype(str).tolist()

    df_out["Bin Code"] = [stock_map.get(k, "") for k in keys]

    if "Document No." not in df_out.columns:
        df_out["Document No."] = ""

    mask_no_stock = (df_out["Bin Code"] == "") | (df_out["Bin Code"] == "67-01-01-01")
    df_out.loc[mask_no_stock, "Document No."] = key_series.astype(str) + "/NERA"

    return df_out





# =====================================================
# Pipeline 4.x – Galutinės lentelės
# =====================================================

def pipeline_4_1_job_journal(df_alloc: pd.DataFrame, project_number: str) -> pd.DataFrame:
    """
    Sukuria Job Journal lentelę NAV formatui:
    - Jei nėra stock → prie Document No. prideda '/NERA'
    - Job Task No. = 1144
    - Location Code = KAUNAS
    """
    st.info("📑 Creating Job Journal table...")

    cols = ["Type","No.","Document No.","Job No.","Job Task No.","Quantity","Location Code","Bin Code"]
    df_out = pd.DataFrame(columns=cols)

    for _, row in df_alloc.iterrows():
        doc_no = str(project_number)
        if str(row.get("Bin Code","")) in ("", "67-01-01-01"):
            doc_no += "/NERA"

        df_out = pd.concat([df_out, pd.DataFrame([{
            "Type": "Item",
            "No.": row.get("No."),
            "Document No.": doc_no,
            "Job No.": project_number,
            "Job Task No.": 1144,
            "Quantity": row.get("Quantity",0),
            "Location Code": "KAUNAS",
            "Bin Code": row.get("Bin Code","")
        }])], ignore_index=True)

    return df_out

def pipeline_4_2_nav_table(df_alloc: pd.DataFrame, df_part_no: pd.DataFrame) -> pd.DataFrame:
    """
    Sukuria NAV užsakymo lentelę iš df_alloc (turi turėti 'No.' ir 'Quantity'):
      - Stulpeliai: Type, No., Quantity, Supplier, Profit, Discount
      - Supplier paimamas iš Part_no (Supplier No.) pagal PartNo_A
      - Profit = 17, o jei gamintojas DANFOSS -> 10
      - Discount = 0
    """
    st.info("🛒 Creating NAV order table...")

    # Užtikrinam reikiamus Part_no stulpelius (jie buvo suvienodinti pipeline_3_3_add_nav_numbers)
    needed = ["PartNo_A", "SupplierNo_E", "Manufacturer_D"]
    for col in needed:
        if col not in df_part_no.columns:
            st.error(f"❌ Part_no sheet missing required column: {col}")
            return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount"])

    # Map'ai iš Part_no
    supplier_map = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["Manufacturer_D"].astype(str)))

    # Paruošiam įvestį
    tmp = df_alloc.copy()
    if "No." not in tmp.columns:
        st.error("❌ NAV table source must contain 'No.' column")
        return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount"])

    tmp["No."] = tmp["No."].astype(str)
    tmp["Quantity"] = pd.to_numeric(tmp.get("Quantity", 0), errors="coerce").fillna(0)

    rows = []
    for _, r in tmp.iterrows():
        part_no = str(r["No."])
        qty = float(r.get("Quantity", 0) or 0)
        manuf = manuf_map.get(part_no, "")
        profit = 10 if "DANFOSS" in manuf.upper() else 17
        supplier = supplier_map.get(part_no, 30093)

        rows.append({
            "Type": "Item",
            "No.": part_no,
            "Quantity": qty,
            "Supplier": supplier,
            "Profit": profit,
            "Discount": 0
        })

    return pd.DataFrame(rows, columns=["Type","No.","Quantity","Supplier","Profit","Discount"])



def pipeline_4_3_calculation(df_bom: pd.DataFrame, df_cubic: pd.DataFrame, df_hours: pd.DataFrame,
                             panel_type: str, grounding: str, project_number: str) -> pd.DataFrame:
    """
    Sukuria sąmatos lentelę:
    - Parts cost, CUBIC cost, Hours cost, Smart supply, Wire set, Extra
    - Total, Total+5%, Total+35%
    """
    st.info("💰 Creating Calculation table...")

    # Užtikrinam, kad skaičiai būtų float
    qty_bom = pd.to_numeric(df_bom.get("Quantity", 0), errors="coerce").fillna(0)
    unit_bom = pd.to_numeric(df_bom.get("Unit Cost", 0), errors="coerce").fillna(0)
    parts_cost = (qty_bom * unit_bom).sum() if not df_bom.empty else 0

    if df_cubic is not None and not df_cubic.empty:
        qty_cubic = pd.to_numeric(df_cubic.get("Quantity", 0), errors="coerce").fillna(0)
        unit_cubic = pd.to_numeric(df_cubic.get("Unit Cost", 0), errors="coerce").fillna(0)
        cubic_cost = (qty_cubic * unit_cubic).sum()
    else:
        cubic_cost = 0

    # Hours pagal projektą
    hours_cost = 0
    if df_hours is not None and not df_hours.empty:
        hourly_rate = pd.to_numeric(df_hours.iloc[1, 4], errors="coerce") if df_hours.shape[1] > 4 else 0
        row_match = df_hours[df_hours.iloc[:, 0].astype(str).str.upper() == str(panel_type).upper()]
        hours_value = 0
        if not row_match.empty:
            if grounding == "TT":
                hours_value = pd.to_numeric(row_match.iloc[0, 1], errors="coerce")
            elif grounding == "TN-S":
                hours_value = pd.to_numeric(row_match.iloc[0, 2], errors="coerce")
            elif grounding == "TN-C-S":
                hours_value = pd.to_numeric(row_match.iloc[0, 3], errors="coerce")
        hours_cost = (hours_value if pd.notna(hours_value) else 0) * (hourly_rate if pd.notna(hourly_rate) else 0)

    smart_supply_cost = 9750.0
    wire_set_cost     = 2500.0

    total = parts_cost + cubic_cost + hours_cost + smart_supply_cost + wire_set_cost
    total_plus_5  = total * 1.05
    total_plus_35 = total * 1.35

    df_calc = pd.DataFrame([
        {"Label": "Parts", "Value": parts_cost},
        {"Label": "Cubic", "Value": cubic_cost},
        {"Label": "Hours cost", "Value": hours_cost},
        {"Label": "Smart supply", "Value": smart_supply_cost},
        {"Label": "Wire set", "Value": wire_set_cost},
        {"Label": "Extra", "Value": 0},
        {"Label": "Total", "Value": total},
        {"Label": "Total+5%", "Value": total_plus_5},
        {"Label": "Total+35%", "Value": total_plus_35},
    ])

    return df_calc

def pipeline_4_3_calculation(df_bom: pd.DataFrame, df_cubic: pd.DataFrame, df_hours: pd.DataFrame,
                             panel_type: str, grounding: str, project_number: str) -> pd.DataFrame:
    """
    Sukuria sąmatos lentelę:
    - Parts cost, CUBIC cost, Hours cost, Smart supply, Wire set, Extra
    - Total, Total+5%, Total+35%
    """
    st.info("💰 Creating Calculation table...")

    # --- Parts cost ---
    if not df_bom.empty:
        qty_bom = pd.to_numeric(df_bom["Quantity"], errors="coerce").fillna(0) if "Quantity" in df_bom.columns else 0
        unit_bom = pd.to_numeric(df_bom["Unit Cost"], errors="coerce").fillna(0) if "Unit Cost" in df_bom.columns else 0
        parts_cost = (qty_bom * unit_bom).sum() if isinstance(qty_bom, pd.Series) else 0
    else:
        parts_cost = 0

    # --- CUBIC cost ---
    if df_cubic is not None and not df_cubic.empty:
        if "Unit Cost" in df_cubic.columns and "Quantity" in df_cubic.columns:
            qty_cubic = pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0)
            unit_cubic = pd.to_numeric(df_cubic["Unit Cost"], errors="coerce").fillna(0)
            cubic_cost = (qty_cubic * unit_cubic).sum()
        elif "Total" in df_cubic.columns:  # jei failas turi tik Total sumą
            cubic_cost = pd.to_numeric(df_cubic["Total"], errors="coerce").fillna(0).sum()
        else:
            cubic_cost = 0
    else:
        cubic_cost = 0

    # --- Hours cost ---
    hours_cost = 0
    if df_hours is not None and not df_hours.empty:
        hourly_rate = pd.to_numeric(df_hours.iloc[1, 4], errors="coerce") if df_hours.shape[1] > 4 else 0
        row_match = df_hours[df_hours.iloc[:, 0].astype(str).str.upper() == str(panel_type).upper()]
        hours_value = 0
        if not row_match.empty:
            if grounding == "TT":
                hours_value = pd.to_numeric(row_match.iloc[0, 1], errors="coerce")
            elif grounding == "TN-S":
                hours_value = pd.to_numeric(row_match.iloc[0, 2], errors="coerce")
            elif grounding == "TN-C-S":
                hours_value = pd.to_numeric(row_match.iloc[0, 3], errors="coerce")
        hours_cost = (hours_value if pd.notna(hours_value) else 0) * (hourly_rate if pd.notna(hourly_rate) else 0)

    # --- Fixed costs ---
    smart_supply_cost = 9750.0
    wire_set_cost     = 2500.0

    # --- Totals ---
    total = parts_cost + cubic_cost + hours_cost + smart_supply_cost + wire_set_cost
    total_plus_5  = total * 1.05
    total_plus_35 = total * 1.35

    df_calc = pd.DataFrame([
        {"Label": "Parts", "Value": parts_cost},
        {"Label": "Cubic", "Value": cubic_cost},
        {"Label": "Hours cost", "Value": hours_cost},
        {"Label": "Smart supply", "Value": smart_supply_cost},
        {"Label": "Wire set", "Value": wire_set_cost},
        {"Label": "Extra", "Value": 0},
        {"Label": "Total", "Value": total},
        {"Label": "Total+5%", "Value": total_plus_5},
        {"Label": "Total+35%", "Value": total_plus_35},
    ])

    return df_calc


# =====================================================
# Main render for Stage 3
# =====================================================
def render():
    st.header("Stage 3: BOM Management")

    # 1. Inputs
    inputs = pipeline_2_1_user_inputs()
    if not inputs:
        return

    # 2. File uploads
    files = pipeline_2_2_file_uploads(inputs["rittal"])
    if not files:
        return

    # Tikrinam ar visi failai yra
    required_keys = ["bom", "data", "ks"]
    if not inputs["rittal"]:  # jei Rittal nėra, dar reikia cubic_bom
        required_keys.append("cubic_bom")

    missing = [k for k in required_keys if k not in files]
    if missing:
        st.warning(f"⚠️ Missing required files: {missing}")
        return

    # 3. Jei viskas yra – rodom mygtuką
    if st.button("🚀 Run BOM Processing"):
        # --- pasiimam reikalingus sheetus iš DATA ---
        df_stock       = get_sheet_safe(files["data"], ["Stock"])
        df_accessories = get_sheet_safe(files["data"], ["Accessories"])
        df_part_no     = get_sheet_safe(files["data"], ["Part_no", "Parts_no", "Part no"])
        df_hours       = get_sheet_safe(files["data"], ["Hours"])

        if df_stock is None or df_part_no is None:
            st.error("❌ DATA.xlsx must contain at least 'Stock' and 'Part_no' sheets")
            return

        # --- vykdom pipelines ---
        df_bom = pipeline_3_1_filtering(files["bom"], df_stock)
        df_bom = pipeline_3_2_add_accessories(df_bom, df_accessories)
        df_bom = pipeline_3_3_add_nav_numbers(df_bom, df_part_no)
        df_bom = pipeline_3_4_check_stock(df_bom, files["ks"])

        # --- paimam jau paruoštą Part_no lentelę iš session ---
        df_part_no_ready = st.session_state.get("part_no", df_part_no)

        # --- galutinės lentelės ---
        job_journal = pipeline_4_1_job_journal(df_bom, inputs["project_number"])
        nav_table   = pipeline_4_2_nav_table(df_bom, df_part_no_ready)
        calc_table  = pipeline_4_3_calculation(
            df_bom,
            files.get("cubic_bom"),
            df_hours,
            inputs["panel_type"],
            inputs["grounding"],
            inputs["project_number"]
        )

       # --- išvedimas ---
        st.success("✅ BOM processing complete!")
        
        st.subheader("📑 Job Journal")
        st.dataframe(job_journal, use_container_width=True)
        
        st.subheader("🛒 NAV Table")
        st.dataframe(nav_table, use_container_width=True)
        
        # Nauja dalis: Nerasti NAV numeriai
        missing_parts = st.session_state.get("missing_parts", pd.DataFrame())
        if not missing_parts.empty:
            st.subheader("⚠️ Nerasti NAV numeriai")
            st.dataframe(missing_parts, use_container_width=True)
        
        st.subheader("💰 Calculation")
        st.dataframe(calc_table, use_container_width=True)
