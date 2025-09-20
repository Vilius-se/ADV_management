import streamlit as st
import pandas as pd
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



import pandas as pd
import streamlit as st

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
    Pvz. Comment = Q1, No need, Wurth, GRM → tokie komponentai nepatenka į BOM.
    """
    st.info("🚦 Filtering BOM according to DATA.xlsx Stock (Comment)...")

    if "Component" not in df_stock.columns or "Comment" not in df_stock.columns:
        st.error("❌ Stock sheet must have 'Component' and 'Comment' columns")
        return df_bom

    exclude = df_stock[df_stock["Comment"].notna()]["Component"].unique()
    filtered = df_bom[~df_bom["Type"].isin(exclude)].reset_index(drop=True)

    st.success(f"✅ BOM filtered: {len(df_bom)} → {len(filtered)} rows")
    return filtered


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
    Rezultatas: grąžina BOM su NAV informacija.
    Originalus Part_no lapas saugomas st.session_state["part_no"].
    """
    # --- Pervadinam stulpelius pagal realų failo turinį ---
    df_part_no = df_part_no_raw.copy()
    df_part_no.columns = [
        'PartNo_A',       # "Item no."
        'PartName_B',     # "Type no."
        'Desc_C',         # "Description"
        'Manufacturer_D', # "Supplier"
        'SupplierNo_E',   # "Supplier No."
        'UnitPrice_F'     # "Cost price / Unit cost DKK"
    ]

    # Normalizuojam Type (PartName_B) kad būtų galima jungti
    df_part_no['Norm_B'] = df_part_no['PartName_B'].astype(str).str.upper().str.replace(" ", "")
    part_map = dict(zip(df_part_no['Norm_B'], df_part_no['PartNo_A']))

    # --- Pridedam NAV numerius į BOM ---
    df_bom = df_bom.copy()
    df_bom['Norm_Type'] = df_bom['Type'].astype(str).str.upper().str.replace(" ", "")
    df_bom['No.'] = df_bom['Norm_Type'].map(part_map)

    # Prijungiam papildomą informaciją iš Part_no lentelės
    df_bom = df_bom.merge(
        df_part_no[['PartNo_A', 'Desc_C', 'Manufacturer_D', 'SupplierNo_E', 'UnitPrice_F', 'Norm_B']],
        left_on='Norm_Type', right_on='Norm_B', how='left'
    )

    # Sutvarkom galutinį formatą
    df_bom = df_bom.drop(columns=['Norm_Type', 'Norm_B'])
    df_bom = df_bom.rename(columns={
        'PartNo_A': 'No.',
        'Desc_C': 'Description',
        'Manufacturer_D': 'Supplier',
        'SupplierNo_E': 'Supplier No.',
        'UnitPrice_F': 'Unit Cost'
    })

    # --- Išsaugom Part_no lentelę į session ---
    st.session_state["part_no"] = df_part_no

    return df_bom

import pandas as pd

def pipeline_3_4_check_stock(df_bom, ks_file):
    """
    Tikrina ar komponentai yra Kauno sandėlyje.
    Failas gali turėti stulpelius bet kokiais pavadinimais – pervadinam.
    """
    df_out = df_bom.copy()

    # Pirmiausia nuskaityti failą
    try:
        df_kaunas = pd.read_excel(ks_file)
    except Exception as e:
        raise ValueError(f"⚠️ Cannot open Kaunas Stock: {e}")

    # Normalizuojam stulpelių pavadinimus
    df_kaunas.columns = [str(c).strip() for c in df_kaunas.columns]

    # Pervadinam į standartą
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

    required_cols = ["Component", "Bin Code", "Quantity"]
    missing = [c for c in required_cols if c not in df_kaunas.columns]
    if missing:
        raise ValueError(f"⚠️ Kaunas Stock missing required columns: {missing}")

    # Sudarom žemėlapį komponento -> sandėlio lokacija
    stock_map = dict(zip(df_kaunas["Component"].astype(str), df_kaunas["Bin Code"].astype(str)))

    # Pridedam Bin Code į BOM
    df_out["Bin Code"] = df_out["Type"].map(lambda x: stock_map.get(str(x), ""))

    # Jei Bin Code tuščias arba "67-01-01-01", žymim kaip NĖRA
    df_out.loc[(df_out["Bin Code"] == "") | (df_out["Bin Code"] == "67-01-01-01"), "Document No."] = df_out["No."].astype(str) + "/NERA"

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
    Sukuria NAV užsakymo lentelę:
    - Type, No., Quantity, Supplier, Profit, Discount
    - Profit = 17, Danfoss → 10
    - Discount = 0
    """
    st.info("🛒 Creating NAV order table...")

    cols = ["Type","No.","Quantity","Supplier","Profit","Discount"]
    df_out = pd.DataFrame(columns=cols)

    supplier_map = dict(zip(df_part_no["PartNo_A"], df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"], df_part_no["Manufacturer_D"]))

    for _, row in df_alloc.iterrows():
        part_no = row.get("No.")
        manuf   = manuf_map.get(part_no,"")
        profit  = 10 if "DANFOSS" in str(manuf).upper() else 17

        df_out = pd.concat([df_out, pd.DataFrame([{
            "Type": "Item",
            "No.": part_no,
            "Quantity": row.get("Quantity",0),
            "Supplier": supplier_map.get(part_no, 30093),
            "Profit": profit,
            "Discount": 0
        }])], ignore_index=True)

    return df_out


def pipeline_4_3_calculation(df_bom: pd.DataFrame, df_cubic: pd.DataFrame, df_hours: pd.DataFrame,
                             panel_type: str, grounding: str, project_number: str) -> pd.DataFrame:
    """
    Sukuria sąmatos lentelę:
    - Parts cost, CUBIC cost, Hours cost, Smart supply, Wire set, Extra
    - Total, Total+5%, Total+35%
    """
    st.info("💰 Creating Calculation table...")

    parts_cost = (df_bom["Quantity"]*df_bom.get("Unit Cost",0)).sum() if not df_bom.empty else 0
    cubic_cost = (df_cubic["Quantity"]*df_cubic.get("Unit Cost",0)).sum() if df_cubic is not None else 0

    # Hours pagal projektą
    hourly_rate = float(df_hours.iloc[1,4]) if df_hours is not None else 0
    row_match = df_hours[df_hours.iloc[:,0].astype(str).str.upper() == str(panel_type).upper()] if df_hours is not None else pd.DataFrame()
    hours_value = 0
    if not row_match.empty:
        if grounding == "TT": hours_value = float(row_match.iloc[0,1])
        elif grounding == "TN-S": hours_value = float(row_match.iloc[0,2])
        elif grounding == "TN-C-S": hours_value = float(row_match.iloc[0,3])
    hours_cost = hours_value * hourly_rate

    smart_supply_cost = 9750.0
    wire_set_cost     = 2500.0

    total = parts_cost + cubic_cost + hours_cost + smart_supply_cost + wire_set_cost
    total_plus_5  = total * 1.05
    total_plus_35 = total * 1.35

    df_calc = pd.DataFrame([
        {"Label":"Parts","Value":parts_cost},
        {"Label":"Cubic","Value":cubic_cost},
        {"Label":"Hours cost","Value":hours_cost},
        {"Label":"Smart supply","Value":smart_supply_cost},
        {"Label":"Wire set","Value":wire_set_cost},
        {"Label":"Extra","Value":0},
        {"Label":"Total","Value":total},
        {"Label":"Total+5%","Value":total_plus_5},
        {"Label":"Total+35%","Value":total_plus_35},
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
        st.warning(f"⚠️ Missing required files")
        return

    # 3. Jei viskas yra – rodom mygtuką
    if st.button("🚀 Run BOM Processing"):
        df_bom = pipeline_3_1_filtering(files["bom"], files["data"]["Stock"])
        df_bom = pipeline_3_2_add_accessories(df_bom, files["data"]["Accessories"])
        df_bom = pipeline_3_3_add_nav_numbers(df_bom, files["data"]["Part_no"])
        df_bom = pipeline_3_4_check_stock(df_bom, files["ks"])

        job_journal = pipeline_4_1_job_journal(df_bom, inputs["project_number"])
        nav_table   = pipeline_4_2_nav_table(df_bom, files["data"]["Part_no"])
        calc_table  = pipeline_4_3_calculation(
            df_bom, files.get("cubic_bom"), files["data"].get("Hours"),
            inputs["panel_type"], inputs["grounding"], inputs["project_number"]
        )

        st.success("✅ BOM processing complete!")
        st.subheader("📑 Job Journal")
        st.dataframe(job_journal, use_container_width=True)

        st.subheader("🛒 NAV Table")
        st.dataframe(nav_table, use_container_width=True)

        st.subheader("💰 Calculation")
        st.dataframe(calc_table, use_container_width=True)
