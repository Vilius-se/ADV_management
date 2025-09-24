import streamlit as st
import pandas as pd
import io
import re
import requests

# ---- NUSTATYMAS: naudoti GitHub failus vietoje upload ----
USE_GITHUB_FILES = True  # <-- čia laikinai užsidedi True, vėliau galėsi pakeisti į False

# Nuorodos į GitHub RAW failus (pvz. master/main branch)
GITHUB_FILES = {
    "cubic_bom": "https://raw.githubusercontent.com/Vilius-se/ADV_management/main/cubic.xls",
    "bom":       "https://raw.githubusercontent.com/Vilius-se/ADV_management/main/bom.XLSX",
    "data":      "https://raw.githubusercontent.com/Vilius-se/ADV_management/main/data.xlsx",
    "ks":        "https://raw.githubusercontent.com/Vilius-se/ADV_management/main/kaunas_stock.xlsm"
}

import requests
import io

# ---- Universalus Excel reader iš GitHub URL ----
def load_excel_from_url(url: str, sheet_name=0):
    """
    Atsisiunčia Excel failą iš GitHub RAW nuorodos ir nuskaito su pandas.
    Automatiškai parenka engine pagal plėtinį.
    """
    r = requests.get(url)
    r.raise_for_status()
    ext = url.lower().split(".")[-1]

    if ext in ["xlsx", "xlsm"]:
        return pd.read_excel(io.BytesIO(r.content), engine="openpyxl", sheet_name=sheet_name)
    elif ext == "xls":
        return pd.read_excel(io.BytesIO(r.content), engine="xlrd", sheet_name=sheet_name)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ---- Įkėlimo iš GitHub logika ----
def load_files_from_github():
    st.info("📥 Loading files from GitHub repository...")
    dfs = {}
    try:
        # --- CUBIC BOM ---
        df_cubic = load_excel_from_url(GITHUB_FILES["cubic_bom"])
        if df_cubic is not None and not df_cubic.empty:
            # Quantity iš E/F/G stulpelių
            if any(col in df_cubic.columns for col in ["E", "F", "G"]):
                df_cubic["Quantity"] = df_cubic[["E", "F", "G"]].bfill(axis=1).iloc[:, 0]
            elif "Quantity" not in df_cubic.columns:
                df_cubic["Quantity"] = 0
            df_cubic["Quantity"] = pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0)

            # Type / Original Type
            if "Item Id" in df_cubic.columns:
                df_cubic["Type"] = df_cubic["Item Id"].astype(str).str.strip()
                df_cubic["Original Type"] = df_cubic["Type"]
            else:
                if "Type" not in df_cubic.columns:
                    df_cubic["Type"] = ""
                if "Original Type" not in df_cubic.columns:
                    df_cubic["Original Type"] = df_cubic["Type"]

            # No.
            if "No." not in df_cubic.columns:
                df_cubic["No."] = df_cubic["Type"]

        dfs["cubic_bom"] = df_cubic

        # --- BOM ---
        df_bom = load_excel_from_url(GITHUB_FILES["bom"])
        if df_bom.shape[1] >= 2:
            colA = df_bom.iloc[:, 0].fillna("").astype(str).str.strip()
            colB = df_bom.iloc[:, 1].fillna("").astype(str).str.strip()
            df_bom["Original Article"] = colA
            df_bom["Original Type"] = colB.where(colB != "", colA)
        else:
            df_bom["Original Article"] = df_bom.iloc[:, 0].fillna("").astype(str).str.strip()
            df_bom["Original Type"] = df_bom["Original Article"]

        if "Type" not in df_bom.columns:
            df_bom["Type"] = df_bom["Original Type"]

        dfs["bom"] = df_bom

        # --- DATA (visi sheet'ai) ---
        dfs["data"] = load_excel_from_url(GITHUB_FILES["data"], sheet_name=None)

        # --- Kaunas stock ---
        dfs["ks"] = load_excel_from_url(GITHUB_FILES["ks"])

        st.success("✅ Files loaded from GitHub")
    except Exception as e:
        st.error(f"❌ Cannot read GitHub files: {e}")
    return dfs


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

def normalize_part_no(df_part_no_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalizuoja Part_no sheetą į standartinius stulpelius:
    PartNo_A, PartName_B, Desc_C, Manufacturer_D, SupplierNo_E, UnitPrice_F
    """
    if df_part_no_raw is None or df_part_no_raw.empty:
        return pd.DataFrame()

    df = df_part_no_raw.copy()
    df = df.rename(columns=lambda c: str(c).strip())

    col_map = {}
    if df.shape[1] >= 1:
        col_map[df.columns[0]] = "PartNo_A"
    if df.shape[1] >= 2:
        col_map[df.columns[1]] = "PartName_B"
    if df.shape[1] >= 3:
        col_map[df.columns[2]] = "Desc_C"
    if df.shape[1] >= 4:
        col_map[df.columns[3]] = "Manufacturer_D"
    if df.shape[1] >= 5:
        col_map[df.columns[4]] = "SupplierNo_E"
    if df.shape[1] >= 6:
        col_map[df.columns[5]] = "UnitPrice_F"

    return df.rename(columns=col_map)

def normalize_no(x):
    """
    Normalizuoja NAV numerius: pašalina kablelius, taškus,
    palieka tik sveiką skaičių kaip string.
    Pvz. '2169732.0' -> '2169732'
    """
    try:
        return str(int(float(str(x).replace(",", ".").strip())))
    except:
        return str(x).strip()


def allocate_from_stock(no, qty_needed, stock_rows):
    allocations = []
    qty_needed = pd.to_numeric(pd.Series([qty_needed]), errors="coerce").fillna(0).iloc[0]
    remaining = int(round(qty_needed))

    if stock_rows is not None and not stock_rows.empty:
        for _, srow in stock_rows.iterrows():
            if remaining <= 0:
                break
            bin_code = str(srow.get("Bin Code", "")).strip()
            stock_qty = pd.to_numeric(pd.Series([srow.get("Quantity", 0)]), errors="coerce").fillna(0).iloc[0]
            if stock_qty <= 0:
                continue
            if bin_code == "67-01-01-01":
                continue

            take = min(int(round(stock_qty)), remaining)
            if take > 0:
                allocations.append({
                    "No.": no,
                    "Bin Code": bin_code,
                    "Allocated Qty": take
                })
                remaining -= take

    if remaining > 0:
        allocations.append({
            "No.": no,
            "Bin Code": "",
            "Allocated Qty": remaining
        })

    return allocations

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
                # Skaitom platesnį bloką (B:G), nes Quantity gali būti E/F/G
                df_cubic = read_excel_any(cubic_bom, skiprows=13, usecols="B,E:F,G")
                df_cubic = df_cubic.rename(columns=lambda c: str(c).strip())

                # Sukuriam Quantity kaip pirmą nenulinę reikšmę tarp E,F,G
                if {"E", "F", "G"}.issubset(df_cubic.columns):
                    df_cubic["Quantity"] = (
                        df_cubic[["E", "F", "G"]]
                        .bfill(axis=1)  # užpildo iš kairės
                        .iloc[:, 0]     # pasiima pirmą reikšmę
                    )
                elif "Quantity" not in df_cubic.columns:
                    df_cubic["Quantity"] = 0

                # Normalizacija
                df_cubic["Quantity"] = pd.to_numeric(df_cubic["Quantity"], errors="coerce").fillna(0)
                df_cubic = df_cubic.rename(columns={"Item Id": "Type"})
                df_cubic["Original Type"] = df_cubic["Type"]

                # Sukuriam No.
                df_cubic["No."] = df_cubic["Type"]

                dfs["cubic_bom"] = df_cubic
            except Exception as e:
                st.error(f"⚠️ Cannot open CUBIC BOM: {e}")

    # --- BOM ---
    st.markdown("<h3 style='color:#0ea5e9; font-weight:700;'>📂 Insert BOM</h3>", unsafe_allow_html=True)
    bom = st.file_uploader("", type=["xls", "xlsx", "xlsm"], key="bom")
    if bom:
        try:
            df_bom = read_excel_any(bom)

            # pasiruošiam pirmus du stulpelius kaip originalius
            if df_bom.shape[1] >= 2:
                colA = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
                colB = df_bom.iloc[:,1].fillna("").astype(str).str.strip()
                df_bom["Original Article"] = colA
                df_bom["Original Type"]    = colB.where(colB != "", colA)
            else:
                df_bom["Original Article"] = df_bom.iloc[:,0].fillna("").astype(str).str.strip()
                df_bom["Original Type"]    = df_bom["Original Article"]

            dfs["bom"] = df_bom
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

def pipeline_3_0_rename_columns(df_bom: pd.DataFrame, df_part_code: pd.DataFrame) -> pd.DataFrame:
    """
    Pervadina BOM stulpelius pagal DATA.xlsx → Part_code.
    1-asis stulpelis = senas pavadinimas, 2-asis = naujas pavadinimas.
    """
    st.info("🔄 Renaming BOM columns according to Part_code...")

    if df_part_code is None or df_part_code.empty:
        st.warning("⚠️ Part_code sheet not found, skipping rename")
        return df_bom

    rename_map = dict(zip(
        df_part_code.iloc[:, 0].astype(str).str.strip(),
        df_part_code.iloc[:, 1].astype(str).str.strip()
    ))

    df_bom = df_bom.rename(columns=rename_map)

    st.success("✅ BOM columns renamed according to Part_code")
    return df_bom
def pipeline_3_1_filtering(df_bom: pd.DataFrame, df_stock: pd.DataFrame, source: str = "BOM") -> pd.DataFrame:
    """
    Filtruojam Project BOM:
    - pašalinam tik tas eilutes, kurias Stock sheet turi su Comment = 'no need'
    - kiti komentarai (Q1, Wurth, ir t.t.) lieka
    """
    st.info(f"🚦 Filtering {source}: removing only 'no need' items...")

    if df_bom is None or df_bom.empty:
        return pd.DataFrame()

    if df_stock is None or df_stock.empty:
        st.warning(f"⚠️ Stock sheet missing or empty — skipping filtering for {source}")
        return df_bom.copy()

    # užtikrinam, kad turim bent 3 stulpelius (Component, ..., Comment)
    cols = list(df_stock.columns)
    if len(cols) < 3:
        st.error("❌ Stock sheet must have at least 3 columns (Component, ..., Comment)")
        return df_bom.copy()

    df_stock = df_stock.rename(columns={cols[0]: "Component", cols[2]: "Comment"})

    # atrenkam tik tuos komponentus, kurių Comment = 'no need'
    excluded_components = (
        df_stock[df_stock["Comment"].astype(str).str.lower().str.strip() == "no need"]["Component"]
        .dropna()
        .astype(str)
    )

    # normalizuojam raktus palyginimui
    excluded_norm = (
        excluded_components
        .str.upper()
        .str.replace(" ", "")
        .str.strip()
        .unique()
    )

    df_in = df_bom.copy()
    df_in["Norm_Type"] = (
        df_in["Type"]
        .astype(str)
        .str.upper()
        .str.replace(" ", "")
        .str.strip()
    )

    filtered = df_in[~df_in["Norm_Type"].isin(excluded_norm)].reset_index(drop=True)

    st.success(
        f"✅ {source} filtered: {len(df_bom)} → {len(filtered)} rows "
        f"(removed {len(df_bom) - len(filtered)} 'no need' items)"
    )

    return filtered.drop(columns=["Norm_Type"])

def pipeline_3_1_filtering_cubic(df_bom: pd.DataFrame, df_stock: pd.DataFrame, source: str = "CUBIC BOM") -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Filtruojam CUBIC BOM:
    - Job Journal: pašalinam visas eilutes, jei Comment != "" (no need, Q1, Wurth ar kitas)
    - NAV Table: pašalinam tik 'no need', o Q1, Wurth ir kiti lieka
    Grąžina tuple: (df_for_journal, df_for_nav)
    """
    st.info(f"🚦 Filtering {source} (different rules for Job Journal & NAV Table)...")

    if df_bom is None or df_bom.empty:
        return pd.DataFrame(), pd.DataFrame()

    if df_stock is None or df_stock.empty:
        st.warning(f"⚠️ Stock sheet missing or empty — skipping filtering for {source}")
        return df_bom.copy(), df_bom.copy()

    # užtikrinam bent 3 stulpelius (Component, ..., Comment)
    cols = list(df_stock.columns)
    if len(cols) < 3:
        st.error("❌ Stock sheet must have at least 3 columns (Component, ..., Comment)")
        return df_bom.copy(), df_bom.copy()

    df_stock = df_stock.rename(columns={cols[0]: "Component", cols[2]: "Comment"})

    # visi su bet kokiu Comment (Job Journal neturi jų turėti)
    excluded_all = df_stock[df_stock["Comment"].astype(str).str.strip() != ""]
    # tik 'no need' (NAV Table neturi jų turėti)
    excluded_no_need = df_stock[df_stock["Comment"].astype(str).str.lower().str.strip() == "no need"]

    excluded_all_norm = (
        excluded_all["Component"].dropna().astype(str).str.upper().str.replace(" ", "").str.strip().unique()
    )
    excluded_no_need_norm = (
        excluded_no_need["Component"].dropna().astype(str).str.upper().str.replace(" ", "").str.strip().unique()
    )

    df_in = df_bom.copy()
    df_in["Norm_Type"] = df_in["Type"].astype(str).str.upper().str.replace(" ", "").str.strip()

    # Job Journal: išmetam visas su Comment
    df_for_journal = df_in[~df_in["Norm_Type"].isin(excluded_all_norm)].reset_index(drop=True)
    # NAV Table: išmetam tik 'no need'
    df_for_nav = df_in[~df_in["Norm_Type"].isin(excluded_no_need_norm)].reset_index(drop=True)

    st.success(
        f"✅ {source} filtered: {len(df_bom)} → {len(df_for_journal)} rows for Job Journal, "
        f"{len(df_for_nav)} rows for NAV Table"
    )

    return df_for_journal.drop(columns=["Norm_Type"]), df_for_nav.drop(columns=["Norm_Type"])




def pipeline_3_2_add_accessories(df_bom: pd.DataFrame, df_accessories: pd.DataFrame) -> pd.DataFrame:
    """
    Prideda accessories pagal DATA.xlsx → Accessories lapą.
    Logika: jei BOM’e yra pagrindinis komponentas, įtraukiami jo accessories.
    Accessories įrašomi į 'Type', kad gautų NAV numerius vėliau.
    """
    st.info("➕ Adding accessories...")

    if df_accessories is None or df_accessories.empty:
        st.warning("⚠️ Accessories sheet not found, skipping")
        return df_bom

    df_out = df_bom.copy()
    added = []

    for _, row in df_bom.iterrows():
        main_item = str(row["Type"]).strip()
        matches = df_accessories[df_accessories.iloc[:, 0].astype(str).str.strip() == main_item]

        for _, acc_row in matches.iterrows():
            acc_values = acc_row.values[1:]  # viskas po pirmo stulpelio
            for i in range(0, len(acc_values), 3):
                if i + 2 >= len(acc_values) or pd.isna(acc_values[i]):
                    break

                acc_item = str(acc_values[i]).strip()  # accessory pavadinimas
                try:
                    acc_qty = float(str(acc_values[i + 1]).replace(",", "."))
                except:
                    acc_qty = 1
                acc_manuf = str(acc_values[i + 2]).strip()

                # accessories įrašom į Type
                df_out = pd.concat([df_out, pd.DataFrame([{
                    "Type": acc_item,
                    "Quantity": acc_qty,
                    "Manufacturer": acc_manuf,
                    "Source": "Accessory"  # žyma identifikacijai
                }])], ignore_index=True)

                added.append(acc_item)

    st.success(f"✅ Added {len(added)} accessories")
    return df_out

def pipeline_3_3_add_nav_numbers(df_bom: pd.DataFrame, df_part_no_raw: pd.DataFrame, source: str = "BOM") -> pd.DataFrame:
    """
    Priskiria NAV numerius ir papildomą info iš Part_no lentelės.
    - 'Type' lieka originalus.
    - 'No.' = sveikas NAV numeris be kablelio (string), jei nerasta -> "".
    - Papildomi laukai: Description, Supplier, Supplier No., Unit Cost (jei Part_no turi).
    """
    st.info(f"🔗 Assigning NAV numbers for {source}...")

    if df_bom is None or df_bom.empty:
        return pd.DataFrame()

    df_out = df_bom.copy()

    if df_part_no_raw is None or df_part_no_raw.empty:
        st.warning("⚠️ Part_no sheet is empty — 'No.' will be left empty")
        df_out["No."] = ""
        return df_out

    # ---- normalizuojam part_no pagal pozicijas (lankstus) ----
    df_part = df_part_no_raw.copy().reset_index(drop=True)
    df_part = df_part.rename(columns=lambda c: str(c).strip())

    col_map = {}
    if df_part.shape[1] >= 1:
        col_map[df_part.columns[0]] = "PartNo_A"
    if df_part.shape[1] >= 2:
        col_map[df_part.columns[1]] = "PartName_B"
    if df_part.shape[1] >= 3:
        col_map[df_part.columns[2]] = "Desc_C"
    if df_part.shape[1] >= 4:
        col_map[df_part.columns[3]] = "Manufacturer_D"
    if df_part.shape[1] >= 5:
        col_map[df_part.columns[4]] = "SupplierNo_E"
    if df_part.shape[1] >= 6:
        col_map[df_part.columns[5]] = "UnitPrice_F"

    df_part = df_part.rename(columns=col_map)

    # Jeigu nėra PartName_B arba PartNo_A — negalime map'inti
    if "PartName_B" not in df_part.columns or "PartNo_A" not in df_part.columns:
        st.warning("⚠️ Part_no lacks PartNo or PartName columns — 'No.' will be left empty")
        df_out["No."] = ""
        return df_out

    # Normalizuojam pavadinimus (PartName_B)
    df_part["Norm_B"] = (
        df_part["PartName_B"]
        .astype(str)
        .str.upper()
        .str.replace(" ", "")
        .str.strip()
    )

    # Normalizuojam PartNo_A — paverčiam į sveiką be .0 (jei įmanoma), paliekam string jei ne
    def norm_partno(x):
        try:
            # jei numeris kaip float (pvz. 2169732.0), paverčiam į int
            s = str(x).strip()
            # bortinam tarpus ir keičiam kablelius
            s2 = s.replace(",", ".")
            if s2 == "":
                return ""
            f = float(s2)
            i = int(round(f))
            return str(i)
        except Exception:
            return str(x).strip()

    df_part["PartNo_A"] = df_part["PartNo_A"].map(norm_partno).fillna("").astype(str)

    # Drop duplicates by Norm_B (keep first)
    df_part = df_part.drop_duplicates(subset=["Norm_B"], keep="first").reset_index(drop=True)

    # Normalizuot BOM Type to match
    df_out["Norm_Type"] = (
        df_out["Type"].astype(str)
        .str.upper()
        .str.replace(" ", "")
        .str.strip()
    )

    # Map: Norm_Type -> PartNo_A
    map_by_type = dict(zip(df_part["Norm_B"], df_part["PartNo_A"]))

    df_out["No."] = df_out["Norm_Type"].map(map_by_type).fillna("").astype(str)

    # Pridėti papildomą informaciją (jei egzistuoja)
    merge_cols = [c for c in ["PartNo_A", "Desc_C", "Manufacturer_D", "SupplierNo_E", "UnitPrice_F", "Norm_B"] if c in df_part.columns]
    if merge_cols:
        # merge pagal No. <-> PartNo_A (mes nekeičiame Type)
        df_out = df_out.merge(
            df_part[merge_cols],
            left_on="No.", right_on="PartNo_A",
            how="left"
        )

        # Pervadinam papildomus lauku pavadinimus jei yra
        rename_map = {}
        if "Desc_C" in df_out.columns:
            rename_map["Desc_C"] = "Description"
        if "Manufacturer_D" in df_out.columns:
            rename_map["Manufacturer_D"] = "Supplier"
        if "SupplierNo_E" in df_out.columns:
            rename_map["SupplierNo_E"] = "Supplier No."
        if "UnitPrice_F" in df_out.columns:
            rename_map["UnitPrice_F"] = "Unit Cost"

        df_out = df_out.rename(columns=rename_map)

        # Pašalinam pagalbinius stulpelius
        dropcols = [c for c in ["Norm_Type", "Norm_B", "PartNo_A"] if c in df_out.columns]
        df_out = df_out.drop(columns=dropcols, errors="ignore")
    else:
        # jei nėra merge stulpelių - tiesiog išvalom Norm_Type
        df_out = df_out.drop(columns=["Norm_Type"], errors="ignore")

    st.success(f"✅ NAV numbers assigned for {source} (found {df_out['No.'].astype(bool).sum()} of {len(df_out)})")
    return df_out

def pipeline_3_4_check_stock(df_bom, ks_file):
    df_out = df_bom.copy()

    # Įsikeliam Kaunas Stock
    if isinstance(ks_file, pd.DataFrame):
        df_stock = ks_file.copy()
    else:
        df_stock = pd.read_excel(io.BytesIO(ks_file.getvalue()), engine="openpyxl")

    df_stock = df_stock.rename(columns=lambda c: str(c).strip())
    # Teisinga tvarka: C (No.), B (Bin Code), D (Quantity)
    df_stock = df_stock[[df_stock.columns[2], df_stock.columns[1], df_stock.columns[3]]]
    df_stock.columns = ["No.", "Bin Code", "Quantity"]

    # Normalizuojam numerius
    df_stock["No."] = df_stock["No."].apply(normalize_no)
    df_out["No."]   = df_out["No."].apply(normalize_no)

    # Sukuriam grupes pagal No.
    stock_groups = {k: v for k, v in df_stock.groupby("No.")}
    df_out["Stock Rows"] = df_out["No."].map(stock_groups)

    return df_out


def pipeline_3_5_prepare_cubic(df_cubic: pd.DataFrame) -> pd.DataFrame:
    if df_cubic is None or df_cubic.empty:
        return pd.DataFrame()

    df_out = df_cubic.copy()
    cols = {c: str(c).strip() for c in df_out.columns}
    df_out = df_out.rename(columns=cols)

    # Quantity iš sujungtų langelių (E:F:G)
    if any(col in df_out.columns for col in ["E", "F", "G"]):
        df_out["Quantity"] = df_out[["E","F","G"]].bfill(axis=1).iloc[:,0]
        df_out["Quantity"] = pd.to_numeric(df_out["Quantity"], errors="coerce").fillna(0)
    elif "Quantity" in df_out.columns:
        df_out["Quantity"] = pd.to_numeric(df_out["Quantity"], errors="coerce").fillna(0)
    else:
        df_out["Quantity"] = 0

    if "Item Id" in df_out.columns:
        df_out["Type"] = df_out["Item Id"].astype(str).str.strip()
        df_out["Original Type"] = df_out["Type"]

    return df_out

# =====================================================
# Pipeline 4.x – Galutinės lentelės
# =====================================================

def pipeline_4_1_job_journal(df_alloc: pd.DataFrame, project_number: str, source: str = "BOM") -> pd.DataFrame:
    st.info(f"📑 Creating Job Journal table from {source}...")

    rows = []
    for _, row in df_alloc.iterrows():
        no = row.get("No.")
        qty_needed = float(row.get("Quantity", 0) or 0)
        stock_rows = row.get("Stock Rows")
        if not isinstance(stock_rows, pd.DataFrame):
            stock_rows = pd.DataFrame(columns=["Bin Code", "Quantity"])

        allocations = allocate_from_stock(no, qty_needed, stock_rows)

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

    return pd.DataFrame(rows)

def pipeline_4_2_nav_table(df_alloc: pd.DataFrame, df_part_no: pd.DataFrame) -> pd.DataFrame:
    """
    Sukuria NAV užsakymo lentelę iš df_alloc:
      - Stulpeliai: Type, No., Quantity, Supplier, Profit, Discount, Description
      - Supplier paimamas iš Part_no pagal PartNo_A
      - Profit = 17, o jei gamintojas DANFOSS -> 10
    """
    st.info("🛒 Creating NAV order table...")

    if df_alloc is None or df_alloc.empty:
        return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])

    # Užtikrinam Part_no stulpelius
    needed = ["PartNo_A", "SupplierNo_E", "Manufacturer_D"]
    for col in needed:
        if col not in df_part_no.columns:
            st.error(f"❌ Part_no sheet missing required column: {col}")
            return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])

    supplier_map = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["SupplierNo_E"]))
    manuf_map    = dict(zip(df_part_no["PartNo_A"].astype(str), df_part_no["Manufacturer_D"].astype(str)))

    tmp = df_alloc.copy()

    # --- Saugikliai ---
    if "No." not in tmp.columns:
        st.error("❌ NAV table source must contain 'No.' column")
        return pd.DataFrame(columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])

    if "Quantity" not in tmp.columns:
        tmp["Quantity"] = 0   # <-- jei trūksta, sukuriam

    if "Description" not in tmp.columns:
        tmp["Description"] = ""

    # Konversijos
    tmp["No."] = tmp["No."].astype(str)
    tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0)

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
            "Discount": 0,
            "Description": r.get("Description", "")
        })

    return pd.DataFrame(rows, columns=["Type","No.","Quantity","Supplier","Profit","Discount","Description"])


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


def pipeline_4_4_missing_nav(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """
    Grąžina lentelę su nerastais NAV numeriais.
    Parodo Original Article, Original Type, Quantity ir 'NAV No.' = None.
    source: "BOM" arba "CUBIC"
    """
    if df is None or df.empty:
        return pd.DataFrame()

    missing = df[df["No."].isna()].copy()
    if missing.empty:
        return pd.DataFrame()

    # Užtikrinam, kad stulpeliai egzistuoja
    if "Original Article" not in missing.columns and df.shape[1] >= 1:
        missing["Original Article"] = df.iloc[:, 0]
    if "Original Type" not in missing.columns and df.shape[1] >= 2:
        missing["Original Type"] = df.iloc[:, 1]

    # Quantity saugiklis
    if "Quantity" in missing.columns:
        qty = pd.to_numeric(missing["Quantity"], errors="coerce").fillna(0).astype(int)
    else:
        qty = 0

    out = pd.DataFrame({
        "Source": source,
        "Original Article (from BOM)": missing.get("Original Article", ""),
        "Original Type (from BOM)": missing.get("Original Type", ""),
        "Quantity": qty,
        "NAV No.": missing["No."]
    })

    return out

# =====================================================
# Main render for Stage 3
# =====================================================
def render():
    st.header("Stage 3: BOM Management")

    # 1. Inputs
    inputs = pipeline_2_1_user_inputs()
    if not inputs:
        return

    # --- Pasirinkimas: GitHub arba įkėlimas ---
    use_github = st.checkbox("📂 Naudoti GitHub failus", value=True)

    # 2. File uploads arba GitHub
    if use_github:
        files = load_files_from_github()
    else:
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

    st.subheader("🔎 Kaunas Stock preview")
    try:
        df_stock_preview = files["ks"].copy()
        st.dataframe(df_stock_preview.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Cannot preview stock: {e}")

    # 3. Jei viskas yra – rodom mygtuką
    if st.button("🚀 Run BOM Processing"):
        # --- pasiimam reikalingus sheetus iš DATA ---
        df_stock       = get_sheet_safe(files["data"], ["Stock"])
        df_part_no_raw = get_sheet_safe(files["data"], ["Part_no", "Parts_no", "Part no"])
        df_part_no     = normalize_part_no(df_part_no_raw)
        df_hours       = get_sheet_safe(files["data"], ["Hours"])
        df_part_code   = get_sheet_safe(files["data"], ["Part_code", "Part code"])

        if df_stock is None or df_part_no.empty:
            st.error("❌ DATA.xlsx must contain at least 'Stock' and 'Part_no' sheets")
            return

        # =====================================================
        # Project BOM processing
        # =====================================================
        st.subheader("📦 Processing Project BOM")
        df_bom = pipeline_3_1_filtering(files["bom"], df_stock)

        # Išsaugom originalius BOM pavadinimus
        if "Original Type" not in df_bom.columns and df_bom.shape[1] >= 2:
            df_bom["Original Type"] = df_bom.iloc[:, 1]
        if "Original Article" not in df_bom.columns and df_bom.shape[1] >= 1:
            df_bom["Original Article"] = df_bom.iloc[:, 0]

        # jei yra Part_code → pakeičiam pavadinimus
        if df_part_code is not None and not df_part_code.empty:
            rename_map = dict(zip(
                df_part_code.iloc[:,0].astype(str).str.strip(),
                df_part_code.iloc[:,1].astype(str).str.strip()
            ))
            df_bom["Type"] = df_bom["Type"].astype(str).map(lambda x: rename_map.get(x, x))

        # NAV numeriai
        df_bom = pipeline_3_3_add_nav_numbers(df_bom, df_part_no, source="Project BOM")

        # NAV Table (kopija be stock)
        df_bom_for_nav = df_bom.copy()
        nav_table_bom = pipeline_4_2_nav_table(df_bom_for_nav, df_part_no)

        # Job Journal (kopija su stock)
        df_bom_for_journal = pipeline_3_4_check_stock(df_bom.copy(), files["ks"])
        job_journal_bom = pipeline_4_1_job_journal(df_bom_for_journal, inputs["project_number"], source="Project BOM")

        # =====================================================
        # CUBIC BOM processing
        # =====================================================
        st.subheader("📦 Processing CUBIC BOM")
        df_cubic = files.get("cubic_bom", pd.DataFrame())
        nav_table_cubic = pd.DataFrame()
        job_journal_cubic = pd.DataFrame()

        if not df_cubic.empty:
            df_cubic = pipeline_3_5_prepare_cubic(df_cubic)

            # Filtravimas pagal komentarus
            df_cubic_for_journal, df_cubic_for_nav = pipeline_3_1_filtering_cubic(
                df_cubic, df_stock, source="CUBIC BOM"
            )

            # NAV Table (CUBIC BOM)
            nav_table_cubic = pipeline_3_2_add_accessories(df_cubic_for_nav, None)  # jei reikia accessories
            nav_table_cubic = pipeline_3_3_add_nav_numbers(df_cubic_for_nav, df_part_no, source="CUBIC BOM")
            nav_table_cubic = pipeline_4_2_nav_table(nav_table_cubic, df_part_no)

            # Job Journal (CUBIC BOM)
            df_cubic_for_journal = pipeline_3_3_add_nav_numbers(df_cubic_for_journal, df_part_no, source="CUBIC BOM")
            df_cubic_for_journal = pipeline_3_4_check_stock(df_cubic_for_journal, files["ks"])
            job_journal_cubic = pipeline_4_1_job_journal(
                df_cubic_for_journal, inputs["project_number"], source="CUBIC BOM"
            )

        # =====================================================
        # Calculation (bendram projektui)
        # =====================================================
        calc_table = pipeline_4_3_calculation(
            df_bom,
            df_cubic,
            df_hours,
            inputs["panel_type"],
            inputs["grounding"],
            inputs["project_number"]
        )

        # =====================================================
        # Output
        # =====================================================
        st.success("✅ BOM processing complete!")

        st.subheader("📑 Job Journal (Project BOM)")
        st.dataframe(job_journal_bom, use_container_width=True)

        if not job_journal_cubic.empty:
            st.subheader("📑 Job Journal (CUBIC BOM)")
            st.dataframe(job_journal_cubic, use_container_width=True)

        st.subheader("🛒 NAV Table (Project BOM)")
        st.dataframe(nav_table_bom, use_container_width=True)

        if not nav_table_cubic.empty:
            st.subheader("🛒 NAV Table (CUBIC BOM)")
            st.dataframe(nav_table_cubic, use_container_width=True)

        st.subheader("💰 Calculation")
        st.dataframe(calc_table, use_container_width=True)



