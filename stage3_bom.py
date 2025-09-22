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
    remaining = qty_needed

    for _, srow in stock_rows.iterrows():
        bin_code = str(srow.get("Bin Code", "")).strip()
        stock_qty = pd.to_numeric(srow.get("Quantity", 0), errors="coerce")
        if pd.isna(stock_qty):
            stock_qty = 0.0

        if bin_code == "67-01-01-01":  # skip netinkamą lokaciją
            continue
        if remaining <= 0:
            break

        take = min(stock_qty, remaining)
        remaining -= take

        allocations.append({
            "No.": no,
            "Bin Code": bin_code,
            "Allocated Qty": int(round(take))
        })

    # jeigu dar liko neišpildytas kiekis – pridedam eilutę su NERA
    if remaining > 0:
        allocations.append({
            "No.": no,
            "Bin Code": "NERA",
            "Allocated Qty": 0
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
    excluded_components = (
        df_stock[df_stock["Comment"].notna()]["Component"]
        .dropna()
        .astype(str)
    )

    # Normalizuojam pavadinimus
    excluded_norm = (
        excluded_components.str.upper()
        .str.replace(" ", "")
        .str.strip()
        .unique()
    )

    df_bom = df_bom.copy()
    df_bom["Norm_Type"] = (
        df_bom["Type"].astype(str)
        .str.upper()
        .str.replace(" ", "")
        .str.strip()
    )

    # Filtravimas: išmetam VISUS, kurie yra excluded
    filtered = df_bom[~df_bom["Norm_Type"].isin(excluded_norm)].reset_index(drop=True)

    st.success(
        f"✅ BOM filtered: {len(df_bom)} → {len(filtered)} rows "
        f"(removed {len(df_bom) - len(filtered)} items with comments)"
    )

    return filtered.drop(columns=["Norm_Type"])


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


def pipeline_3_3_add_nav_numbers(df_bom, df_part_no_raw):
    if df_bom is None or df_bom.empty:
        return pd.DataFrame()

    # Originalių laukų išsaugojimas
    if "Original Type" not in df_bom.columns:
        df_bom["Original Type"] = df_bom["Type"]
    if "Original Article" not in df_bom.columns and "Article No." in df_bom.columns:
        df_bom["Original Article"] = df_bom["Article No."]

    # Part_no
    df_part_no = df_part_no_raw.copy()
    df_part_no.columns = [
        'PartNo_A', 'PartName_B', 'Desc_C',
        'Manufacturer_D', 'SupplierNo_E', 'UnitPrice_F'
    ]

    # Normalizuojam NAV numerius
    def normalize_no(x):
        try:
            return str(int(float(str(x).replace(",", ".").strip())))
        except:
            return str(x).strip()

    df_part_no["PartNo_A"] = df_part_no["PartNo_A"].map(normalize_no).fillna("")
    df_part_no['Norm_B']   = df_part_no['PartName_B'].astype(str).str.upper().str.replace(" ", "")

    # Map by Type
    map_by_type = dict(zip(df_part_no['Norm_B'], df_part_no['PartNo_A']))

    # BOM
    df_bom = df_bom.copy()
    df_bom['Norm_Type'] = (
        df_bom['Type'].astype(str).str.upper().str.replace(" ", "")
    )
    df_bom['No.'] = df_bom['Norm_Type'].map(map_by_type)

    # Normalizuojam BOM No.
    df_bom["No."] = df_bom["No."].map(normalize_no).fillna("")

    # Backup Quantity
    qty_backup = df_bom.get("Quantity", None)

    # --- Svarbiausia: prieš merge abu key paverčiam į string ---
    df_bom["No."]       = df_bom["No."].astype(str)
    df_part_no["PartNo_A"] = df_part_no["PartNo_A"].astype(str)

    # Merge
    df_bom = df_bom.merge(
        df_part_no[['PartNo_A','Desc_C','Manufacturer_D','SupplierNo_E','UnitPrice_F','Norm_B']],
        left_on='No.', right_on='PartNo_A', how='left'
    )

    # Tvarkymas
    df_bom = df_bom.drop(columns=['Norm_Type','Norm_B','PartNo_A'])
    df_bom = df_bom.rename(columns={
        'Desc_C': 'Description',
        'Manufacturer_D': 'Supplier',
        'SupplierNo_E': 'Supplier No.',
        'UnitPrice_F': 'Unit Cost'
    })

    if "Quantity" not in df_bom.columns and qty_backup is not None:
        df_bom["Quantity"] = qty_backup

    st.session_state["part_no"] = df_part_no
    return df_bom


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
        qty_needed = float(row.get("Quantity", 0))

        stock_rows = row.get("Stock Rows")
        if not isinstance(stock_rows, pd.DataFrame):
            stock_rows = pd.DataFrame(columns=["Bin Code", "Quantity"])

        allocations = allocate_from_stock(no, qty_needed, stock_rows)

        for alloc in allocations:
            doc_no = project_number
            if alloc["Bin Code"] == "NERA":
                doc_no += "/NERA"

            rows.append({
                "Type": "Item",
                "No.": no,
                "Document No.": doc_no,
                "Job No.": project_number,
                "Job Task No.": 1144,
                "Quantity": alloc["Allocated Qty"],
                "Location Code": "KAUNAS",
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
        df_accessories = get_sheet_safe(files["data"], ["Accessories"])
        df_part_no     = get_sheet_safe(files["data"], ["Part_no", "Parts_no", "Part no"])
        df_hours       = get_sheet_safe(files["data"], ["Hours"])
        df_part_code   = get_sheet_safe(files["data"], ["Part_code", "Part code"])

        if df_stock is None or df_part_no is None:
            st.error("❌ DATA.xlsx must contain at least 'Stock' and 'Part_no' sheets")
            return

        # ================================
        # --- BOM processing ---
        # ================================
        df_bom = files["bom"].copy()

        df_bom = pipeline_3_1_filtering(df_bom, df_stock)

        if df_part_code is not None and not df_part_code.empty:
            rename_map_bom = dict(zip(
                df_part_code.iloc[:,0].astype(str).str.strip(),
                df_part_code.iloc[:,1].astype(str).str.strip()
            ))
            df_bom["Type"] = df_bom["Type"].map(lambda x: rename_map_bom.get(x, x))

        df_bom = pipeline_3_2_add_accessories(df_bom, df_accessories)
        df_bom = pipeline_3_3_add_nav_numbers(df_bom, df_part_no)
        df_bom = pipeline_3_4_check_stock(df_bom, files["ks"])

        # ================================
        # --- CUBIC processing ---
        # ================================
        df_cubic = files.get("cubic_bom", pd.DataFrame())
        if not df_cubic.empty:
            df_cubic = pipeline_3_5_prepare_cubic(df_cubic)
            df_cubic = pipeline_3_1_filtering(df_cubic, df_stock)

            if df_part_code is not None and not df_part_code.empty:
                rename_map_cubic = dict(zip(
                    df_part_code.iloc[:,0].astype(str).str.strip(),
                    df_part_code.iloc[:,1].astype(str).str.strip()
                ))
                df_cubic["Type"] = df_cubic["Type"].map(lambda x: rename_map_cubic.get(x, x))

            df_cubic = pipeline_3_3_add_nav_numbers(df_cubic, df_part_no)
            df_cubic = pipeline_3_4_check_stock(df_cubic, files["ks"])

        # ================================
        # --- Missing NAV numbers ---
        # ================================
        missing_bom   = pipeline_4_4_missing_nav(df_bom, "BOM")
        missing_cubic = pipeline_4_4_missing_nav(df_cubic, "CUBIC")

        if not missing_bom.empty:
            st.subheader("📋 Missing NAV numbers (BOM)")
            st.dataframe(missing_bom, use_container_width=True)

        if not missing_cubic.empty:
            st.subheader("📋 Missing NAV numbers (CUBIC)")
            st.dataframe(missing_cubic, use_container_width=True)

        # --- paimam jau paruoštą Part_no lentelę iš session ---
        df_part_no_ready = st.session_state.get("part_no", df_part_no)

        # ================================
        # --- Atskiras Job Journal ---
        # ================================
        job_journal_bom = pipeline_4_1_job_journal(df_bom, inputs["project_number"], source="BOM")

        job_journal_cubic = pd.DataFrame()
        if not df_cubic.empty:
            job_journal_cubic = pipeline_4_1_job_journal(df_cubic, inputs["project_number"], source="CUBIC")

        # ================================
        # --- NAV Tables ---
        # ================================
        nav_table_bom = pipeline_4_2_nav_table(df_bom, df_part_no_ready)

        nav_table_cubic = pd.DataFrame()
        if not df_cubic.empty:
            nav_table_cubic = pipeline_4_2_nav_table(df_cubic, df_part_no_ready)

        # ================================
        # --- Calculation ---
        # ================================
        calc_table = pipeline_4_3_calculation(
            df_bom,
            df_cubic,
            df_hours,
            inputs["panel_type"],
            inputs["grounding"],
            inputs["project_number"]
        )

        # ================================
        # --- Output ---
        # ================================
        st.success("✅ BOM processing complete!")

        st.subheader("📑 Job Journal (BOM)")
        st.dataframe(job_journal_bom, use_container_width=True)

        if not job_journal_cubic.empty:
            st.subheader("📑 Job Journal (CUBIC)")
            st.dataframe(job_journal_cubic, use_container_width=True)

        st.subheader("🛒 NAV Table (BOM)")
        st.dataframe(nav_table_bom, use_container_width=True)

        if not nav_table_cubic.empty:
            st.subheader("🛒 NAV Table (CUBIC)")
            st.dataframe(nav_table_cubic, use_container_width=True)

        st.subheader("💰 Calculation")
        st.dataframe(calc_table, use_container_width=True)

