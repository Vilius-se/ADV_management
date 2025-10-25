import pandas as pd
import os
from collections import defaultdict
import re
import csv


def friendly_file_type(filetype: str, filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext in [".xls", ".xlsx"]:
        return "Excel"
    if ext == ".csv":
        return "CSV"
    if ext in [".txt"]:
        return "Text"
    # fallback
    if filetype.startswith("application/vnd.openxmlformats"):
        return "Excel"
    if filetype.startswith("text/"):
        return "Text"
    return filetype.split("/")[-1].capitalize()

def stage1_pipeline_1(df: pd.DataFrame):
    df = df.copy()
    df = df.fillna("")
    df = df.astype(str)
    if 'Line-Article' in df.columns:
        df = df.drop('Line-Article', axis=1)
    if 'Name' in df.columns and 'C.Label' in df.columns:
        mask = df['C.Label'].str.startswith('J')
        df.loc[mask, 'Name'] = (df.loc[mask, 'Name'] + " " + df.loc[mask, 'C.Label']).str.strip()
    if 'C.Label' in df.columns:
        df = df.drop('C.Label', axis=1)
    if 'Name.1' in df.columns and 'C.Label.1' in df.columns:
        mask = df['C.Label.1'].str.startswith('J')
        df.loc[mask, 'Name.1'] = (df.loc[mask, 'Name.1'] + " " + df.loc[mask, 'C.Label.1']).str.strip()
    if 'C.Label.1' in df.columns:
        df = df.drop('C.Label.1', axis=1)
    if 'C.Label.2' in df.columns:
        df = df.drop('C.Label.2', axis=1)
    if 'Name' in df.columns and 'Name.1' in df.columns:
        df = df[df['Name'] != df['Name.1']]
    n_before = len(df)
    df = df.drop_duplicates(ignore_index=True)
    removed_duplicates = n_before - len(df)
    return df, removed_duplicates

def stage1_pipeline_2(df: pd.DataFrame) -> pd.DataFrame:
    block_values = [
        "cable", "Cable", "External", "GNYE", "Interal cable", "Internal Cable",
        "internal cable", "Internal cable", "Power", "power"
    ]
    if 'Line-Function' in df.columns:
        df = df[~df['Line-Function'].isin(block_values)]
    mask_pe = ~df.astype(str).apply(lambda col: col.str.contains("PE", na=False)).any(axis=1)
    df = df[mask_pe]
    return df

def stage1_pipeline_3(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    replacements = {
        "24VDC.": "24VDC",
        "DBL": "DBU",
        "DBLWH": "DBU/WH",
        "DBUWH": "DBU/WH",
        "RDWH": "RD/WH",
        "LBL": "BU",
        "N'": "N2",
        "N´": "N2",
        "N`": "N2",
        "'": "",
        "\u23F6": "~",
        ".P": "",
        ".S": "",
        "1,0": "0,75",
        "\u23E6": "~",
        "0VDC.": "0VDC",
        "24VDC.": "24VDC",
        "24VDC1.": "24VDC1",
        "24VDC2.": "24VDC2"
    }
    for col in df.columns:
        df[col] = df[col].astype(str)
        for old, new in replacements.items():
            df[col] = df[col].str.replace(old, new, regex=False)
    return df

def stage1_pipeline_4(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    col6_name = df.columns[5] if len(df.columns) > 5 else None
    col1_name = df.columns[0] if len(df.columns) > 0 else None
    col2_name = df.columns[1] if len(df.columns) > 1 else None
    if not col6_name or not col1_name or not col2_name:
        return df
    cond_col6 = df[col6_name].isin(['BK', 'LBL'])
    pattern = r'^-F6\d{2}:'
    cond_col1 = df[col1_name].astype(str).str.match(pattern)
    cond_col2 = df[col2_name].astype(str).str.match(pattern)
    to_remove = cond_col6 & (cond_col1 | cond_col2)
    df_filtered = df[~to_remove]
    return df_filtered

def stage1_pipeline_5(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if 'Name' not in df.columns or 'Name.1' not in df.columns:
        return df
    mask_r = df['Name'].str.startswith('-R')
    df.loc[mask_r, ['Name', 'Name.1']] = df.loc[mask_r, ['Name.1', 'Name']].values
    relay_map = {
        row['Name.1']: row['Name']
        for _, row in df.iterrows()
        if isinstance(row['Name.1'], str) and row['Name.1'].startswith('-R')
    }
    rows_out = []
    for idx, row in df.iterrows():
        name_left = row['Name']
        name_right = row['Name.1']
        walked = set()
        while isinstance(name_right, str) and name_right.startswith('-R') and name_right in relay_map:
            if name_right in walked:
                break
            walked.add(name_right)
            name_right = relay_map[name_right]
        new_row = row.copy()
        new_row['Name.1'] = name_right
        rows_out.append(new_row)
    result_df = pd.DataFrame(rows_out, columns=df.columns)
    mask_duplicates = result_df['Name'] == result_df['Name.1']
    result_df = result_df[~mask_duplicates].reset_index(drop=True)
    result_df = result_df.drop_duplicates(ignore_index=True)
    return result_df

def stage1_pipeline_6(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    suffixes = [':A', ':B', ':C', ':D', ':E', ':F']
    def check_pattern(val: str) -> bool:
        if not isinstance(val, str):
            return False
        if val.startswith('-K') and any(val.endswith(suffix) for suffix in suffixes):
            return True
        return False
    mask_name = df['Name'].apply(check_pattern) if 'Name' in df.columns else pd.Series(False, index=df.index)
    mask_name1 = df['Name.1'].apply(check_pattern) if 'Name.1' in df.columns else pd.Series(False, index=df.index)
    mask_to_keep = ~(mask_name | mask_name1)
    df_filtered = df.loc[mask_to_keep].reset_index(drop=True)
    return df_filtered

def stage1_pipeline_7(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if not all(col in df.columns for col in ['Name', 'Name.1', 'Wireno']):
        return df
    df['Wireno'] = df['Wireno'].fillna("").astype(str)
    wireno_map = {}
    name_wireno = df.loc[df['Wireno'] != "", ['Name', 'Wireno']].drop_duplicates(subset=['Name'])
    for _, row in name_wireno.iterrows():
        wireno_map[row['Name']] = row['Wireno']
    name1_wireno = df.loc[df['Wireno'] != "", ['Name.1', 'Wireno']].drop_duplicates(subset=['Name.1'])
    for _, row in name1_wireno.iterrows():
        if row['Name.1'] not in wireno_map:
            wireno_map[row['Name.1']] = row['Wireno']
    def fill_wireno(row):
        if row['Wireno'] != "":
            return row['Wireno']
        for key in [row['Name'], row['Name.1']]:
            if key in wireno_map and wireno_map[key] != "":
                return wireno_map[key]
        return ""
    df['Wireno'] = df.apply(fill_wireno, axis=1)
    return df

def stage1_pipeline_7_1(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    targets = ["-X923:N", "-X924:N", "-X927:N", "-X928:N"]
    has_flags = {key: False for key in targets}
    stored_line_name = None
    stored_line_function = None
    for _, row in df.iterrows():
        names = [row['Name'], row['Name.1']]
        if "-X923:N" in names or "-X924:N" in names:
            has_flags["-X923:N"] = "-X923:N" in names or has_flags["-X923:N"]
            has_flags["-X924:N"] = "-X924:N" in names or has_flags["-X924:N"]
            if stored_line_name is None:
                stored_line_name = row.get("Line-Name", "")
                stored_line_function = row.get("Line-Function", "")
        for val in ("-X927:N", "-X928:N"):
            if val in names:
                has_flags[val] = True
        if all(has_flags.values()) and stored_line_name is not None:
            break
    base_cols = df.columns.tolist()
    def create_new_row(target_name1_val):
        new_row = {col: "" for col in base_cols}
        new_row["Name"] = "-X0100:N"
        new_row["Name.1"] = target_name1_val
        new_row["Line-Name"] = stored_line_name if stored_line_name is not None else "1,5"
        new_row["Line-Function"] = stored_line_function if stored_line_function is not None else ""
        return new_row
    new_rows = []
    for val in targets:
        if has_flags[val]:
            new_rows.append(create_new_row(val))
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    delete_pairs = {
        ("-X923:N", "-X923:N"),
        ("-X924:N", "-X924:N"),
        ("-X923:N", "-X924:N"),
        ("-X924:N", "-X923:N"),
        ("-X924:N", "-X927:N"),
        ("-X923:N", "-X927:N"),
        ("-X927:N", "-X927:N"),
        ("-X928:N", "-X927:N"),
        ("-X928:N", "-X928:N"),
        ("-X923:N", "-X928:N"),
        ("-X924:N", "-X928:N"),
        ("-X927:N", "-X928:N"),
        # Extended pairs with :230VN
        ("-X923:230VN", "-X923:230VN"),
        ("-X924:230VN", "-X924:230VN"),
        ("-X923:230VN", "-X924:230VN"),
        ("-X924:230VN", "-X923:230VN"),
        ("-X924:230VN", "-X927:230VN"),
        ("-X923:230VN", "-X927:230VN"),
        ("-X927:230VN", "-X927:230VN"),
        ("-X928:230VN", "-X927:230VN"),
        ("-X928:230VN", "-X928:230VN"),
        ("-X923:230VN", "-X928:230VN"),
        ("-X924:230VN", "-X928:230VN"),
        ("-X927:230VN", "-X928:230VN"),
        ("-X0100:L", "-X0100:L3"),
    }
    for idx in reversed(df.index):
        row = df.loc[idx]
        val_pair = (row["Name"], row["Name.1"])
        if val_pair in delete_pairs:
            df.drop(index=idx, inplace=True)
    # Wireno fill
    wireno_map = {}
    for idx, row in df.iterrows():
        for key in [row["Name"], row["Name.1"]]:
            wireno_val = row.get("Wireno", "")
            if pd.notna(wireno_val) and wireno_val != "":
                if key not in wireno_map:
                    wireno_map[key] = wireno_val
    def fill_wireno(row):
        if pd.isna(row.get("Wireno", "")) or row.get("Wireno", "") == "":
            for key in [row["Name"], row["Name.1"]]:
                if key in wireno_map:
                    return wireno_map[key]
            return ""
        else:
            return row["Wireno"]
    df["Wireno"] = df.apply(fill_wireno, axis=1)
    return df.reset_index(drop=True)


def stage1_pipeline_8(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if not all(col in df.columns for col in ['Name', 'Name.1']):
        return df

    # 0) Identify X102 rows to force-keep
    force_keep = {
        idx
        for idx, row in df.iterrows()
        if str(row['Name']).startswith('-X102:') or str(row['Name.1']).startswith('-X102:')
    }

    # 1) Build connectivity map, skipping forced rows
    value_to_rows = defaultdict(set)
    for idx, row in df.iterrows():
        if idx in force_keep:
            continue
        name = str(row['Name']).strip()
        name1 = str(row['Name.1']).strip()
        if name and name != 'nan':
            value_to_rows[name].add(idx)
        if name1 and name1 != 'nan':
            value_to_rows[name1].add(idx)

    parent = {}
    def find(x):
        parent.setdefault(x, x)
        if parent[x] != x:
            parent[x] = find(parent[x])
        return parent[x]
    def union(a, b):
        pa, pb = find(a), find(b)
        if pa != pb:
            parent[pa] = pb

    # 2) Initialize all indices
    for idx in df.index:
        parent[idx] = idx

    # 3) Union only groups with >1 connection, skipping forced rows
    for group in value_to_rows.values():
        members = list(group)
        for other in members[1:]:
            union(members[0], other)

    # 4) Build daisy‐chain numbers
    groups = defaultdict(list)
    for idx in df.index:
        groups[find(idx)].append(idx)

    daisy_mapping = {}
    daisy_no = 1
    for members in groups.values():
        if len(members) > 1:
            for idx in members:
                daisy_mapping[idx] = daisy_no
            daisy_no += 1
        else:
            daisy_mapping[members[0]] = 0

    # 5) Enforce DaisyNo=0 for forced rows
    for idx in force_keep:
        daisy_mapping[idx] = 0

    df['DaisyNo'] = df.index.map(daisy_mapping)

    # 6) Final sort
    df = df.sort_values(by=['DaisyNo', 'Wireno'], ascending=[True, True]).reset_index(drop=True)
    return df



# ------------------------------------------------------------------ #
#  Stage-1 Pipeline 9 – assign “Line-Name” / “Line-Function” values
#  and delete superfluous 2,5 rows that are NOT battery related
# ------------------------------------------------------------------ #


def stage1_pipeline_9(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # 0) Force-keep any -X102:* rows
    force_keep = df[
        df['Name'].str.startswith('-X102:', na=False) |
        df['Name.1'].str.startswith('-X102:', na=False)
    ]

    # 1) Work on the remaining rows
    working = df.drop(force_keep.index)

    # ── original pipeline_9 logic ──
    line_name_values = {
        '-F903:2','-F903:N2','-F903.1:2','-F903.1:N2','-F904:2','-F904.1:N2',
        '-T901','-C903:10','-C903:11','-F903.2:2','-F903.2:1','-F901.1:1',
        '-F901:2','-F901:N2','-F903:1','-F903.1:1','-F904:1','-F904.1:1',
        '-K918:11','-K918:14','-F902:2','-F902:N2','-G90A3:OUT+','-G90A3:OUT-',
        '-T901:0V','-T901:0 V',"-T901:0 V'",'-F901:4',"-T901:115 V'",'-T901:400V',
        '-T901:230 V'
    }
    exact_mask = working['Name'].isin(line_name_values) | working['Name.1'].isin(line_name_values)
    working.loc[exact_mask, 'Line-Name'] = '1,5'

    pat_bat = re.compile(r'-G.*BAT', re.IGNORECASE)
    pat_in  = re.compile(r'-G.*IN',  re.IGNORECASE)
    mask_bat = working['Name'].str.contains(pat_bat, na=False) | working['Name.1'].str.contains(pat_bat, na=False)
    mask_in  = working['Name'].str.contains(pat_in,  na=False) | working['Name.1'].str.contains(pat_in,  na=False)
    working.loc[mask_bat,                   'Line-Name'] = '2,5'
    working.loc[mask_in & ~mask_bat,        'Line-Name'] = '1,5'

    mask_bu = (working['Name']=='-F104:5') | (working['Name.1']=='-F104:5')
    working.loc[mask_bu, 'Line-Function'] = 'BU'

    new_symbols = line_name_values  # same set as above
    mask_new = working['Name'].isin(new_symbols) | working['Name.1'].isin(new_symbols)
    working.loc[mask_new & working['Line-Name'].eq(""), 'Line-Name'] = '1,5'

    has_g90  = working[['Name','Name.1']].apply(lambda s: s.str.contains(r'\bG90A3\b',na=False)).any().any()
    has_out = working[['Name','Name.1']].apply(lambda s: s.eq('G90A3:OUT+')).any().any()
    if has_g90 and not has_out:
        base = list(working.columns)
        new = {c:"" for c in base}
        new.update({
            'Name':'G90A3:OUT+','Name.1':'-X0102:24VDC','Wireno':'24VDC',
            'Line-Name':'2,5','Line-Function':'DBU'
        })
        working = pd.concat([working, pd.DataFrame([new])], ignore_index=True)
        
    # old 2,5 filtering rule
    # mask_25 = working['Line-Name']=='2,5'
    # mask_bat = working['Name'].str.contains('BAT',na=False,case=False) | working['Name.1'].str.contains('BAT',na=False,case=False)
    # working = working[~(mask_25 & ~mask_bat)]

    mask_25 = working['Line-Name']=='2,5'
    mask_bat = working['Name'].str.contains('BAT',na=False,case=False) | working['Name.1'].str.contains('BAT',na=False,case=False)
    mask_transformer = working['Name'].str.contains(r'-T8\d', na=False, regex=True) | working['Name.1'].str.contains(r'-T8\d', na=False, regex=True)
    working = working[~(mask_25 & ~mask_bat & ~mask_transformer)]

    # ── end original logic ──

    # 2) Re-attach the -X102:* rows
    result = pd.concat([working, force_keep], ignore_index=True)

    # 3) Reset index (and preserve DaisyNo from prior stage)
    return result.reset_index(drop=True)



def parse_component_functions(df_f):
    # Expects: one-column DataFrame, filter only those starting with '='
    func_col = df_f.columns[0]
    results = {}
    pat = re.compile(r"^=([A-Z0-9_]+)-(.+)$", re.IGNORECASE)

    raw_values = df_f[func_col].tolist()

    for idx, val in enumerate(raw_values):
        if not isinstance(val, str) or not val.strip().startswith("="):
            continue
        stripped = val.strip()
        m = pat.match(stripped)
        if m:
            group = m.group(1).upper()
            symbol = "-" + m.group(2).strip() if not m.group(2).startswith("-") else m.group(2).strip()
            results.setdefault(group, []).append(symbol)

    return results


def stage1_pipeline_10(df: pd.DataFrame, group_symbols: dict) -> pd.DataFrame:
    """
    Enhanced Pipeline 10 with:
    • G90A3 & special terminals (Line-Name = "1,5")
    • Daisy chains (Line-Name = "0,75" for power Wirenos; "1,5" for control Wirenos)
    • Unique terminal rows with no duplicate Name or Name.1
    • Exclude any daisy‐chain row that duplicates a unique (Name,Name.1) or its reverse
    • Preservation and de‐duplication
    """
    SPECIAL_WIRENOS = [
        "0VDC", "24VDC", "24VDC1", "24VDC2",
        "230VL", "230VN", "230VL2", "230VN2",
        "F903/L3", "F903/N"
    ]
    TERMINAL_MAP = {
        "230VL":   "-X0101:230VL",
        "230VN":   "-X0101:230VN",
        "230VL2":  "-X0100:230VL2",
        "230VN2":  "-X0100:230VN2",
        "0VDC":    "-X0102:0VDC",
        "24VDC":   "-X0102:24VDC",
        "24VDC1":  "-X0102:24VDC1",
        "24VDC2":  "-X0102:24VDC2",
        "F903/L3": "-X0100:L3",
        "F903/N":  "-X0100:N",
    }
    CONTROL_WIRENOS = {"F903/N", "F903/L3", "230VL2", "230VN2"}

    def get_first_row_mapping(wireno: str, section: pd.DataFrame) -> str:
        syms = {
            str(s).strip()
            for col in ("Name", "Name.1")
            for s in section[col]
            if pd.notna(s) and str(s).strip() != "nan"
        }
        has_v2    = any("230VL2" in s or "230VN2" in s for s in syms)
        has_g90   = any("-G90A3" in s for s in syms)
        has_f9031 = any("-F903.1" in s for s in syms)
        has_f903  = any("-F903" in s and "-F903." not in s for s in syms)
        has_f904  = any("-F904" in s and "-F904." not in s for s in syms)
        has_f9041 = any("-F904.1" in s for s in syms)

        if wireno == "230VL":
            return "-F901:2" if has_v2 else "-F901.1:2"
        if wireno == "230VN":
            return "-F901:N2" if has_v2 else "-T901:0 V"
        if wireno in ("F903/L", "F903/L3"):
            return "-F903:2"
        if wireno == "230VL2":
            return "-F104:4"
        if wireno == "230VN2":
            return "-F104:6"
        if wireno == "F903/N":
            return "-F903:N2"
        if wireno == "0VDC":
            return "-G90A3:OUT-" if has_g90 else "-C903:11"
        if wireno == "24VDC":
            return "-G90A3:OUT+" if has_g90 else "-C903:10"
        if wireno == "24VDC1":
            if has_f9031: return "-F903.1:2"
            if has_f903:  return "-F903:2"
            if has_f904:  return "-F904:2"
            return "-C903:10"
        if wireno == "24VDC2":
            if has_f904:  return "-F904:2"
            if has_f9041: return "-F904.1:2"
            return "-C903:10"
        return f"-X0101:{wireno}"

    df = df.copy()

    # 1. Preserve -M92X:N rows
    m_pat = r"-M92[345]:N"
    m_rows = df[
        df["Name"].str.contains(m_pat, na=False, regex=True) |
        df["Name.1"].str.contains(m_pat, na=False, regex=True)
    ].copy()

    # 2. Global flags
    has_g90 = df[["Name", "Name.1"]].apply(
        lambda s: s.str.contains(r"-G90A3", na=False)
    ).any().any()
    has_v2 = (
        df["Name"].str.contains("230VL2|230VN2", na=False).any() or
        df["Name.1"].str.contains("230VL2|230VN2", na=False).any() or
        df["Wireno"].str.contains("230VL2|230VN2", na=False).any()
    )

    # 3. Prepare base columns
    base_cols = list(df.columns)
    if "DaisyNo" not in base_cols:
        base_cols.append("DaisyNo")

    # 4. Build unique terminal rows (Line-Name="1,5")
    unique_rows = []
    seen = set()

    def add_unique(name, name1, wireno, func=""):
        if (name, name1) in seen or (name1, name) in seen:
            return
        seen.add((name, name1))
        row = {c: "" for c in base_cols}
        row.update({
            "Name":         name,
            "Name.1":       name1,
            "Wireno":       wireno,
            "DaisyNo":      "0",
            "Line-Name":    "1,5",
            "Line-Function":func
        })
        unique_rows.append(row)

    # Core unique rows (swapped)
    if has_v2:
        add_unique("-F104:4",  "-X0100:230VL2_MAIN",  "230VL2", "BK")
        add_unique("-F104:6",  "-X0100:230VN2_MAIN",  "230VN2", "BU")
        add_unique("-F901:2",  "-X0101:230VL_MAIN",   "230VL",  "RD")
        add_unique("-F901:N2", "-X0101:230VN_MAIN",   "230VN",  "RD/WH")
    else:
        add_unique("-F903:2",  "-X0100:L3_MAIN",      "F903/L3", "BK")
        add_unique("-F903:N2", "-X0100:N_MAIN",       "F903/N",  "BU")
        add_unique("-F901.1:2","-X0101:230VL_MAIN",   "230VL",   "RD")
        add_unique("-T901:0 V","-X0101:230VN_MAIN",   "230VN",   "RD/WH")
        # ── ALWAYS add 24VDC1/24VDC2 unique terminals ──────────────────────────
        add_unique("-F903.1:2", "-X0102:24VDC1_MAIN", "24VDC1")
        add_unique("-F904:2",   "-X0102:24VDC2_MAIN", "24VDC2")
    if has_g90:
        add_unique("-K918:14",  "-X0102:24VDC",  "24VDC", "DBU")
        add_unique("-G90A3:OUT+","-X0102:24VDC_MAIN", "24VDC","DBU")
        add_unique("-G90A3:OUT-","-X0102:0VDC_MAIN",  "0VDC", "DBU/WH")
    else:
        add_unique("-K918:14",  "-X0102:24VDC",  "24VDC","DBU")
        add_unique("-C903:10",  "-X0102:24VDC_MAIN",  "24VDC","DBU")
        add_unique("-C903:11",  "-X0102:0VDC_MAIN",   "0VDC", "DBU/WH")

    # X921 terminal rows
    if has_v2:
        add_unique("-X921:N",  "-X0100:N",      "F903/N", "BU")
        add_unique("-X921:L",  "-X0100:L3",     "F903/L3", "BK")
        add_unique("-X921:L",  "-X0100:230VL2", "230VL2", "BK")
        add_unique("-X921:N",  "-X0100:230VN2", "230VN2", "BU")
    else:
        add_unique("-X921:N",  "-X0100:N",      "F903/N", "BU")
        add_unique("-X921:L",  "-X0100:L3",     "F903/L3", "BK")
        
    # 1) Identify EKF components only if they appear with all of these suffixes:
    ekf_suffixes = {"A1S1", "A2S1", "B1S1", "B2S1", "GND", "BAT+"}
    # Build a mapping from component base name to set of observed suffixes
    comp_suffixes = defaultdict(set)
    for _, row in df.iterrows():
        for side in ("Name", "Name.1"):
            val = row.get(side, "")
            if isinstance(val, str) and ":" in val:
                comp, suffix = val.split(":", 1)
                if suffix in ekf_suffixes:
                    comp_suffixes[comp].add(suffix)
    
    # Filter only those components that have all required suffixes
    ekf_comps = {comp for comp, suffixes in comp_suffixes.items() if ekf_suffixes.issubset(suffixes)}
    
    # 2) Drop any existing rows for those components where Name or Name.1 ends with ":GND"
    df = df[~df.apply(
        lambda r: any(
            r[col].endswith(":GND") and r[col].split(":", 1)[0] in ekf_comps 
            for col in ("Name", "Name.1")
        ),
        axis=1
    )].reset_index(drop=True)
    
    # 3) Create unique grounding rows for each EKF component
    for comp in sorted(ekf_comps):
        add_unique(f"{comp}:GND",    "-X0102:0VDC", "0VDC", "DBU/WH")
        add_unique(f"{comp}:~/-",     "-X0102:0VDC", "0VDC", "DBU/WH")


    # 5. Collect meta for SPECIAL_WIRENOS
    meta = {}
    for _, r in df.iterrows():
        w = r.get("Wireno", "")
        if w in SPECIAL_WIRENOS:
            for c in ("Name", "Name.1"):
                s = str(r[c]).strip()
                if s and s != "nan":
                    meta.setdefault(w, {})[s] = {
                        "Line-Name":     r.get("Line-Name", ""),
                        "Line-Function": r.get("Line-Function", "")
                    }

    # 6. Drop rows where Name == Name.1
    if {"Name", "Name.1"}.issubset(df.columns):
        df = df[df["Name"] != df["Name.1"]]

    # 7. Build daisy chains
    daisy_rows = []
    for wireno in SPECIAL_WIRENOS:
        sec = df[df["Wireno"] == wireno]
        if sec.empty:
            continue
        syms = {
            str(s).strip()
            for s in pd.concat([sec["Name"], sec["Name.1"]])
            if pd.notna(s) and str(s).strip()
        }
        terminal = TERMINAL_MAP.get(wireno, f"-X0101:{wireno}")
        first    = get_first_row_mapping(wireno, sec)
        ln       = "1,5" if wireno in CONTROL_WIRENOS else "0,75"

        for grp, funcs in group_symbols.items():
            matches = sorted({
                s for s in syms
                # for f in funcs if f in s  # <-- REPLACE THIS LINE OLD LOGIC
                for f in funcs if re.match(rf'^{re.escape(f)}(:|$)', s)  # <-- NEW LINE
                if (s, terminal) not in seen and (terminal, s) not in seen and s != first
            })
            if not matches:
                continue

            # first daisy row
            row = {c: "" for c in base_cols}
            row.update({
                "Name":       first,
                "Name.1":     terminal,
                "Wireno":     wireno,
                "DaisyNo":    grp,
                "Line-Name":  ln,
                **meta.get(wireno, {}).get(first, {})
            })
            daisy_rows.append(row)

            # second
            row = {c: "" for c in base_cols}
            row.update({
                "Name":       matches[0],
                "Name.1":     terminal,
                "Wireno":     wireno,
                "DaisyNo":    grp,
                "Line-Name":  ln,
                **meta.get(wireno, {}).get(matches[0], {})
            })
            daisy_rows.append(row)

            # chain rows
            for left, right in zip(matches, matches[1:]):
                row = {c: "" for c in base_cols}
                row.update({
                    "Name":       left,
                    "Name.1":     right,
                    "Wireno":     wireno,
                    "DaisyNo":    grp,
                    "Line-Name":  ln,
                    **meta.get(wireno, {}).get(left, {})
                })
                daisy_rows.append(row)

    # 8. Rebuild output
    rebuilt = {r["Wireno"] for r in daisy_rows}
    out = df[~(df["Wireno"].isin(rebuilt) & df["Wireno"].isin(SPECIAL_WIRENOS))]
    out = pd.concat([out, pd.DataFrame(daisy_rows)], ignore_index=True)
    out = pd.concat([out, pd.DataFrame(unique_rows)], ignore_index=True)
    if not m_rows.empty:
        out = pd.concat([out, m_rows], ignore_index=True)

    # 9. Remove duplicates & finalize, preferring DaisyNo="0"
    if {"Name", "Name.1", "DaisyNo"}.issubset(out.columns):
        # Identify all (Name, Name.1) pairs that have a DaisyNo of "0"
        zero_pairs = set(
            out.loc[out["DaisyNo"] == "0", ["Name", "Name.1"]]
               .itertuples(index=False, name=None)
        )
        # Exclude any rows with the same pair but DaisyNo != "0"
        out = out[~out.apply(
            lambda r: (r["Name"], r["Name.1"]) in zero_pairs and r["DaisyNo"] != "0",
            axis=1
        )]
        # Finally drop any true duplicates, keeping the first
        out = out.drop_duplicates(subset=["Name", "Name.1"], keep="first")
    if "DaisyNo" in out.columns:
        out["DaisyNo"] = out["DaisyNo"].astype(str)


    return out.reset_index(drop=True)






def stage1_pipeline_11(df: pd.DataFrame) -> pd.DataFrame:
    """
    Enhanced pipeline 11 that Line-Name and Line-Function values using EXACT matching only
    2. Searches both Name AND Name.1 columns for exact matches
    3. Removes rows where Name and Name.1 are identical
    4. Sorts by DaisyNo first, then by Line-Name second
    """
    df = df.copy()

    # Step 1: Build exact match dictionaries from both Name and Name.1 columns
    def build_exact_mappings(df):
        line_name_map = {}
        line_function_map = {}

        # Collect all non-empty values for mapping from BOTH Name and Name.1 columns
        for idx, row in df.iterrows():
            for col in ['Name', 'Name.1']:
                value = str(row[col]).strip()
                if not value or value == 'nan' or value == '':
                    continue

                # Collect Line-Name mappings (first non-empty occurrence wins)
                line_name = str(row.get('Line-Name', '')).strip()
                if line_name and line_name != '' and line_name != 'nan':
                    if value not in line_name_map:
                        line_name_map[value] = line_name

                # Collect Line-Function mappings (first non-empty occurrence wins)
                line_function = str(row.get('Line-Function', '')).strip()
                if line_function and line_function != '' and line_function != 'nan':
                    if value not in line_function_map:
                        line_function_map[value] = line_function

        return line_name_map, line_function_map

    line_name_map, line_function_map = build_exact_mappings(df)

    # Step 2: Apply exact matching to fill empty values
    for idx, row in df.iterrows():
        # Check and fill Line-Name if empty
        current_line_name = str(row.get('Line-Name', '')).strip()
        if not current_line_name or current_line_name == '' or current_line_name == 'nan':
            for col in ['Name', 'Name.1']:
                value = str(row[col]).strip()
                if value and value in line_name_map:  # EXACT match only
                    df.at[idx, 'Line-Name'] = line_name_map[value]
                    break

        # Check and fill Line-Function if empty
        current_line_function = str(row.get('Line-Function', '')).strip()
        if not current_line_function or current_line_function == '' or current_line_function == 'nan':
            for col in ['Name', 'Name.1']:
                value = str(row[col]).strip()
                if value and value in line_function_map:  # EXACT match only
                    df.at[idx, 'Line-Function'] = line_function_map[value]
                    break

    # Step 3: Remove rows where Name and Name.1 are identical
    if 'Name' in df.columns and 'Name.1' in df.columns:
        df = df[df['Name'] != df['Name.1']]

    # Step 4: Sort by DaisyNo first, then by Line-Name second
    sort_columns = []
    if 'Wireno' in df.columns:
        sort_columns.append('Wireno')
    if 'DaisyNo' in df.columns:
        sort_columns.append('DaisyNo')
    if 'Line-Name' in df.columns:
        sort_columns.append('Line-Name')

    if sort_columns:
        df = df.sort_values(by=sort_columns, ascending=True).reset_index(drop=True)

    if 'DaisyNo' in df.columns:
        df['DaisyNo'] = df['DaisyNo'].astype(str)

    return df


def stage1_pipeline_12(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare dataframe for manual editing by:
    1. Replacing cells containing 'Error' with empty strings
    2. Identifying blank cells and ensuring proper data types for the data editor
    """
    df = df.copy()

    # Replace any cell containing 'Error' with empty string across entire DataFrame
    # This handles partial matches (case-insensitive)
    df = df.replace(to_replace=r'.*[Ee]rror.*', value='', regex=True)

    # Ensure DaisyNo is string type to avoid Arrow serialization issues
    if 'DaisyNo' in df.columns:
        df['DaisyNo'] = df['DaisyNo'].astype(str)

    # Replace various representations of empty values with empty strings for consistency
    df = df.fillna("")
    df = df.replace(['nan', 'None', 'null', 'NULL'], "")

    # Convert all columns to string for consistent editing
    for col in df.columns:
        df[col] = df[col].astype(str).replace('nan', "")

    return df


def identify_blank_cells(df: pd.DataFrame) -> dict:
    """
    Identify all blank/empty cells in the dataframe and return their locations
    """
    blank_cells = {}

    def is_blank(value):
        if pd.isna(value):
            return True
        str_val = str(value).strip()
        return str_val == "" or str_val.lower() in ['nan', 'none', 'null']

    for col in df.columns:
        blank_rows = []  # Initialize blank_rows for each column
        for idx, value in enumerate(df[col]):
            if is_blank(value):
                blank_rows.append(idx)

        if blank_rows:  # Only add to dict if there are blank rows
            blank_cells[col] = blank_rows

    return blank_cells


def stage1_pipeline_14(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 14 - Remove rows with specific pattern and POWER in DaisyNo

    Removes rows where:
    1. Name or Name.1 matches pattern -F***2:1 (exactly 8 symbols total)
    2. DaisyNo column contains 'POWER' (case-insensitive)

    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame

    Returns:
    --------
    pd.DataFrame
        Filtered DataFrame with matching rows removed
    """

    df = df.copy()

    # Pattern for -F***2:1 where *** can be any 3 characters and total length is 8
    pattern = re.compile(r'^-F.{3}2:1$')

    # Check if DaisyNo column exists
    if 'DaisyNo' not in df.columns:
        return df

    # Create masks for the conditions
    mask_name = df['Name'].astype(str).str.match(pattern, na=False) if 'Name' in df.columns else pd.Series(False,
                                                                                                           index=df.index)
    mask_name1 = df['Name.1'].astype(str).str.match(pattern, na=False) if 'Name.1' in df.columns else pd.Series(False,
                                                                                                                index=df.index)

    # Check for POWER in DaisyNo (case-insensitive)
    mask_power = df['DaisyNo'].astype(str).str.contains('POWER', case=False, na=False)

    # Combine conditions: remove rows where (Name OR Name.1 matches pattern) AND DaisyNo contains POWER
    mask_to_remove = (mask_name | mask_name1) & mask_power

    # Keep rows that don't match the removal criteria
    df_filtered = df[~mask_to_remove].reset_index(drop=True)

    return df_filtered


def stage1_pipeline_15(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 15 - Correct Line-Function values based on Wireno mapping

    Checks if the Line-Function value matches the expected value for each Wireno
    according to POTENTIAL_MAP and corrects it if it doesn't match.

    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame

    Returns:
    --------
    pd.DataFrame
        DataFrame with corrected Line-Function values
    """

    df = df.copy()

    # Define the potential mapping
    POTENTIAL_MAP = {
        "230VL": "RD",
        "230VN": "RD/WH",
        "F903/L": "BK",
        "F903/L3": "BK",
        "230VL2": "BK",
        "F903/N": "BU",
        "230VN2": "BU",
        "0VDC": "DBU/WH",
        "24VDC": "DBU",
        "24VDC1": "DBU",
        "24VDC2": "DBU",
    }

    # Check if required columns exist
    if 'Wireno' not in df.columns or 'Line-Function' not in df.columns:
        return df

    # Iterate through rows and correct Line-Function values
    for idx, row in df.iterrows():
        wireno = str(row['Wireno']).strip()

        # If this Wireno is in our mapping
        if wireno in POTENTIAL_MAP:
            expected_line_function = POTENTIAL_MAP[wireno]
            current_line_function = str(row['Line-Function']).strip()

            # If current value doesn't match expected value, correct it
            if current_line_function != expected_line_function:
                df.at[idx, 'Line-Function'] = expected_line_function

    return df


def stage1_pipeline_16(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 16 – Enhanced with -X102 protection
    """
    df = df.copy()
    
    # 0) Force-keep any -X102:* rows
    force_keep = df[
        df['Name'].str.startswith('-X102:', na=False) |
        df['Name.1'].str.startswith('-X102:', na=False)
    ]
    
    # 1) Work on remaining rows
    working = df.drop(force_keep.index)

    # 1. Drop -Fxxx:x rows except -F9xx:x
    pattern = re.compile(r'^-F(?!9\d{2}:\d$)(?!\d{2}8:2$)(?!8\d{2}:\d$)\d{3}:\d$')
    mask_drop = working['Name'].str.match(pattern, na=False) | working['Name.1'].str.match(pattern, na=False)
    working = working.loc[~mask_drop].reset_index(drop=True)


    # 2. Check for 230VL2/VN2 anywhere
    has_230vl2_or_230vn2 = (
        working['Name'].str.contains('230VL2|230VN2', na=False).any() or
        working['Name.1'].str.contains('230VL2|230VN2', na=False).any() or
        working['Wireno'].str.contains('230VL2|230VN2', na=False).any()
    )

    if has_230vl2_or_230vn2:
        mask_t901 = working['Name'].astype(str).str.startswith('-T901:') | working['Name.1'].astype(str).str.startswith('-T901:')
        working = working.loc[~mask_t901].reset_index(drop=True)

    # 3. Only add T901 supply row if NO 230VL2/VN2 present anywhere
    if not has_230vl2_or_230vn2:
        base_cols = working.columns.tolist()
        new_row = {c: "" for c in base_cols}
        new_row.update({
            'Name': "-T901:0 V'",
            'Name.1': "-T901:115 V'",
            'Wireno': '90:10',
            'Line-Name': '1,5',
            'Line-Function': 'RD',
            'DaisyNo': '0'
        })
        working = pd.concat([working, pd.DataFrame([new_row])], ignore_index=True)

    # 4) Re-attach the -X102:* rows
    result = pd.concat([working, force_keep], ignore_index=True)
    
    return result.reset_index(drop=True)


def stage1_pipeline_17(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 17 – Ensure Line-Name is '1,5' for given symbols in Name or Name.1.

    For any row where Name or Name.1 appears in new_symbols,
    set 'Line-Name' to '1,5'.
    """
    df = df.copy()
    new_symbols = {
        '-F903:2',   '-F903:N2',  '-F903.1:2', '-F903.1:N2', '-F904:2',  '-F904.1:N2',
        '-T901',     '-C903:10',  '-C903:11',  '-F903.2:2',  '-F903.2:1', '-F901.1:1'
        '-F901:2',   '-F901:N2',  '-F903:1',   '-F903.1:1',  '-F904:1',  '-F904.1:1',
        '-K918:11',  '-K918:14',  '-F902:2',   '-F902:N2',   '-G90A3:OUT+', '-G90A3:OUT-'
    }
    mask = df['Name'].isin(new_symbols) | df['Name.1'].isin(new_symbols)
    df.loc[mask, 'Line-Name'] = '1,5'
    return df

def stage1_pipeline_18(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 18 - Remove swapped duplicate rows
    
    Removes duplicate rows where Name and Name.1 are swapped but
    Wireno, Line-Function, and Line-Name are identical.
    
    Example:
    Row 1: Name="-F901:1", Name.1="-F901:2", Wireno="230VL", Line-Function="RD"
    Row 2: Name="-F901:2", Name.1="-F901:1", Wireno="230VL", Line-Function="RD"
    → Keep Row 1, remove Row 2
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame
        
    Returns:
    --------
    pd.DataFrame
        DataFrame with swapped duplicates removed
    """
    df = df.copy()
    
    if len(df) == 0:
        return df
    
    # Check if required columns exist
    required_cols = ['Name', 'Name.1']
    if not all(col in df.columns for col in required_cols):
        print("⚠️ Pipeline 18: Required columns (Name, Name.1) not found")
        return df
    
    print(f"🔄 Pipeline 18: Processing {len(df)} rows for swapped duplicates")
    
    # Create normalized identifier for each row
    def create_normalized_key(row):
        name = str(row.get('Name', '')).strip()
        name1 = str(row.get('Name.1', '')).strip()
        
        # Skip rows where both names are empty or identical
        if not name or not name1 or name == name1:
            # Use original values to avoid false matches
            return (name, name1, 
                   str(row.get('Wireno', '')).strip(),
                   str(row.get('Line-Function', '')).strip(), 
                   str(row.get('Line-Name', '')).strip())
        
        # Sort names alphabetically to normalize swapped pairs
        sorted_names = tuple(sorted([name, name1]))
        
        # Include other fields that must match for true duplicates
        wireno = str(row.get('Wireno', '')).strip()
        line_function = str(row.get('Line-Function', '')).strip()
        line_name = str(row.get('Line-Name', '')).strip()
        
        return (sorted_names, wireno, line_function, line_name)
    
    # Apply normalization and track original indices
    df_with_keys = df.copy()
    df_with_keys['_temp_key'] = df.apply(create_normalized_key, axis=1)
    df_with_keys['_original_index'] = df.index
    
    # Find duplicates before removal
    duplicate_keys = df_with_keys['_temp_key'].duplicated(keep=False)
    duplicates = df_with_keys[duplicate_keys]
    
    if len(duplicates) > 0:
        print(f"🔍 Pipeline 18: Found {len(duplicates)} rows with potential swapped duplicates")
        
        # Group by normalized key to show what's being removed
        for key, group in duplicates.groupby('_temp_key'):
            if len(group) > 1:
                print(f"   📋 Duplicate group (keeping first):")
                for idx, row in group.iterrows():
                    status = "KEEP" if idx == group.index[0] else "REMOVE"
                    print(f"      {status}: {row['Name']} ↔ {row['Name.1']} | {row.get('Wireno', '')} | {row.get('Line-Function', '')}")
    
    # Keep first occurrence of each normalized key
    df_dedup = df_with_keys.drop_duplicates(subset=['_temp_key'], keep='first')
    
    # Remove temporary columns
    df_dedup = df_dedup.drop(['_temp_key', '_original_index'], axis=1)
    
    # Report results
    removed_count = len(df) - len(df_dedup)
    if removed_count > 0:
        print(f"✅ Pipeline 18: Removed {removed_count} swapped duplicate rows")
        print(f"📊 Pipeline 18: {len(df)} → {len(df_dedup)} rows")
    else:
        print("✅ Pipeline 18: No swapped duplicates found")
    
    return df_dedup.reset_index(drop=True)


def stage1_pipeline_19(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 19 - Handle 230VN2/230VL2 presence and related cleanup
    
    Logic:
    1. Check if Wireno column contains 230VN2 or 230VL2
    2. If true:
       - Delete all rows containing -T901 and -F901.1 in Name or Name.1
       - Check if there are any -F901: values in Name or Name.1
       - If no -F901: values exist, create two specific rows
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame
        
    Returns:
    --------
    pd.DataFrame
        Cleaned DataFrame with appropriate modifications
    """
    
    df = df.copy()
    
    # Step 1: Check if Wireno column contains 230VN2 or 230VL2
    has_230vn2_or_230vl2 = False
    if 'Wireno' in df.columns:
        has_230vn2_or_230vl2 = df['Wireno'].str.contains('230VN2|230VL2', na=False).any()
    
    print(f"🔍 Pipeline 19: 230VN2/230VL2 detected in Wireno: {has_230vn2_or_230vl2}")
    
    if has_230vn2_or_230vl2:
        # Step 2: Delete all rows that contain -T901 OR -F901.1 in Name or Name.1
        initial_count = len(df)
        mask_t901_f901_1 = (
            df['Name'].astype(str).str.contains('-T901|-F901\\.1', na=False, regex=True) |
            df['Name.1'].astype(str).str.contains('-T901|-F901\\.1', na=False, regex=True)
        )
        df = df[~mask_t901_f901_1].reset_index(drop=True)
        removed_rows = initial_count - len(df)
        if removed_rows > 0:
            print(f"✂️ Pipeline 19: Removed {removed_rows} rows containing -T901 or -F901.1")
        
        # Step 3: Check if there are any -F901: values in Name or Name.1
        has_f901 = (
            df['Name'].astype(str).str.contains('-F901:', na=False).any() or
            df['Name.1'].astype(str).str.contains('-F901:', na=False).any()
        )
        
        print(f"🔍 Pipeline 19: -F901: values detected: {has_f901}")
        
        # Step 4: If no -F901: values, create the two specific rows
        if not has_f901:
            base_cols = df.columns.tolist()
            
            # Ensure all required columns exist
            required_cols = ['Name', 'Name.1', 'Wireno', 'Line-Name', 'Line-Function', 'DaisyNo']
            for col in required_cols:
                if col not in base_cols:
                    base_cols.append(col)
            
            new_rows = [
                # Row 1: -F901:2 → -X0101:230VL
                {col: "" for col in base_cols} | {
                    'Name': '-F901:2',
                    'Name.1': '-X0101:230VL',
                    'Wireno': '230VL',
                    'Line-Name': '1,5',
                    'Line-Function': 'RD',
                    'DaisyNo': 'CONTROLS'
                },
                # Row 2: -F901:N2 → -X0101:230VN
                {col: "" for col in base_cols} | {
                    'Name': '-F901:N2',
                    'Name.1': '-X0101:230VN',
                    'Wireno': '230VN',
                    'Line-Name': '1,5',
                    'Line-Function': 'RD/WH',
                    'DaisyNo': 'CONTROLS'
                }
            ]
            
            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
            print("✅ Pipeline 19: Added 2 new -F901: rows")
    
    # Final cleanup and sorting
    sort_columns = []
    if 'Wireno' in df.columns:
        sort_columns.append('Wireno')
    if 'DaisyNo' in df.columns:
        sort_columns.append('DaisyNo')
    if 'Line-Name' in df.columns:
        sort_columns.append('Line-Name')
    
    if sort_columns:
        df = df.sort_values(by=sort_columns, ascending=True).reset_index(drop=True)
    
    if 'DaisyNo' in df.columns:
        df['DaisyNo'] = df['DaisyNo'].astype(str)
    
    print(f"📊 Pipeline 19: Final result - {len(df)} rows")
    return df


def stage1_pipeline_20(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 20 - Clean identical Name/Name.1 rows and ensure terminal rows exist

    Logic:
    1. Delete rows where Name == Name.1 (identical columns)
    2. Check for -X923:, -X924:, -X927:, -X928: values in Name or Name.1 (skip those ending in 230VL2)
    3. Check if Wireno contains 230VN2 or 230VL2
    4. Create missing terminal rows:
       - If has 230VN2/VL2: create -X0100:230VN2 rows
       - If no 230VN2/VL2: create -X0100:N rows
    5. Remove duplicate rows where Name and Name.1 combinations are identical
    6. Remove explicit unwanted (Name,Name.1) pairs
    7. Final cleanup and sorting

    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame

    Returns:
    --------
    pd.DataFrame
        Cleaned DataFrame with ensured terminal rows
    """

    df = df.copy()

    # Step 1: Remove rows where Name == Name.1
    if 'Name' in df.columns and 'Name.1' in df.columns:
        df = df[df['Name'] != df['Name.1']].reset_index(drop=True)

    # Step 2: Detect 230VN2/VL2 in Wireno
    has_230vn2_or_230vl2 = ('Wireno' in df.columns and
        df['Wireno'].str.contains('230VN2|230VL2', na=False).any()
    )
    print(f"🔍 Pipeline 20: 230VN2/230VL2 detected: {has_230vn2_or_230vl2}")

    # Step 3: Find target prefixes (skip those containing 230VL2)
    prefixes = ['-X923:', '-X924:', '-X927:', '-X928:']
    found = set()
    # Only include symbols ending with ':N' or ':230VN2'
    for p in prefixes:
        mask1 = (df['Name'].str.startswith(p, na=False)
                 & (df['Name'].str.endswith(':N', na=False)
                    | df['Name'].str.endswith(':230VN2', na=False)))
        mask2 = (df['Name.1'].str.startswith(p, na=False)
                 & (df['Name.1'].str.endswith(':N', na=False)
                    | df['Name.1'].str.endswith(':230VN2', na=False)))
        found.update(df.loc[mask1, 'Name'])
        found.update(df.loc[mask2, 'Name.1'])
    found = [v for v in found if isinstance(v, str) and v]
    if not found:
        # Deduplicate and return early
        if {'Name','Name.1'}.issubset(df.columns):
            df = df.drop_duplicates(['Name','Name.1'], keep='first').reset_index(drop=True)
        return df

    print(f"🔍 Pipeline 20: Valid X-values: {sorted(found)}")

    # Step 4: Create missing terminal rows
    suffix = '230VN2' if has_230vn2_or_230vl2 else 'N'
    terminal = f'-X0100:{suffix}'
    print(f"🔧 Pipeline 20: Using terminal {terminal}")

    base_cols = list(df.columns)
    for c in ['Name','Name.1','Wireno','Line-Name','Line-Function','DaisyNo']:
        if c not in base_cols: base_cols.append(c)

    existing = set()
    for _,r in df.iterrows():
        n,n1 = str(r['Name']).strip(), str(r['Name.1']).strip()
        if n==terminal and n1 in found: existing.add(n1)
        if n1==terminal and n in found: existing.add(n)

    missing = [x for x in found if x not in existing]
    if missing:
        print(f"🔧 Creating {len(missing)} missing terminal rows")
        new_rows = []
        for x in missing:
            sample = df[(df['Name']==x)|(df['Name.1']==x)].iloc[0]
            row = {c:"" for c in base_cols}
            row.update({
                'Name': x,
                'Name.1': terminal,
                'Wireno': suffix if has_230vn2_or_230vl2 else 'F903/N',
                'Line-Name': '1,5',
                'Line-Function': 'BU',
                'DaisyNo': sample.get('DaisyNo','0')
            })
            new_rows.append(row)
            print(f"  ✅ {x} → {terminal}")
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    else:
        print("✅ Pipeline 20: All terminal rows already exist")

    # Step 5: Remove duplicate Name/Name.1 rows
    if {'Name','Name.1'}.issubset(df.columns):
        before = len(df)
        df = df.drop_duplicates(subset=['Name','Name.1'], keep='first').reset_index(drop=True)
        print(f"✂️ Removed {before-len(df)} duplicate Name/Name.1 rows")

    # Step 6: Remove explicit unwanted pairs
    delete_pairs = {
        ("-X923:N","-X923:N"),("-X924:N","-X924:N"),
        ("-X923:N","-X924:N"),("-X924:N","-X923:N"),
        ("-X924:N","-X927:N"),("-X923:N","-X927:N"),
        ("-X927:N","-X927:N"),("-X928:N","-X927:N"),
        ("-X928:N","-X928:N"),("-X923:N","-X928:N"),
        ("-X924:N","-X928:N"),("-X927:N","-X928:N"),
        ("-X923:230VN2","-X923:230VN2"),("-X924:230VN2","-X924:230VN2"),
        ("-X923:230VN2","-X924:230VN2"),("-X924:230VN2","-X923:230VN2"),
        ("-X924:230VN2","-X927:230VN2"),("-X923:230VN2","-X927:230VN2"),
        ("-X927:230VN2","-X927:230VN2"),("-X928:230VN2","-X927:230VN2"),
        ("-X928:230VN2","-X928:230VN2"),("-X923:230VN2","-X928:230VN2"),
        ("-X924:230VN2","-X928:230VN2"),("-X927:230VN2","-X928:230VN2")
    }
    for idx in reversed(df.index):
        pair = (df.at[idx,"Name"], df.at[idx,"Name.1"])
        if pair in delete_pairs:
            df.drop(idx, inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Final cleanup & sorting
    sort_cols = [c for c in ['Wireno','DaisyNo','Line-Name'] if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)
    if 'DaisyNo' in df.columns:
        df['DaisyNo'] = df['DaisyNo'].astype(str)

    print(f"📊 Pipeline 20 final rows: {len(df)}")
    return df

def stage1_pipeline_21(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pipeline 21:
    For Wirenos in {"0VDC","24VDC","24VDC1","24VDC2","230VL","230VN"}:
      – If DaisyNo != "0", set Line-Name to "0,75"
      – If DaisyNo == "0", leave Line-Name unchanged
    Additionally:
      – If Name or Name.1 equals "-F901.1:1" or "-T901:115 V", set Line-Name to "1,5"
      – If Name=='-F104:4' & Name.1=='-X0100:230VL2' or Name=='-F104:6' & Name.1=='-X0100:230VN2',
        set Line-Name to "2,5"
    Finally:
      – If any 230VL2 or 230VN2 appears in Wireno, append four -X102 rows.
    """
    WIRENOS = {"0VDC", "24VDC", "24VDC1", "24VDC2", "230VL", "230VN"}
    SPECIAL_TERMS = {"-F901.1:1", "-T901:115 V"}

    df = df.copy()

    # Ensure required columns exist
    if "Wireno" not in df.columns or "Line-Name" not in df.columns:
        return df.reset_index(drop=True)
    if "DaisyNo" not in df.columns:
        df["DaisyNo"] = ""

    # Base transformation for WIRENOS
    mask_power = df["Wireno"].isin(WIRENOS) & (df["DaisyNo"] != "0")
    df.loc[mask_power, "Line-Name"] = "0,75"

    # Special override for specific terms
    mask_special = (
        df["Name"].isin(SPECIAL_TERMS) |
        df["Name.1"].isin(SPECIAL_TERMS)
    )
    df.loc[mask_special, "Line-Name"] = "1,5"

    # Force Line-Name="2,5" for the two specified pairs
    mask_pair1 = (df["Name"] == "-F104:4") & (df["Name.1"] == "-X0100:230VL2")
    mask_pair2 = (df["Name"] == "-F104:6") & (df["Name.1"] == "-X0100:230VN2")
    df.loc[mask_pair1 | mask_pair2, "Line-Name"] = "2,5"

    # Detect 230VL2 or 230VN2 for appending extra rows
    has_v2 = df["Wireno"].str.contains("230VL2|230VN2", na=False).any()
    if has_v2:
        extra = [
            {"Name": "-X102:1", "Name.1": "-F102:4",  "Wireno": "10:01", "Line-Name": "2,5", "Line-Function": "BK"},
            {"Name": "-X102:2", "Name.1": "-F102:6",  "Wireno": "10:02", "Line-Name": "2,5", "Line-Function": "BK"},
            {"Name": "-X102:3", "Name.1": "-F104:1",  "Wireno": "10:04", "Line-Name": "2,5", "Line-Function": "BK"},
            {"Name": "-X102:4", "Name.1": "-F104:5",  "Wireno": "10:05", "Line-Name": "2,5", "Line-Function": "BU"},
        ]
        # Ensure each extra row includes all existing columns
        for row in extra:
            for col in df.columns:
                row.setdefault(col, "")
        df = pd.concat([df, pd.DataFrame(extra)], ignore_index=True)

    return df.reset_index(drop=True)


def stage1_pipeline_22(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 22 – Handle ventilator components and append VENTS rows.
    Only considers M_PREFS and X_PREFS entries that end with ':N'.
    Special case: when has_k924 is True, map '-M925:N' to '-X924:N'.
    """
    df = df.copy()
    # Flags
    has_k924 = df[['Name', 'Name.1']].apply(lambda col: col.str.startswith('-K924', na=False)).any().any()
    has_vnl2 = df['Wireno'].astype(str).str.contains('230VN2|230VL2', na=False).any()

    # Prefix lists
    M_PREFS = ['-M923:', '-M924:', '-M925:']
    X_PREFS = ['-X923:', '-X924:', '-X927:', '-X928:']

    # Gather matching symbols that end with ':N'
    m_syms = {
        val for col in ['Name', 'Name.1']
        for val in df[col].astype(str)
        for p in M_PREFS
        if val.startswith(p) and val.endswith(':N')
    }
    x_syms = {
        val for col in ['Name', 'Name.1']
        for val in df[col].astype(str)
        for p in X_PREFS
        if val.startswith(p) and val.endswith(':N')
    }

    # Remove ventilator rows
    suffix = '230VN2' if has_vnl2 else 'N'
    def to_remove(row):
        for sym in m_syms | x_syms:
            if (row['Name'].startswith(sym) or row['Name.1'].startswith(sym)) and (
               row['Name'].endswith(suffix) or row['Name.1'].endswith(suffix)):
                return True
        return False
    df_filtered = df.loc[~df.apply(to_remove, axis=1)].reset_index(drop=True)

    # Build new VENTS rows
    target_wireno = '230VN2' if has_vnl2 else 'F903/N'
    terminal = '-X0100:230VN2' if has_vnl2 else '-X0100:N'
    new_rows = []

    # Determine fallback X symbol for M chain when no K924
    fallback_x = sorted(x_syms)[0] if x_syms else terminal

    for sym in sorted(x_syms):
        new_rows.append({
            'Name': sym,
            'Name.1': terminal,
            'Wireno': target_wireno,
            'Line-Name': '1,5',
            'Line-Function': 'BU',
            'DaisyNo': 'VENTS'
        })
    for sym in sorted(m_syms):
        if has_k924 and sym == '-M925:N':
            chain_to = '-X924:N'
        elif has_k924:
            chain_to = sym.replace('-M', '-X')
        else:
            chain_to = fallback_x
        new_rows.append({
            'Name': sym,
            'Name.1': chain_to,
            'Wireno': target_wireno,
            'Line-Name': '1,5',
            'Line-Function': 'BU',
            'DaisyNo': 'VENTS'
        })

    if new_rows:
        df_filtered = pd.concat([df_filtered, pd.DataFrame(new_rows)], ignore_index=True)
    return df_filtered.reset_index(drop=True)



def stage1_pipeline_23(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 23 – Preserve only _MAIN terminal rows.

    1. Identify all (Name, base_Name1) pairs where Name.1 contains '_MAIN'.
       base_Name1 is Name.1 with '_MAIN' stripped.
    2. For each such pair, delete any row where Name and Name.1 match the pair
       but Name.1 does NOT contain '_MAIN'.
    """
    df = df.copy()
    # 1. Collect pairs from rows with '_MAIN' in Name.1
    main_pairs = {
        (row['Name'], row['Name.1'].replace('_MAIN', '', 1))
        for _, row in df.iterrows()
        if isinstance(row.get('Name.1'), str) and '_MAIN' in row['Name.1']
    }
    # 2. Filter out non-_MAIN duplicates
    def keep_row(row):
        for name, base_name1 in main_pairs:
            if row['Name'] == name and row['Name.1'] == base_name1:
                # drop this non-MAIN duplicate
                return False
        return True

    return df[df.apply(keep_row, axis=1)].reset_index(drop=True)

def stage1_pipeline_24(df):
    """
    Stage 1 Pipeline 24 - Remove duplicate 0VDC and 230VN rows for DOOR DaisyNo
    
    Logic:
    1. Check for DaisyNo containing 'DOOR' or 'door' (case insensitive)
    2. For rows with DOOR DaisyNo, remove all 0VDC rows except the first occurrence
    3. For rows with DOOR DaisyNo, remove all 230VN rows except the first occurrence
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame (expected from Stage 1 processing)
        
    Returns:
    --------
    pd.DataFrame
        DataFrame with duplicate DOOR 0VDC and 230VN rows removed
    """
    import pandas as pd
    
    df = df.copy()
    
    print(f"🔧 Pipeline 24: Processing {len(df)} rows for DOOR duplicates")
    
    # Check if required columns exist
    if 'DaisyNo' not in df.columns or 'Wireno' not in df.columns:
        print("⚠️ Pipeline 24: Required columns (DaisyNo, Wireno) not found")
        return df
    
    # Step 1: Identify rows with DOOR DaisyNo (case insensitive)
    door_mask = df['DaisyNo'].astype(str).str.contains('door', case=False, na=False)
    door_rows = df[door_mask]
    non_door_rows = df[~door_mask]
    
    if door_rows.empty:
        print("✅ Pipeline 24: No DOOR rows found, no changes needed")
        return df
    
    print(f"🔍 Pipeline 24: Found {len(door_rows)} DOOR rows")
    
    # Step 2: Process DOOR rows to keep only first occurrence of 0VDC and 230VN
    processed_door_rows = []
    seen_0vdc = False
    seen_230vn = False
    removed_count = 0
    
    for idx, row in door_rows.iterrows():
        wireno = str(row['Wireno']).strip()
        
        # Check for 0VDC
        if wireno == '0VDC':
            if not seen_0vdc:
                seen_0vdc = True
                processed_door_rows.append(row)
                print(f"✅ Pipeline 24: Keeping first 0VDC DOOR row (index {idx})")
            else:
                removed_count += 1
                print(f"🗑️ Pipeline 24: Removing duplicate 0VDC DOOR row (index {idx})")
                continue
        
        # Check for 230VN
        elif wireno == '230VN':
            if not seen_230vn:
                seen_230vn = True
                processed_door_rows.append(row)
                print(f"✅ Pipeline 24: Keeping first 230VN DOOR row (index {idx})")
            else:
                removed_count += 1
                print(f"🗑️ Pipeline 24: Removing duplicate 230VN DOOR row (index {idx})")
                continue
        
        # Keep all other DOOR rows
        else:
            processed_door_rows.append(row)
    
    # Step 3: Combine processed DOOR rows with non-DOOR rows
    if processed_door_rows:
        processed_door_df = pd.DataFrame(processed_door_rows)
        result_df = pd.concat([non_door_rows, processed_door_df], ignore_index=True)
    else:
        result_df = non_door_rows.copy()
    
    # Report results
    if removed_count > 0:
        print(f"✂️ Pipeline 24: Removed {removed_count} duplicate DOOR rows")
        print(f"📊 Pipeline 24: {len(df)} → {len(result_df)} rows")
    else:
        print("✅ Pipeline 24: No duplicate DOOR rows to remove")
    
    return result_df.reset_index(drop=True)


def stage1_pipeline_25(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 1 Pipeline 25 - Add PE (protective earth) grounding rows based on various conditions

    Logic:
    1. Check for "J" symbol after ":" in Name or Name.1 to set has_carel flag
    2. Add K1011 PE row based on has_carel flag
    3. Add K5511 PE row if has_carel is False and K5511 exists
    4. Add transformer PE rows for specific transformers found
    5. Add motor PE rows based on K924 presence and motor patterns
    6. Add capacitor PE rows for C903 and C90A1
    7. Add transformer S2 PE rows for T8x:S2 patterns
    8. Add T901 PE rows if T901 exists
    9. Always add X921:PE row if not exists
    """
    df = df.copy()

    # Collect all symbols from Name and Name.1
    all_symbols = set()
    for col in ('Name', 'Name.1'):
        if col in df.columns:
            all_symbols.update(df[col].astype(str).dropna().tolist())
    all_symbols = {s for s in all_symbols if s and s != 'nan' and s.strip()}

    # 1) has_carel flag
    has_carel = any(':' in s and 'J' in s.split(':', 1)[1] for s in all_symbols)

    # Prepare base columns
    base_cols = list(df.columns)
    for col in ('Name','Name.1','Wireno','Line-Name','Line-Function','DaisyNo'):
        if col not in base_cols:
            base_cols.append(col)

    new_rows = []

    # 2) K1011 row
    if has_carel:
        new_rows.append({
            'Name':'-K1011:GND J24','Name.1':'-XPE:PE','Wireno':'PE',
            'Line-Name':'0,75','Line-Function':'GNYE','DaisyNo':'CONTROL'
        })
    else:
        new_rows.append({
            'Name':'-K1011:PE','Name.1':'-XPE:PE','Wireno':'PE',
            'Line-Name':'0,75','Line-Function':'GNYE','DaisyNo':'CONTROL'
        })

    # 3) K5511 row if no carel
    if not has_carel and any('-K5511:' in s for s in all_symbols):
        new_rows.append({
            'Name':'-K5511:PE','Name.1':'-XPE:PE','Wireno':'PE',
            'Line-Name':'0,75','Line-Function':'GNYE','DaisyNo':'CONTROL'
        })

    # 4) Transformer rows (only if has_carel)
    if has_carel:
        for t in ('-T1011','-T2011','-T3011','-T4011','-T5011','-T5511','-T5711'):
            if any(t in s for s in all_symbols):
                new_rows.append({
                    'Name':f'{t}:-','Name.1':'-XPE:PE','Wireno':'PE',
                    'Line-Name':'0,75','Line-Function':'GNYE','DaisyNo':'CONTROL'
                })

    # 5) Motor rows
    has_k924 = any('-K924' in s for s in all_symbols)
    motors = [m for m in ('-M923','-M924','-M925') if any(m in s for s in all_symbols)]
    if motors:
        if has_k924:
            if '-M925' in motors:
                motor_map = [
                    ('-M923:PE','-X923:PE'),
                    ('-M924:PE','-X924:PE'),
                    ('-M925:PE','-X924:PE'),
                    ('-X923:PE','-XPE:PE'),
                    ('-X924:PE','-XPE:PE'),
                ]
            else:
                motor_map = [
                    ('-M923:PE','-X923:PE'),
                    ('-M924:PE','-X924:PE'),
                    ('-X923:PE','-XPE:PE'),
                    ('-X924:PE','-XPE:PE'),
                ]
        else:
            if len(motors)==3:
                motor_map = [
                    ('-M923:PE','-X923:PE'),
                    ('-M924:PE','-X923:PE'),
                    ('-M925:PE','-X923:PE'),
                    ('-X923:PE','-XPE:PE'),
                ]
            elif len(motors)==2:
                motor_map = [
                    ('-M923:PE','-X923:PE'),
                    ('-M924:PE','-X923:PE'),
                    ('-X923:PE','-XPE:PE'),
                ]
            else:
                motor_map = [
                    ('-M923:PE','-X923:PE'),
                    ('-X923:PE','-XPE:PE'),
                ]
        for n1,n2 in motor_map:
            new_rows.append({'Name':n1,'Name.1':n2,'Wireno':'PE',
                             'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'POWER'})

    # 6) Capacitor rows
    if any('-C903:' in s for s in all_symbols):
        for c in ('-C903:11','-C903:1'):
            new_rows.append({'Name':c,'Name.1':'-XPE:PE','Wireno':'PE',
                             'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'CONTROL'})
    if any('-C90A1:' in s for s in all_symbols):
        for c in ('-C90A1:PE','-C90A1:-'):
            new_rows.append({'Name':c,'Name.1':'-XPE:PE','Wireno':'PE',
                             'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'CONTROL'})

    # 7) T8x:S2 rows
    t8_matches = [s for s in all_symbols if re.match(r'-T8.*:S2', s)]
    if t8_matches:
        for t in ('-T81:S2','-T81.1:S2','-T81.2:S2'):
            new_rows.append({'Name':t,'Name.1':'-XPE:PE','Wireno':'PE',
                             'Line-Name':'2,5','Line-Function':'GNYE','DaisyNo':'POWER'})

    # 8) T901 rows
    if any('-T901:' in s for s in all_symbols):
        for t in ('-T901:PE','-T901:0 V'):
            new_rows.append({'Name':t,'Name.1':'-XPE:PE','Wireno':'PE',
                             'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'CONTROL'})

    # 9) Always add X921:PE if missing
    exists = any(
        (r.get('Name')=='-X921:PE' and r.get('Name.1')=='-XPE:PE')
        for _,r in df.iterrows()
    )
    if not exists:
        new_rows.append({'Name':'-X921:PE','Name.1':'-XPE:PE','Wireno':'PE',
                         'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'CONTROL'})
    
    # 10) X927/X928 PE rows - NEW ADDITION
    x_terminals = ['-X927', '-X928']
    for x_term in x_terminals:
        if any(x_term in s for s in all_symbols):
            new_rows.append({
                'Name':f'{x_term}:PE','Name.1':'-XPE:PE','Wireno':'PE',
                'Line-Name':'1,5','Line-Function':'GNYE','DaisyNo':'POWER'
            })
    # Ensure all keys
    for row in new_rows:
        for c in base_cols:
            row.setdefault(c, '')

    # Append and return
    if new_rows:
        return pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True).reset_index(drop=True)
    else:
        return df.reset_index(drop=True)


def stage2_pipeline_1(uploaded_file) -> pd.DataFrame:
    """
    Stage-2 Pipeline 1 (KOMAX CSV) with conditional space removal:
        - Only remove spaces in Pin / Pin.1 if Betriebsmittelkennzeichen or Betriebsmittelkennzeichen.1 starts with –K
        - Otherwise, preserve internal spaces
    """
    # 1. Load CSV with auto-detected delimiter and encoding
    uploaded_file.seek(0)
    sample = uploaded_file.read(8192).decode("utf-8", errors="replace")
    uploaded_file.seek(0)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
        sep = dialect.delimiter
    except csv.Error:
        sep = ","
    df = None
    for enc in ("utf-8", "latin1", "cp1252"):
        uploaded_file.seek(0)
        try:
            df = pd.read_csv(
                uploaded_file,
                sep=sep,
                engine="python",
                dtype=str,
                encoding=enc,
                keep_default_na=False
            )
            break
        except UnicodeDecodeError:
            continue
    if df is None:
        raise ValueError("Unable to read CSV with utf-8 / latin1 / cp1252.")

    # 2. Split on delimiter if only one wide column (common in Excel export)
    if df.shape[1] == 1:
        first = df.columns[0]
        df = (
            df[first]
            .str.split(sep, expand=True)
            .applymap(lambda x: x.strip() if isinstance(x, str) else x)
        )

    cols = df.columns.tolist()

    # 3. Clean PINs (columns 3 and 10): if >1 colon, keep everything from second colon onward
    def strip_until_second_colon(val: str):
        if not isinstance(val, str): return val
        parts = val.split(":")
        return ":" + ":".join(parts[2:]) if len(parts) > 2 else val

    for idx in (3, 10):
        if idx < len(cols):
            df[cols[idx]] = df[cols[idx]].apply(strip_until_second_colon)

    # -------- CONDITIONAL SPACE REMOVAL ONLY ON -K ROWS ------
    # Mask for rows with -K in Betriebsmittelkennzeichen or Betriebsmittelkennzeichen.1
    name_col  = cols[2] if len(cols) > 2 else None    # Betriebsmittelkennzeichen
    name1_col = cols[9] if len(cols) > 9 else None    # Betriebsmittelkennzeichen.1
    pin_cols  = [cols[3] if len(cols) > 3 else None, cols[10] if len(cols) > 10 else None]
    if name_col and name1_col:
        mask_k = (
            df[name_col].astype(str).str.startswith("-K", na=False)
            | df[name1_col].astype(str).str.startswith("-K", na=False)
        )
        for pin_col in pin_cols:
            if pin_col:
                # Remove all spaces for those rows in Pin/Pin.1 only
                df.loc[mask_k, pin_col] = df.loc[mask_k, pin_col].str.replace(" ", "", regex=False)
    # Leave all other rows untouched (including cases with spaces in pin names)

    # 4. Enforce minimum length in column 17 ('Länge in mm'): min 270
    if 17 < len(cols):
        col_r = cols[17]
        def enforce_min(val):
            try:
                num = float(str(val).replace(",", "."))
                return "270" if num < 270 else str(val)
            except Exception:
                return val
        df[col_r] = df[col_r].apply(enforce_min)

    # 5. Replace single space " " cells with empty string
    df.replace(to_replace="^ $", value="", regex=True, inplace=True)

    # 6. Drop duplicate rows based on columns C-D-J-K (indices 2-3-9-10)
    indices_needed = [2, 3, 9, 10]
    subset_cols = [cols[i] for i in indices_needed if i < len(cols)]
    if subset_cols:
        df = df.drop_duplicates(subset=subset_cols, keep="first").reset_index(drop=True)

    return df


def stage2_pipeline_2(df: pd.DataFrame) -> pd.DataFrame:
    """
    Stage 2 Pipeline 2 - Detect daisy chains and mark ferrule/common connections
    
    Analogous to stage1_pipeline_8 logic, but for KOMAX data:
    - Columns C,D (indices 2,3) are like Name (Betriebsmittelkennzeichen + Pin)
    - Columns J,K (indices 9,10) are like Name.1 (Betriebsmittelkennzeichen.1 + Pin.1)
    - Column E (index 4) and Column L (index 11) need ferrule/common marking
    
    Logic:
    - True endpoints get 'ferrule', middle connections get 'common'
    - Special override: 'common' -> 'common_ferrule' for specific component types
    """
    
    df = df.copy()
    
    # Get column names by index (more reliable than assuming names)
    cols = df.columns.tolist()
    if len(cols) < 12:
        return df  # Not enough columns
    
    col_C = cols[2]   # Betriebsmittelkennzeichen (like Name)
    col_D = cols[3]   # Pin (like Name suffix) 
    col_J = cols[9]   # Betriebsmittelkennzeichen.1 (like Name.1)
    col_K = cols[10]  # Pin.1 (like Name.1 suffix)
    col_E = cols[6]   # Hülse (ferrule marking)
    col_L = cols[13]  # Hülse.1 (ferrule marking)
    
    # Create composite identifiers like "symbol+pin"
    df['Left_ID'] = df[col_C].astype(str) + df[col_D].astype(str)
    df['Right_ID'] = df[col_J].astype(str) + df[col_K].astype(str)
    
    # Build connection map (similar to stage1_pipeline_8 logic)
    value_to_rows = defaultdict(set)
    
    for idx, row in df.iterrows():
        left_id = str(row['Left_ID']).strip()
        right_id = str(row['Right_ID']).strip()
        
        if left_id and left_id != 'nan' and left_id != '':
            value_to_rows[left_id].add(idx)
        if right_id and right_id != 'nan' and right_id != '':
            value_to_rows[right_id].add(idx)
    
    # Union-Find to group connected components
    parent = {}
    
    def find(x):
        if x not in parent:
            parent[x] = x
        if parent[x] != x:
            parent[x] = find(parent[x])
        return parent[x]
    
    def union(x, y):
        px, py = find(x), find(y)
        if px != py:
            parent[px] = py
    
    # Initialize
    for idx in df.index:
        parent[idx] = idx
    
    # Group rows that share common identifiers
    for value, row_indices in value_to_rows.items():
        if len(row_indices) > 1:
            row_list = list(row_indices)
            for i in range(1, len(row_list)):
                union(row_list[0], row_list[i])
    
    # Find groups
    groups = defaultdict(list)
    for idx in df.index:
        root = find(idx)
        groups[root].append(idx)
    
    # Mark ferrule/common for each group
    for root, row_indices in groups.items():
        if len(row_indices) > 1:  # Only process daisy chains (multiple connections)
            # Sort by row index to maintain order
            row_indices.sort()
            
            # Count how many times each component appears in the chain
            component_count = defaultdict(int)
            for idx in row_indices:
                left_id = str(df.at[idx, 'Left_ID']).strip()
                right_id = str(df.at[idx, 'Right_ID']).strip()
                if left_id and left_id != 'nan':
                    component_count[left_id] += 1
                if right_id and right_id != 'nan':
                    component_count[right_id] += 1
            
            # Identify true endpoints (appear only once)
            endpoints = {comp for comp, count in component_count.items() if count == 1}
            
            # Mark each row based on whether its components are endpoints
            for idx in row_indices:
                left_id = str(df.at[idx, 'Left_ID']).strip()
                right_id = str(df.at[idx, 'Right_ID']).strip()
                
                # Left side (column E)
                if left_id in endpoints:
                    df.at[idx, col_E] = 'Ferrule'
                else:
                    df.at[idx, col_E] = 'common'
                
                # Right side (column L)  
                if right_id in endpoints:
                    df.at[idx, col_L] = 'Ferrule'
                else:
                    df.at[idx, col_L] = 'common'
    
    # Helper function to check special component patterns
    def should_be_common_ferrule(component_str):
        """Check if component should be marked as common_ferrule"""
        if not isinstance(component_str, str):
            return False
        
        component_str = component_str.strip()
        
        # Check prefixes
        if component_str.startswith(('-S', '-P', '-Q', '-X010')) and not component_str.startswith('-Q81'):
            return True
            
        # Check exact matches
        if component_str in ('-X923:N', '-X924:N', '-X927:N', '-X928:N'):
            return True
            
        return False
    
    # Apply special override rules for common -> common_ferrule
    for idx in df.index:
        # Check column E (Left side - columns C+D)
        if df.at[idx, col_E] == 'common':
            left_component = str(df.at[idx, col_C]).strip()
            if should_be_common_ferrule(left_component):
                df.at[idx, col_E] = 'common_ferrule'
        
        # Check column L (Right side - columns J+K) 
        if df.at[idx, col_L] == 'common':
            right_component = str(df.at[idx, col_J]).strip()
            if should_be_common_ferrule(right_component):
                df.at[idx, col_L] = 'common_ferrule'
    
    # Clean up temporary columns
    df = df.drop(['Left_ID', 'Right_ID'], axis=1)
    
    return df


def stage2_pipeline_4(df):
    """
    Stage 2 Pipeline 4 - Overwrite with 'Ferrule' based on specific conditions
    
    Logic:
    1. When column C (Betriebsmittelkennzeichen) is in target_prefixes AND 
       column D (Pin) is ':N' → overwrite column G (Hülse) with 'Ferrule'
    2. When column J (Betriebsmittelkennzeichen.1) is in target_prefixes AND 
       column K (Pin.1) is ':N' → overwrite column N (Hülse.1) with 'Ferrule'
    
    Target prefixes: [-M923, -M924, -M925, -X923, -X924, -X927, -X928]
    
    Parameters:
    -----------
    df : pd.DataFrame
        Input DataFrame (KOMAX format)
        Expected column indices:
        - Column C (index 2): Betriebsmittelkennzeichen
        - Column D (index 3): Pin
        - Column G (index 6): Hülse
        - Column J (index 9): Betriebsmittelkennzeichen.1  
        - Column K (index 10): Pin.1
        - Column N (index 13): Hülse.1
    
    Returns:
    --------
    pd.DataFrame
        DataFrame with Ferrule values updated according to the rules
    """
    
    df = df.copy()
    
    # Get column list
    cols = df.columns.tolist()
    
    # Check if we have enough columns
    if len(cols) < 14:
        print(f"⚠️ Pipeline 4: Not enough columns (need at least 14, got {len(cols)})")
        return df
    
    # Define column indices based on the KOMAX Excel structure
    col_C_idx = 2   # Betriebsmittelkennzeichen
    col_D_idx = 3   # Pin  
    col_G_idx = 6   # Hülse (will be updated to 'Ferrule')
    col_J_idx = 9   # Betriebsmittelkennzeichen.1
    col_K_idx = 10  # Pin.1
    col_N_idx = 13  # Hülse.1 (will be updated to 'Ferrule')
    
    # Get column names
    col_C = cols[col_C_idx]
    col_D = cols[col_D_idx] 
    col_G = cols[col_G_idx]
    col_J = cols[col_J_idx]
    col_K = cols[col_K_idx]
    col_N = cols[col_N_idx]
    
    # Target prefixes that trigger the transformation
    target_prefixes = ['-M923', '-M924', '-M925', '-X923', '-X924', '-X927', '-X928', '-XPE']
    
    print(f"🔧 Pipeline 4: Processing {len(df)} rows for Ferrule updates")
    
    # Counters for reporting
    updates_G = 0
    updates_N = 0
    
    # Process each row
    # define the suffixes you want to catch
    suffixes = {':N', ':L3', ':PE'}
    for idx, row in df.iterrows():
        # Rule 1: Check column C and D for updating column G
        col_C_val = str(row[col_C]).strip() if pd.notna(row[col_C]) else ''
        col_D_val = str(row[col_D]).strip() if pd.notna(row[col_D]) else ''
        
        if col_C_val in target_prefixes and col_D_val in suffixes:
            df.at[idx, col_G] = 'Ferrule'
            updates_G += 1
        
        # Rule 2: Check column J and K for updating column N  
        col_J_val = str(row[col_J]).strip() if pd.notna(row[col_J]) else ''
        col_K_val = str(row[col_K]).strip() if pd.notna(row[col_K]) else ''
        
        if col_J_val in target_prefixes and col_K_val in suffixes:
            df.at[idx, col_N] = 'Ferrule'
            updates_N += 1
    
    # Report results
    total_updates = updates_G + updates_N
    if total_updates > 0:
        print(f"✅ Pipeline 4: Updated {updates_G} G columns and {updates_N} N columns to 'Ferrule'")
        print(f"📊 Pipeline 4: Total {total_updates} transformations applied")
    else:
        print("✅ Pipeline 4: No matching conditions found, no updates made")
    
    return df
