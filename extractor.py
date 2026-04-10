"""
extractor.py
Reads AuST and CP DOR .xlsm files and returns structured monthly data.
"""
import io
from collections import defaultdict
import pandas as pd
from openpyxl import load_workbook


# ── Cost tables (from SSPC Cost 1/13/25) ─────────────────────────────────────
# cost[0]=AuST good, cost[1]=CP good, cost[2..8]=AuST stations, cost[9..17]=CP stations
COSTS = {
    "BLN":   [17.93,27.08,1.90,2.71,9.74,10.55,11.67,13.16,16.15,0.18,21.83,21.97,22.40,22.61,23.06,13.46,13.71,19.01],
    "SMG":   [18.11,26.57,1.67,2.39,8.15,8.87,9.90,13.83,16.99,0.18,22.16,22.30,22.73,22.94,23.38,13.48,13.73,18.65],
    "CTI":   [17.93,27.08,1.90,2.71,9.74,10.55,11.67,13.16,16.15,0.18,21.83,21.97,22.40,22.61,23.06,13.46,13.71,19.01],
    "BMD":   [17.93,27.08,1.90,2.71,9.74,10.55,11.67,13.16,16.15,0.18,21.83,21.97,22.40,22.61,23.06,13.46,13.71,19.01],
    "MP3":   [9.86, 13.23,0.28,0.28,5.69,5.69,6.03,6.71,8.08,0.03,9.86,9.86,9.90,9.90,9.90,0.01,0.01,13.23],
    "NFS":   [17.93,27.08,1.90,2.71,9.74,10.55,11.67,13.16,16.15,0.18,21.83,21.97,22.40,22.61,23.06,13.46,13.71,19.01],
}

# Per-defect scrap costs (AuST, CenterPoint)
DEFECT_COSTS_AUST = {
    "Leak - valve":16.43,"Leak - bond":16.43,"Leak":16.43,
    "Scratch":16.43,"Dirty":16.43,"Wire at Tip":11.08,"Wire at Hub":13.38,
    "Flash ID":11.08,"Destroyed Tip":11.08,"OD Bulge":16.43,
    "Gap during Reflow":9.21,"Void at Hub Joint":13.38,"OD Flash at Hub Joint":11.08,
    "OD Tip Flash":11.08,"Side Exposed Wire":16.43,"Bubble":9.21,
    "Fiber/ Embedded Particulate":13.38,"Dent":13.38,"Skive":9.21,
    "Kink":16.43,"Burn":16.43,"Damaged/ Melted Hub":16.43,
    "Bad Cut Valve Visual":16.43,"Unknown Reflow Time":9.21,
    "Tip Crumple @ Mandrel Removal":11.08,"Cut Short":16.43,
    "Pin Holes in Heat Shrink":9.21,
    "Failed Shape (Handle Wrong Direction/ doesn't conform to template)":16.43,
    "Extrusions in Wrong Order":9.21,"Tip Bleed":11.08,"Irregular Braid":1.82,
    "Hole at Tip":11.08,"Marker":16.43,"Wrong Valve Position":16.43,
    "Flash OD":11.08,"Stress Marks (Hub)":16.43,"Other":16.43,
}
DEFECT_COSTS_CP = {k:22.08 for k in DEFECT_COSTS_AUST}
DEFECT_COSTS_CP.update({"Leak - valve":22.51,"Leak - bond":22.51,"Leak":22.51,
                         "Unknown Reflow Time":21.94,"Extrusions in Wrong Order":21.94,
                         "Flash OD":22.08,"Stress Marks (Hub)":22.08})

# ── AuST DOR column layout ────────────────────────────────────────────────────
AUST_PRODUCTS = {
    "SMG":  {"svc":14,"ss":32},
    "BLN":  {"svc":15,"ss":71},
    "CTI":  {"svc":16,"ss":110},
    "BMD":  {"svc":17,"ss":149},
    "MP3-S":{"svc":18,"ss":188},
    "MP3-M":{"svc":19,"ss":227},
    "NFE":  {"svc":20,"ss":266},
}
AUST_REASONS = [
    "Bad Liner","Irregular Braid","Open Braid","Unknow Reflow Time","Scratch","Dirty",
    "Fiber/Emb. Part.","Skive","Kink","Burn","Destroyed Tip","Tip Bleed",
    "Valve Orientation","Cut Short","OD Tip Flash","Wire at Tip","Damaged/Melt Hub",
    "Pin Holes at HS","Hole at Tip","Marker","Flash OD","Leak - valve","Leak - bond",
    "Destructive Test","Gap","Layup Wrong Extrusion Order","Irregular Liner at Tip",
    "Stress Marks (Hub)","other"
]
AUST_MAP = {
    "Bad Liner":None,"Irregular Braid":"Irregular Braid","Open Braid":"Other",
    "Unknow Reflow Time":"Unknown Reflow Time","Scratch":"Scratch","Dirty":"Dirty",
    "Fiber/Emb. Part.":"Fiber/ Embedded Particulate","Skive":"Skive","Kink":"Kink",
    "Burn":"Burn","Destroyed Tip":"Destroyed Tip","Tip Bleed":"Tip Bleed",
    "Valve Orientation":"Wrong Valve Position","Cut Short":"Cut Short",
    "OD Tip Flash":None,"Wire at Tip":"Wire at Tip","Damaged/Melt Hub":"Damaged/ Melted Hub",
    "Pin Holes at HS":"Pin Holes in Heat Shrink","Hole at Tip":"Hole at Tip",
    "Marker":"Marker","Flash OD":"Flash OD","Leak - valve":"Leak - valve",
    "Leak - bond":"Leak - bond","Destructive Test":"Other","Gap":"Gap during Reflow",
    "Layup Wrong Extrusion Order":"Extrusions in Wrong Order",
    "Irregular Liner at Tip":"Other","Stress Marks (Hub)":"Stress Marks (Hub)","other":"Other",
}

# ── CP DOR column layout ──────────────────────────────────────────────────────
CP_PRODUCTS = {
    "SMG":  {"svc":13,"ss":32},
    "BLN":  {"svc":14,"ss":66},
    "CTI":  {"svc":15,"ss":100},
    "BMD":  {"svc":16,"ss":134},
    "MP3-S":{"svc":17,"ss":168},
    "MP3-M":{"svc":18,"ss":202},
    "NFE":  {"svc":19,"ss":236},
}
CP_REASONS = [
    "Leak - Bond","Leak - Valve","Scratch","Dirty","Wire At Tip","Flash ID",
    "Destroyed Tip","Od Tip Flash","Fiber/Embedded Particulate","Skive","Kink","Burn",
    "Damage/Melted Hub","Unknown Reflow Time","Cut Short","Pin Holes in Heat Shrink",
    "Failed Shape (Handle Wrong Direction/doest conform to template)","Tip Bleed",
    "Irregular Bread","Hole at Tip","Marker","Bad Liner (Splotchy coating)",
    "Valve Deformed/Dammage/dirty","other"
]
CP_MAP = {
    "Leak - Bond":"Leak - bond","Leak - Valve":"Leak - valve",
    "Scratch":"Scratch","Dirty":"Dirty","Wire At Tip":"Wire at Tip",
    "Flash ID":None,"Destroyed Tip":"Destroyed Tip","Od Tip Flash":None,
    "Fiber/Embedded Particulate":"Fiber/ Embedded Particulate","Skive":"Skive",
    "Kink":"Kink","Burn":"Burn","Damage/Melted Hub":"Damaged/ Melted Hub",
    "Unknown Reflow Time":"Unknown Reflow Time","Cut Short":"Cut Short",
    "Pin Holes in Heat Shrink":"Pin Holes in Heat Shrink",
    "Failed Shape (Handle Wrong Direction/doest conform to template)":
        "Failed Shape (Handle Wrong Direction/ doesn't conform to template)",
    "Tip Bleed":"Tip Bleed","Irregular Bread":"Irregular Braid",
    "Hole at Tip":"Hole at Tip","Marker":"Marker",
    "Bad Liner (Splotchy coating)":None,
    "Valve Deformed/Dammage/dirty":"Bad Cut Valve Visual","other":"Other",
}

ALL_DEFECT_COLS = [
    "Leak","Leak - valve","Leak - bond","Scratch","Dirty","Wire at Tip","Wire at Hub",
    "Flash ID","Destroyed Tip","OD Bulge","Gap during Reflow","Void at Hub Joint",
    "OD Flash at Hub Joint","OD Tip Flash","Side Exposed Wire","Bubble",
    "Fiber/ Embedded Particulate","Dent","Skive","Kink","Burn","Damaged/ Melted Hub",
    "Bad Cut Valve Visual","Unknown Reflow Time","Tip Crumple @ Mandrel Removal",
    "Cut Short","Pin Holes in Heat Shrink",
    "Failed Shape (Handle Wrong Direction/ doesn't conform to template)",
    "Extrusions in Wrong Order","Tip Bleed","Irregular Braid","Hole at Tip","Marker",
    "Wrong Valve Position","Flash OD","Stress Marks (Hub)","Other",
    "White Line at Hub Joint","Smashed","Glue on Shaft",
]

EXCLUDE_FROM_TOP10 = {"Other","Unknown Issue Samples in Retains",
                       "Bad Liner (splotchy coating, etc)","Leak"}


def _sf(v):
    if v is None: return 0.0
    try:
        f = float(v)
        return 0.0 if f != f else f  # NaN guard
    except: return 0.0


def _extract_rows(ws, qty_col, lot_col, date_col, products, reasons, mapping, entity):
    """Extract all months from a DOR worksheet into row dicts."""
    agg = defaultdict(lambda: {
        "qty":0, "rejects":0, "lot":"",
        **{c:0 for c in ALL_DEFECT_COLS}
    })

    for r in range(4, ws.max_row + 1):
        date = ws.cell(r, date_col).value
        if not hasattr(date, "month"): continue
        rt = ws.cell(r, qty_col).value
        if rt not in ("Qty", "Reject"): continue

        lot_raw = str(ws.cell(r, lot_col).value or "").strip()
        tokens  = [t for t in lot_raw.split() if t.startswith(("ML","CL","ml","cl"))]
        lot_key = tokens[0] if tokens else (lot_raw.split()[0] if lot_raw else "")

        month_key = (date.year, date.month)

        for prod, cfg in products.items():
            svc_val = ws.cell(r, cfg["svc"]).value
            if not svc_val: continue
            try: count = int(float(svc_val))
            except: continue
            if count <= 0: continue

            prod_code = "MP3" if prod in ("MP3-S","MP3-M") else ("NFS" if prod=="NFE" else prod)
            key = (month_key, prod_code, entity)

            if rt == "Qty":
                agg[key]["qty"] += count
                if lot_key and not agg[key]["lot"]:
                    agg[key]["lot"] = lot_key
                ss = cfg["ss"]
                for i, reason in enumerate(reasons):
                    dest = mapping.get(reason)
                    if dest is None: continue
                    v = ws.cell(r, ss + i).value
                    if v:
                        try: n = int(float(v))
                        except: n = 0
                        if n > 0:
                            agg[key][dest] += n

            elif rt == "Reject":
                agg[key]["rejects"] += count

    rows = []
    for (month_key, prod_code, ent), data in agg.items():
        qty = data["qty"]; rej = data["rejects"]
        if qty == 0 and rej == 0 and not any(data[c] > 0 for c in ALL_DEFECT_COLS):
            continue
        yld = round((qty - rej) / qty, 4) if qty > 0 else None
        row = {
            "year": month_key[0], "month": month_key[1],
            "product": prod_code, "entity": ent,
            "lot": data["lot"] or "",
            "lot_size": float(qty) if qty > 0 else 0,
            "retains": float(rej) if rej > 0 else 0,
            "yield": yld,
        }
        # combined leak
        row["Leak"] = data.get("Leak - valve",0) + data.get("Leak - bond",0)
        for c in ALL_DEFECT_COLS:
            row[c] = data.get(c, 0)
        rows.append(row)
    return rows


def extract_from_files(aust_bytes, cp_bytes):
    """
    Main entry point.
    aust_bytes, cp_bytes: file-like bytes objects from st.file_uploader

    Returns:
        df_rows   : pd.DataFrame of all scrap rows
        df_copq   : pd.DataFrame of monthly COPQ summary
        months    : sorted list of (year, month) tuples found
        errors    : list of error strings (empty if all OK)
    """
    errors = []

    try:
        wb_aust = load_workbook(io.BytesIO(aust_bytes), data_only=True, keep_vba=True)
        ws_aust = wb_aust["AuST-SSPC DATA"]
    except Exception as e:
        errors.append(f"Could not read AuST file: {e}")
        return None, None, [], errors

    try:
        wb_cp = load_workbook(io.BytesIO(cp_bytes), data_only=True, keep_vba=True)
        ws_cp = wb_cp["CP-SSPC DATA"]
    except Exception as e:
        errors.append(f"Could not read CP file: {e}")
        return None, None, [], errors

    aust_rows = _extract_rows(ws_aust, 13, 22, 2, AUST_PRODUCTS, AUST_REASONS, AUST_MAP, "AuST")
    cp_rows   = _extract_rows(ws_cp,   12, 22, 2, CP_PRODUCTS,   CP_REASONS,   CP_MAP,   "CenterPoint")

    all_rows = aust_rows + cp_rows
    if not all_rows:
        errors.append("No data found in the uploaded files.")
        return None, None, [], errors

    df = pd.DataFrame(all_rows)
    months = sorted(df[["year","month"]].drop_duplicates().apply(tuple, axis=1).tolist())

    # ── Compute COPQ per month ────────────────────────────────────────────────
    copq_rows = []
    for (yr, mo) in months:
        mdf = df[(df["year"]==yr) & (df["month"]==mo)]

        aust = mdf[mdf["entity"]=="AuST"]
        cp   = mdf[mdf["entity"]=="CenterPoint"]

        def good_cost(sub, entity):
            total = 0.0
            for _, row in sub.iterrows():
                p    = row["product"]
                cost = COSTS.get(p, COSTS["BLN"])
                rate = cost[0] if entity=="AuST" else cost[1]
                good = max(row["lot_size"] - row["retains"], 0)
                total += good * rate
            return total

        def scrap_cost(sub, entity):
            dc = DEFECT_COSTS_AUST if entity=="AuST" else DEFECT_COSTS_CP
            total = 0.0
            for _, row in sub.iterrows():
                for col in ALL_DEFECT_COLS:
                    if col == "Leak": continue  # use valve+bond instead
                    q = _sf(row.get(col, 0))
                    if q > 0:
                        total += q * dc.get(col, 16.43)
            return total

        aust_good  = good_cost(aust,  "AuST")
        cp_good    = good_cost(cp,    "CenterPoint")
        aust_scrap = scrap_cost(aust, "AuST")
        cp_scrap   = scrap_cost(cp,   "CenterPoint")

        tgc = aust_good + cp_good
        tbc = aust_scrap + cp_scrap
        good_cp_units  = sum(max(r["lot_size"]-r["retains"],0) for _,r in cp.iterrows())
        good_aust_units= sum(max(r["lot_size"]-r["retains"],0) for _,r in aust.iterrows())

        costed_yield  = tgc/(tgc+tbc)      if (tgc+tbc) > 0 else None
        copq_per_part = tbc/good_cp_units   if good_cp_units > 0 else None

        leak_aust = float(aust["Leak"].sum())
        leak_cp   = float(cp["Leak"].sum())
        lr_aust   = leak_aust/(good_aust_units+leak_aust) if (good_aust_units+leak_aust) > 0 else 0
        lr_cp     = leak_cp  /(good_cp_units  +leak_cp)   if (good_cp_units  +leak_cp)   > 0 else 0
        cumul_lr  = (leak_aust+leak_cp)/(good_cp_units+leak_cp) if (good_cp_units+leak_cp) > 0 else 0

        # Per-product yields
        prod_yields = {}
        for prod in ["SMG","BLN","CTI","BMD","MP3","NFS"]:
            prows = mdf[mdf["product"]==prod]
            if prows.empty:
                prod_yields[prod] = None
                continue
            pg = sum(good_cost(prows[prows["entity"]=="AuST"], "AuST") +
                     good_cost(prows[prows["entity"]=="CenterPoint"], "CenterPoint")
                     for _ in [1])
            pb = sum(scrap_cost(prows[prows["entity"]=="AuST"], "AuST") +
                     scrap_cost(prows[prows["entity"]=="CenterPoint"], "CenterPoint")
                     for _ in [1])
            prod_yields[prod] = pg/(pg+pb) if (pg+pb) > 0 else None

        copq_rows.append({
            "year":yr,"month":mo,
            "aust_good":aust_good,"aust_scrap":aust_scrap,
            "cp_good":cp_good,"cp_scrap":cp_scrap,
            "total_good":tgc,"total_scrap":tbc,
            "costed_yield":costed_yield,"copq_per_part":copq_per_part,
            "leak_aust":leak_aust,"leak_cp":leak_cp,
            "good_aust":good_aust_units,"good_cp":good_cp_units,
            "leak_rate_aust":lr_aust,"leak_rate_cp":lr_cp,"cumul_leak":cumul_lr,
            "bln_yield":prod_yields["BLN"],"smg_yield":prod_yields["SMG"],
            "cti_yield":prod_yields["CTI"],"bmd_yield":prod_yields["BMD"],
            "mp3_yield":prod_yields["MP3"],
        })

    df_copq = pd.DataFrame(copq_rows)
    return df, df_copq, months, errors


def get_top_defects(df_rows, year, month, top_n=10):
    """Top N defects by quantity for a given month."""
    mdf = df_rows[(df_rows["year"]==year) & (df_rows["month"]==month)]
    totals = {}
    for col in ALL_DEFECT_COLS:
        if col in EXCLUDE_FROM_TOP10: continue
        v = float(mdf[col].sum())
        if v > 0: totals[col] = v
    return sorted(totals.items(), key=lambda x:-x[1])[:top_n]


def get_leak_trend(df_rows, df_copq, n_months=6):
    """Last n months of leak rate data."""
    months = sorted(df_copq[["year","month"]].apply(tuple, axis=1).tolist())[-n_months:]
    rows = []
    for (yr, mo) in months:
        import calendar
        r = df_copq[(df_copq["year"]==yr) & (df_copq["month"]==mo)]
        if r.empty: continue
        rows.append({
            "label": f"{calendar.month_abbr[mo]} {yr}",
            "AuST": float(r["leak_rate_aust"].iloc[0]),
            "CenterPoint": float(r["leak_rate_cp"].iloc[0]),
        })
    return pd.DataFrame(rows)


def get_tip_trend(df_rows, df_copq, n_months=6):
    """Last n months of destroyed tip rate."""
    months = sorted(df_copq[["year","month"]].apply(tuple, axis=1).tolist())[-n_months:]
    rows = []
    for (yr, mo) in months:
        import calendar
        mdf = df_rows[(df_rows["year"]==yr) & (df_rows["month"]==mo)]
        for entity in ["AuST","CenterPoint"]:
            edf = mdf[mdf["entity"]==entity]
            tips = float(edf["Destroyed Tip"].sum())
            lots = float(edf["lot_size"].sum())
        aust_edf = mdf[mdf["entity"]=="AuST"]
        cp_edf   = mdf[mdf["entity"]=="CenterPoint"]
        aust_tips = float(aust_edf["Destroyed Tip"].sum())
        aust_lots = float(aust_edf["lot_size"].sum())
        cp_tips   = float(cp_edf["Destroyed Tip"].sum())
        cp_lots   = float(cp_edf["lot_size"].sum())
        rows.append({
            "label": f"{calendar.month_abbr[mo]} {yr}",
            "AuST":        aust_tips/(aust_tips+aust_lots) if (aust_tips+aust_lots)>0 else 0,
            "CenterPoint": cp_tips  /(cp_tips  +cp_lots)   if (cp_tips  +cp_lots)  >0 else 0,
        })
    return pd.DataFrame(rows)
