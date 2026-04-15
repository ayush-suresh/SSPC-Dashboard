"""
extractor.py  —  reads AuST and CP DOR .xlsm files.
Extracts: scrap, COPQ, daily production (day/swing), demand plan,
          attendance, and actions.
"""
import io
import calendar
from collections import defaultdict
import pandas as pd
from openpyxl import load_workbook

# ── Cost tables ───────────────────────────────────────────────────────────────
COSTS = {
    "BLN":  [17.93,27.08],
    "SMG":  [18.11,26.57],
    "CTI":  [17.93,27.08],
    "BMD":  [17.93,27.08],
    "MP3":  [9.86, 13.23],
    "NFS":  [17.93,27.08],
}
DEFECT_COST_AUST = {
    "Leak":16.43,"Leak - valve":16.43,"Leak - bond":16.43,
    "Scratch":16.43,"Dirty":16.43,"Wire at Tip":11.08,"Wire at Hub":13.38,
    "Flash ID":11.08,"Destroyed Tip":11.08,"OD Bulge":16.43,
    "Gap during Reflow":9.21,"Void at Hub Joint":13.38,
    "OD Flash at Hub Joint":11.08,"OD Tip Flash":11.08,
    "Side Exposed Wire":16.43,"Bubble":9.21,
    "Fiber/ Embedded Particulate":13.38,"Dent":13.38,"Skive":9.21,
    "Kink":16.43,"Burn":16.43,"Damaged/ Melted Hub":16.43,
    "Bad Cut Valve Visual":16.43,"Unknown Reflow Time":9.21,
    "Cut Short":16.43,"Pin Holes in Heat Shrink":9.21,
    "Failed Shape (Handle Wrong Direction/ doesn't conform to template)":16.43,
    "Extrusions in Wrong Order":9.21,"Tip Bleed":11.08,"Irregular Braid":1.82,
    "Hole at Tip":11.08,"Marker":16.43,"Wrong Valve Position":16.43,
    "Flash OD":11.08,"Stress Marks (Hub)":16.43,"Other":16.43,
}
DEFECT_COST_CP = {k: 22.08 for k in DEFECT_COST_AUST}
DEFECT_COST_CP.update({
    "Leak":22.51,"Leak - valve":22.51,"Leak - bond":22.51,
    "Unknown Reflow Time":21.94,"Extrusions in Wrong Order":21.94,
})
DEFECT_COST_MAP = {"AuST": DEFECT_COST_AUST, "CenterPoint": DEFECT_COST_CP}

# ── AuST scrap column config ──────────────────────────────────────────────────
AUST_PRODUCTS = {
    "SMG":{"svc":14,"ss":32},"BLN":{"svc":15,"ss":71},
    "CTI":{"svc":16,"ss":110},"BMD":{"svc":17,"ss":149},
    "MP3-S":{"svc":18,"ss":188},"MP3-M":{"svc":19,"ss":227},
    "NFE":{"svc":20,"ss":266},
}
AUST_REASONS = [
    "Bad Liner","Irregular Braid","Open Braid","Unknow Reflow Time","Scratch","Dirty",
    "Fiber/Emb. Part.","Skive","Kink","Burn","Destroyed Tip","Tip Bleed",
    "Valve Orientation","Cut Short","OD Tip Flash","Wire at Tip","Damaged/Melt Hub",
    "Pin Holes at HS","Hole at Tip","Marker","Flash OD","Leak - valve","Leak - bond",
    "Destructive Test","Gap","Layup Wrong Extrusion Order","Irregular Liner at Tip",
    "Stress Marks (Hub)","other",
]
AUST_MAP = {
    "Bad Liner":None,"Irregular Braid":"Irregular Braid","Open Braid":"Other",
    "Unknow Reflow Time":"Unknown Reflow Time","Scratch":"Scratch","Dirty":"Dirty",
    "Fiber/Emb. Part.":"Fiber/ Embedded Particulate","Skive":"Skive","Kink":"Kink",
    "Burn":"Burn","Destroyed Tip":"Destroyed Tip","Tip Bleed":"Tip Bleed",
    "Valve Orientation":"Wrong Valve Position","Cut Short":"Cut Short",
    "OD Tip Flash":None,"Wire at Tip":"Wire at Tip",
    "Damaged/Melt Hub":"Damaged/ Melted Hub","Pin Holes at HS":"Pin Holes in Heat Shrink",
    "Hole at Tip":"Hole at Tip","Marker":"Marker","Flash OD":"Flash OD",
    "Leak - valve":"Leak - valve","Leak - bond":"Leak - bond",
    "Destructive Test":"Other","Gap":"Gap during Reflow",
    "Layup Wrong Extrusion Order":"Extrusions in Wrong Order",
    "Irregular Liner at Tip":"Other","Stress Marks (Hub)":"Stress Marks (Hub)",
    "other":"Other",
}

# ── CP scrap column config ────────────────────────────────────────────────────
CP_PRODUCTS = {
    "SMG":{"svc":13,"ss":32},"BLN":{"svc":14,"ss":66},
    "CTI":{"svc":15,"ss":100},"BMD":{"svc":16,"ss":134},
    "MP3-S":{"svc":17,"ss":168},"MP3-M":{"svc":18,"ss":202},
    "NFE":{"svc":19,"ss":236},
}
CP_REASONS = [
    "Leak - Bond","Leak - Valve","Scratch","Dirty","Wire At Tip","Flash ID",
    "Destroyed Tip","Od Tip Flash","Fiber/Embedded Particulate","Skive","Kink","Burn",
    "Damage/Melted Hub","Unknown Reflow Time","Cut Short","Pin Holes in Heat Shrink",
    "Failed Shape (Handle Wrong Direction/doest conform to template)","Tip Bleed",
    "Irregular Bread","Hole at Tip","Marker","Bad Liner (Splotchy coating)",
    "Valve Deformed/Dammage/dirty","other",
]
CP_MAP = {
    "Leak - Bond":"Leak - bond","Leak - Valve":"Leak - valve",
    "Scratch":"Scratch","Dirty":"Dirty","Wire At Tip":"Wire at Tip",
    "Flash ID":None,"Destroyed Tip":"Destroyed Tip","Od Tip Flash":None,
    "Fiber/Embedded Particulate":"Fiber/ Embedded Particulate",
    "Skive":"Skive","Kink":"Kink","Burn":"Burn",
    "Damage/Melted Hub":"Damaged/ Melted Hub",
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
    "Bad Cut Valve Visual","Unknown Reflow Time","Cut Short","Pin Holes in Heat Shrink",
    "Failed Shape (Handle Wrong Direction/ doesn't conform to template)",
    "Extrusions in Wrong Order","Tip Bleed","Irregular Braid","Hole at Tip","Marker",
    "Wrong Valve Position","Flash OD","Stress Marks (Hub)","Other",
]
EXCLUDE_FROM_TOP10 = {"Other","Unknown Issue Samples in Retains",
                       "Bad Liner (splotchy coating, etc)","Leak"}


def _sf(v):
    if v is None: return 0.0
    try:
        f = float(v)
        return 0.0 if f != f else f
    except: return 0.0


# ══════════════════════════════════════════════════════════════════════════════
# SCRAP EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════
def _extract_scrap(ws, qty_col, lot_col, date_col, products, reasons, mapping, entity):
    agg = defaultdict(lambda: {"qty":0,"retains":0,"lot":"",
                                **{c:0 for c in ALL_DEFECT_COLS}})
    for r in range(4, ws.max_row+1):
        date = ws.cell(r, date_col).value
        if not hasattr(date,"month"): continue
        rt = ws.cell(r, qty_col).value
        if rt not in ("Qty","Reject"): continue
        lot_raw = str(ws.cell(r, lot_col).value or "").strip()
        tokens  = [t for t in lot_raw.split() if t.startswith(("ML","CL","ml","cl"))]
        lot_key = tokens[0] if tokens else (lot_raw.split()[0] if lot_raw else "")
        for prod, cfg in products.items():
            svc = ws.cell(r, cfg["svc"]).value
            if not svc: continue
            try: count = int(float(svc))
            except: continue
            if count <= 0: continue
            pc  = "MP3" if prod in ("MP3-S","MP3-M") else ("NFS" if prod=="NFE" else prod)
            key = ((date.year, date.month), pc, entity)
            if rt == "Qty":
                agg[key]["qty"] += count
                if lot_key and not agg[key]["lot"]: agg[key]["lot"] = lot_key
                ss = cfg["ss"]
                for i, reason in enumerate(reasons):
                    dest = mapping.get(reason)
                    if dest is None: continue
                    v = ws.cell(r, ss+i).value
                    if v:
                        try: n = int(float(v))
                        except: n = 0
                        if n > 0: agg[key][dest] += n
            elif rt == "Reject":
                agg[key]["retains"] += count
    rows = []
    for (mk, pc, ent), data in agg.items():
        qty = data["qty"]; ret = data["retains"]
        if qty==0 and ret==0 and not any(data[c]>0 for c in ALL_DEFECT_COLS): continue
        data["Leak"] = data.get("Leak - valve",0) + data.get("Leak - bond",0)
        yld = round((qty-ret)/qty,4) if qty>0 else None
        row = {"year":mk[0],"month":mk[1],"product":pc,"entity":ent,
               "lot":data["lot"] or "","lot_size":float(qty),"retains":float(ret),
               "yield":yld}
        for c in ALL_DEFECT_COLS: row[c] = data.get(c,0)
        rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# DAILY PRODUCTION (day/swing split)
# ══════════════════════════════════════════════════════════════════════════════
def _extract_daily_production(ws_aust, ws_cp):
    rows = []

    def parse(ws, facility, qty_col, date_col, shift_col,
              prod_cols, total_col, people_col, hours_col, peoples_col, notes_col):
        for r in range(4, ws.max_row+1):
            date  = ws.cell(r, date_col).value
            if not hasattr(date,"month"): continue
            rt    = ws.cell(r, qty_col).value
            shift = ws.cell(r, shift_col).value
            if rt not in ("Qty","Reject"): continue
            if not shift: continue
            for prod, col in prod_cols.items():
                val = _sf(ws.cell(r, col).value)
                if val == 0: continue
                pc = "MP3" if prod in ("MP3-S","MP3-M") else ("NFS" if prod=="NFE" else prod)
                rows.append({
                    "date":date,"year":date.year,"month":date.month,"day":date.day,
                    "facility":facility,"shift":shift,"product":pc,"row_type":rt,
                    "value":val,
                    "people": _sf(ws.cell(r, people_col).value),
                    "hours":  _sf(ws.cell(r, hours_col).value),
                    "peoples":str(ws.cell(r, peoples_col).value or ""),
                    "notes":  str(ws.cell(r, notes_col).value or ""),
                })

    parse(ws_aust,"AuST",13,2,6,
          {"SMG":14,"BLN":15,"CTI":16,"BMD":17,"MP3-S":18,"MP3-M":19,"NFE":20},
          21,297,296,299,300)
    parse(ws_cp,"CenterPoint",12,2,6,
          {"SMG":13,"BLN":14,"CTI":15,"BMD":16,"MP3-S":17,"MP3-M":18,"NFE":19},
          21,263,264,265,266)

    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows)
    qty_df = df[df["row_type"]=="Qty"][
        ["date","year","month","day","facility","shift","product",
         "value","people","hours","peoples","notes"]
    ].rename(columns={"value":"qty"})
    rej_df = df[df["row_type"]=="Reject"][
        ["date","facility","shift","product","value"]
    ].rename(columns={"value":"rejects"})
    merged = qty_df.merge(rej_df, on=["date","facility","shift","product"], how="left")
    merged["rejects"]     = merged["rejects"].fillna(0)
    merged["good_units"]  = merged["qty"] - merged["rejects"]
    merged["reject_rate"] = merged.apply(
        lambda r: r["rejects"]/r["qty"] if r["qty"]>0 else 0, axis=1)
    return merged.sort_values(["date","facility","shift","product"]).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# DEMAND / PLAN
# ══════════════════════════════════════════════════════════════════════════════
def _extract_demand(ws):
    months = {}
    for c in range(5, ws.max_column+1):
        v = ws.cell(3, c).value
        if hasattr(v,"month"): months[c] = v
    rows = []
    # AuST plan rows 16-24, CP plan rows 29-37
    for facility, row_range in [("AuST", range(16,25)),("CenterPoint", range(29,38))]:
        for r in row_range:
            prod = ws.cell(r, 3).value
            if not prod or not isinstance(prod, str): continue
            for col, dt in months.items():
                val = _sf(ws.cell(r, col).value)
                rows.append({"year":dt.year,"month":dt.month,"product":prod,
                              "facility":facility,"plan":val})
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# ATTENDANCE
# ══════════════════════════════════════════════════════════════════════════════
def _extract_attendance(ws):
    dates = {}
    for c in range(2, ws.max_column+1):
        v = ws.cell(3, c).value
        if hasattr(v,"month"): dates[c] = v
    STATUS = {"a":"On Time","i":"Late","r":"No Show"}
    rows = []
    for r in range(4, 15):
        team = ws.cell(r, 1).value
        if not team or not isinstance(team, str): continue
        for col, dt in dates.items():
            raw = ws.cell(r, col).value
            if raw is None: continue
            if isinstance(raw, str):
                status = STATUS.get(raw.lower().strip(), raw)
            elif isinstance(raw, (int,float)):
                status = f"{raw:.0%}" if 0 < raw <= 1 else str(int(raw))
            else: continue
            rows.append({"date":dt,"year":dt.year,"month":dt.month,
                          "day":dt.day,"team":team,"status":status})
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# ACTIONS
# ══════════════════════════════════════════════════════════════════════════════
def _extract_actions(ws):
    rows = []
    for r in range(2, ws.max_row+1):
        date     = ws.cell(r,1).value
        category = ws.cell(r,2).value
        urgency  = ws.cell(r,3).value
        desc     = ws.cell(r,4).value
        task     = ws.cell(r,5).value
        assigned = ws.cell(r,6).value
        due      = ws.cell(r,7).value
        status   = ws.cell(r,8).value
        if date is None and category is None: continue
        if isinstance(date, str):
            try: date = pd.to_datetime(date, errors="coerce")
            except: date = None
        rows.append({
            "date":date,
            "year":  date.year  if hasattr(date,"year")  else None,
            "month": date.month if hasattr(date,"month") else None,
            "category":    str(category or "").strip(),
            "urgency":     str(urgency  or "").strip(),
            "description": str(desc     or "").strip(),
            "task":        str(task     or "").strip(),
            "assigned_to": str(assigned or "").strip(),
            "due_date":    due,
            "status":      str(status   or "Open").strip(),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# COPQ COMPUTATION
# ══════════════════════════════════════════════════════════════════════════════
def _compute_copq(df_scrap):
    months = sorted(df_scrap[["year","month"]].drop_duplicates().apply(tuple,axis=1).tolist())
    copq_rows = []
    for (yr, mo) in months:
        mdf  = df_scrap[(df_scrap["year"]==yr)&(df_scrap["month"]==mo)]
        aust = mdf[mdf["entity"]=="AuST"]
        cp   = mdf[mdf["entity"]=="CenterPoint"]

        def good_cost(sub, entity):
            total = 0.0
            for _,row in sub.iterrows():
                rate = COSTS.get(row["product"], COSTS["BLN"])[0 if entity=="AuST" else 1]
                total += max(row["lot_size"]-row["retains"],0)*rate
            return total

        def scrap_cost(sub, entity):
            dc = DEFECT_COST_MAP[entity]; total = 0.0
            for _,row in sub.iterrows():
                for col in ALL_DEFECT_COLS:
                    if col == "Leak": continue
                    q = _sf(row.get(col,0))
                    if q > 0: total += q*dc.get(col,16.43)
            return total

        def py(sub_all, prod):
            p = sub_all[sub_all["product"]==prod]
            if p.empty: return None
            pg = good_cost(p[p["entity"]=="AuST"],"AuST") + \
                 good_cost(p[p["entity"]=="CenterPoint"],"CenterPoint")
            pb = scrap_cost(p[p["entity"]=="AuST"],"AuST") + \
                 scrap_cost(p[p["entity"]=="CenterPoint"],"CenterPoint")
            return pg/(pg+pb) if (pg+pb)>0 else None

        ag=good_cost(aust,"AuST");   cg=good_cost(cp,"CenterPoint")
        ab=scrap_cost(aust,"AuST");  cb=scrap_cost(cp,"CenterPoint")
        tgc=ag+cg; tbc=ab+cb
        good_cp   = sum(max(r["lot_size"]-r["retains"],0) for _,r in cp.iterrows())
        good_aust = sum(max(r["lot_size"]-r["retains"],0) for _,r in aust.iterrows())
        leak_aust = float(aust["Leak"].sum())
        leak_cp   = float(cp["Leak"].sum())
        lr_a = leak_aust/(good_aust+leak_aust) if (good_aust+leak_aust)>0 else 0
        lr_c = leak_cp/(good_cp+leak_cp)       if (good_cp+leak_cp)>0    else 0
        cumul= (leak_aust+leak_cp)/(good_cp+leak_cp) if (good_cp+leak_cp)>0 else 0

        copq_rows.append({
            "year":yr,"month":mo,
            "aust_good":ag,"aust_scrap":ab,"cp_good":cg,"cp_scrap":cb,
            "total_good":tgc,"total_scrap":tbc,
            "costed_yield":tgc/(tgc+tbc) if (tgc+tbc)>0 else None,
            "copq_per_part":tbc/good_cp  if good_cp>0   else None,
            "leak_aust":leak_aust,"leak_cp":leak_cp,
            "good_aust":good_aust,"good_cp":good_cp,
            "leak_rate_aust":lr_a,"leak_rate_cp":lr_c,"cumul_leak":cumul,
            "bln_yield":py(mdf,"BLN"),"smg_yield":py(mdf,"SMG"),
            "cti_yield":py(mdf,"CTI"),"bmd_yield":py(mdf,"BMD"),
            "mp3_yield":py(mdf,"MP3"),
        })
    return pd.DataFrame(copq_rows)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def extract_from_files(aust_bytes, cp_bytes):
    errors = []
    try:
        wb_aust = load_workbook(io.BytesIO(aust_bytes), data_only=True, keep_vba=True)
    except Exception as e:
        errors.append(f"Could not read AuST file: {e}"); return {}, errors
    try:
        wb_cp = load_workbook(io.BytesIO(cp_bytes), data_only=True, keep_vba=True)
    except Exception as e:
        errors.append(f"Could not read CP file: {e}"); return {}, errors

    ws_aust_data = wb_aust["AuST-SSPC DATA"]
    ws_cp_data   = wb_cp["CP-SSPC DATA"]

    aust_scrap = _extract_scrap(ws_aust_data,13,22,2,AUST_PRODUCTS,AUST_REASONS,AUST_MAP,"AuST")
    cp_scrap   = _extract_scrap(ws_cp_data,  12,22,2,CP_PRODUCTS,  CP_REASONS,  CP_MAP,  "CenterPoint")
    df_scrap   = pd.DataFrame(aust_scrap + cp_scrap)

    df_copq   = _compute_copq(df_scrap) if not df_scrap.empty else pd.DataFrame()
    df_prod   = _extract_daily_production(ws_aust_data, ws_cp_data)
    df_demand = _extract_demand(wb_aust["Demand Data"]) if "Demand Data"  in wb_aust.sheetnames else pd.DataFrame()
    df_att    = _extract_attendance(wb_aust["Attendance"]) if "Attendance" in wb_aust.sheetnames else pd.DataFrame()
    df_actions= _extract_actions(wb_aust["Actions"]) if "Actions" in wb_aust.sheetnames else pd.DataFrame()

    months = sorted(df_scrap[["year","month"]].drop_duplicates()
                    .apply(tuple,axis=1).tolist()) if not df_scrap.empty else []

    return {"scrap":df_scrap,"copq":df_copq,"prod":df_prod,"demand":df_demand,
            "att":df_att,"actions":df_actions,"months":months}, errors


# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def get_top_defects(df_scrap, year, month, top_n=10):
    mdf = df_scrap[(df_scrap["year"]==year)&(df_scrap["month"]==month)]
    totals = {}
    for col in ALL_DEFECT_COLS:
        if col in EXCLUDE_FROM_TOP10: continue
        v = float(mdf[col].sum())
        if v > 0: totals[col] = v
    total_all = sum(totals.values())
    return [(k,v,v/total_all if total_all>0 else 0)
            for k,v in sorted(totals.items(),key=lambda x:-x[1])[:top_n]]


def get_rolling_stats(df_copq, col, n=6):
    df  = df_copq.sort_values(["year","month"]).tail(n).copy()
    df["label"] = df.apply(lambda r: f"{calendar.month_abbr[int(r.month)]} {int(r.year)}",axis=1)
    vals = df[col].dropna().tolist()
    avg  = sum(vals)/len(vals) if vals else 0
    std  = pd.Series(vals).std() if len(vals)>1 else 0
    return df, avg, std


def get_leak_trend(df_scrap, df_copq, n_months=6):
    months = sorted(df_copq[["year","month"]].apply(tuple,axis=1).tolist())[-n_months:]
    rows = []
    for (yr,mo) in months:
        r = df_copq[(df_copq["year"]==yr)&(df_copq["month"]==mo)]
        if r.empty: continue
        rows.append({"label":f"{calendar.month_abbr[mo]} {yr}",
                     "AuST":float(r["leak_rate_aust"].iloc[0]),
                     "CenterPoint":float(r["leak_rate_cp"].iloc[0])})
    return pd.DataFrame(rows)


def get_tip_trend(df_scrap, df_copq, n_months=6):
    months = sorted(df_copq[["year","month"]].apply(tuple,axis=1).tolist())[-n_months:]
    rows = []
    for (yr,mo) in months:
        mdf  = df_scrap[(df_scrap["year"]==yr)&(df_scrap["month"]==mo)]
        aust = mdf[mdf["entity"]=="AuST"]; cp = mdf[mdf["entity"]=="CenterPoint"]
        at=float(aust["Destroyed Tip"].sum()); al=float(aust["lot_size"].sum())
        ct=float(cp["Destroyed Tip"].sum());   cl=float(cp["lot_size"].sum())
        rows.append({"label":f"{calendar.month_abbr[mo]} {yr}",
                     "AuST":at/(at+al) if (at+al)>0 else 0,
                     "CenterPoint":ct/(ct+cl) if (ct+cl)>0 else 0})
    return pd.DataFrame(rows)
