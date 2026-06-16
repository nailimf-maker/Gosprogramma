#!/usr/bin/env python3
"""
Агрегатор реестра расселения аварийного жилья → компактный data.json.

Модель «фактов»: на клиент уходит по одной записи на помещение в КОЛОНОЧНОМ виде,
все категории закодированы целыми (словари — отдельно). Это позволяет считать любые
комбинации фильтров (район × причина признания × год признания) прямо в браузере
за один проход, без предрасчёта всех сочетаний. ~33k строк пересчитываются мгновенно.

ID-колонки (ID СФ/МР/МО дома/Дома/помещения) отбрасываются и в дашборд не попадают.
Персональные данные жильцов не выгружаются — только обезличенные признаки помещений.
"""
import json
import sys
import os
import pandas as pd

SRC = sys.argv[1] if len(sys.argv) > 1 else "Новая_версия2.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "data.json"

C_MR     = "Наименование МР"
C_ADDR   = "Адрес дома"
C_LIFE   = "Стадия жизненного цикла"
C_PREM   = "Тип помещения"
C_OWN    = "Тип собственности"
C_AREA   = "Общая площадь помещения, кв."
C_ROOMS  = "Количество комнат, ед."
C_STATUS = "Состояние расселения"
C_DONE   = "Фактическая дата заверешния расселения"
C_FAMILY = "Количество постоянно проживающих членов семьи, чел. "
C_REASON = "Причина признания дома аварийным"
C_EMERG  = "Дата признания дома аварийным"

DROP = ["ID СФ", "ID МР", "ID МО дома", "ID Дома", "ID помещения"]
STATUS_ORDER = ["Расселено", "Подлежит расселению", "Пустующие",
                "Розыск собственника", "В суде", "Не заполнено", "Не указано"]
ROOMS_LABELS = ["Студия/0", "1", "2", "3", "4", "5+"]

df = pd.read_excel(SRC, sheet_name="Помещения")
df = df.drop(columns=[c for c in DROP if c in df.columns]).dropna(how="all")

df[C_AREA]   = pd.to_numeric(df[C_AREA], errors="coerce")
df[C_FAMILY] = pd.to_numeric(df[C_FAMILY], errors="coerce")
df[C_ROOMS]  = pd.to_numeric(df[C_ROOMS], errors="coerce")
df[C_EMERG]  = pd.to_datetime(df[C_EMERG], errors="coerce")
df[C_DONE]   = pd.to_datetime(df[C_DONE], errors="coerce")
for c in [C_MR, C_STATUS, C_PREM, C_OWN, C_REASON, C_LIFE, C_ADDR]:
    df[c] = df[c].astype("string").fillna("Не указано").str.strip().replace("", "Не указано")

def dict_for(col, fixed_order=None):
    counts = df[col].value_counts()
    if fixed_order:
        labels = [x for x in fixed_order if x in counts.index]
        labels += [x for x in counts.index if x not in fixed_order]
    else:
        labels = list(counts.index)
    return labels

def codes(col, labels):
    idx = {lab: i for i, lab in enumerate(labels)}
    return df[col].map(idx).astype(int).tolist()

d_mr     = dict_for(C_MR)
d_status = dict_for(C_STATUS, STATUS_ORDER)
d_prem   = dict_for(C_PREM)
d_own    = dict_for(C_OWN)
d_life   = dict_for(C_LIFE)
d_reason = dict_for(C_REASON)
d_house  = dict_for(C_ADDR)

house_mr = {}
for addr, mr in df[[C_ADDR, C_MR]].drop_duplicates(C_ADDR).itertuples(index=False):
    house_mr[addr] = mr
mr_idx = {lab: i for i, lab in enumerate(d_mr)}
d_house_mr = [mr_idx[house_mr[a]] for a in d_house]

col_mr = codes(C_MR, d_mr); col_st = codes(C_STATUS, d_status)
col_pr = codes(C_PREM, d_prem); col_ow = codes(C_OWN, d_own)
col_lf = codes(C_LIFE, d_life); col_rs = codes(C_REASON, d_reason)
col_hs = codes(C_ADDR, d_house)

def rooms_bucket(v):
    if pd.isna(v): return -1
    v = int(v)
    if v <= 0: return 0
    if v >= 5: return 5
    return v
col_rm = [rooms_bucket(v) for v in df[C_ROOMS]]

col_ar10 = [int(round(v*10)) if pd.notna(v) else 0 for v in df[C_AREA]]
col_re   = [int(v) if pd.notna(v) else 0 for v in df[C_FAMILY]]
col_ey   = [int(v.year) if pd.notna(v) else 0 for v in df[C_EMERG]]
col_ry   = [int(v.year*100+v.month) if pd.notna(v) else 0 for v in df[C_DONE]]

emerg_years = sorted({y for y in col_ey if y >= 1997})

meta = {
    "period": str(df["Отчетный период"].dropna().iloc[0]) if "Отчетный период" in df.columns and df["Отчетный период"].notna().any() else "—",
    "region": str(df["Наименование СФ"].dropna().iloc[0]) if "Наименование СФ" in df.columns and df["Наименование СФ"].notna().any() else "—",
    "rows": len(df), "mr_count": len(d_mr),
    "year_min": emerg_years[0], "year_max": emerg_years[-1],
    "generated": pd.Timestamp.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
    "data_updated": str(df["Дата заполнения"].max())[:10] if "Дата заполнения" in df.columns else "—",
}

payload = {
    "meta": meta,
    "dict": {"mr": d_mr, "status": d_status, "premise": d_prem, "ownership": d_own,
             "lifecycle": d_life, "reason": d_reason, "rooms": ROOMS_LABELS, "house_mr": d_house_mr},
    "houses": d_house,
    "cols": {"mr": col_mr, "st": col_st, "pr": col_pr, "ow": col_ow, "lf": col_lf,
             "rs": col_rs, "rm": col_rm, "hs": col_hs,
             "ar10": col_ar10, "re": col_re, "ey": col_ey, "ry": col_ry},
}

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

print(f"OK -> {OUT} ({os.path.getsize(OUT)/1024:.0f} KB)")
print("rows:", len(df), "| районов:", len(d_mr), "| домов:", len(d_house),
      "| годы:", emerg_years[0], "-", emerg_years[-1])
print("причины:", d_reason)
