#!/usr/bin/env python3
"""保险产品知识库构建脚本: 读取 data/raw/ 三份调研原始数据,
规范化后生成 data/*.json 主数据、csv/*.csv 明细、Excel 总表。
用法: python3 scripts/build_kb.py
"""
import csv
import json
import datetime as dt
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
CATS = [
    ("critical_illness", "重疾险", "CI"),
    ("medical", "医疗险", "MED"),
    ("accident", "意外险", "ACC"),
]

COVERAGE_KEYS = {
    "critical_illness": [
        "severe_illness", "moderate_illness", "mild_illness",
        "special_extra", "death_benefit", "waiver", "highlights",
    ],
    "medical": [
        "general_medical_limit", "critical_medical_limit", "deductible",
        "reimbursement_ratio", "special_drug", "proton_therapy",
        "value_added", "highlights",
    ],
    "accident": [
        "accidental_death", "disability", "accidental_medical",
        "hospital_subsidy", "transport_extra", "sudden_death", "highlights",
    ],
}


def txt(v):
    if v is None:
        return "—"
    if isinstance(v, list):
        return "\n".join(str(x) for x in v)
    return str(v)


def num(v):
    return v if isinstance(v, (int, float)) else None


def load_and_normalize(slug, prefix):
    path = RAW / f"{slug}.json"
    items = json.loads(path.read_text(encoding="utf-8"))
    out = []
    for i, p in enumerate(items, 1):
        rec = dict(p)
        rec["product_id"] = f"{prefix}-{i:02d}"
        # 统一 sum_insured_options 到顶层(意外险原始数据在顶层,已是数组)
        cov_keys = COVERAGE_KEYS[slug]
        filled = sum(1 for k in cov_keys if rec.get("coverage", {}).get(k) not in (None, "", []))
        top_keys = ["sum_insured_options", "underwriting_age", "guarantee_period",
                    "waiting_period", "hesitation_period", "payment_period_options",
                    "renewal", "premium_reference"]
        filled += sum(1 for k in top_keys if rec.get(k) not in (None, "", []))
        pr = rec.get("premium_reference") or {}
        premium_filled = sum(1 for k in ("male_annual", "female_annual", "annual_premium") if pr.get(k))
        total = len(cov_keys) + len(top_keys) + 1
        rec["data_completeness"] = round((filled + (1 if premium_filled else 0)) / total, 2)
        out.append(rec)
    return out


def base_row(p):
    row = {
        "product_id": p["product_id"],
        "产品名称": p["product_name"],
        "承保公司": txt(p.get("company")),
        "渠道": txt(p.get("channel")),
        "在售状态": txt(p.get("status")),
        "投保年龄": txt(p.get("underwriting_age")),
        "保障期间": txt(p.get("guarantee_period")),
        "等待期": txt(p.get("waiting_period")),
        "犹豫期": txt(p.get("hesitation_period")),
        "置信度": p.get("confidence", ""),
        "数据抓取日": txt(p.get("retrieved_at")),
    }
    return row


def write_csv(path, rows, columns):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def build_xlsx(sheets, out_path, meta):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    head_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    head_fill = PatternFill("solid", start_color="2F5597")
    body_font = Font(name="微软雅黑", size=10)
    wrap = Alignment(wrap_text=True, vertical="top")

    for name, rows, columns in sheets:
        ws = wb.create_sheet(name)
        ws.append(columns)
        for r in rows:
            ws.append([r.get(c, "") for c in columns])
        for cell in ws[1]:
            cell.font, cell.fill, cell.alignment = head_font, head_fill, Alignment(
                wrap_text=True, vertical="center", horizontal="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font, cell.alignment = body_font, wrap
        for j, col in enumerate(columns, 1):
            width = 12
            if col in ("产品名称", "承保公司"):
                width = 34
            elif col in ("渠道", "在售状态", "投保年龄", "保障期间", "等待期", "犹豫期", "缴费期间"):
                width = 22
            elif "保费" in col or "限额" in col or "保额" in col:
                width = 14
            elif len(col) <= 6:
                width = 14
            else:
                width = 40
            ws.column_dimensions[get_column_letter(j)].width = width
        ws.freeze_panes = "B2"

    ws = wb.create_sheet("数据说明")
    notes = [
        ["保险产品知识库 · 数据说明", ""],
        ["", ""],
        ["数据抓取日期", meta["retrieved_at"]],
        ["最近复核日期", meta.get("last_verified", meta["retrieved_at"])],
        ["生成时间", meta["generated_at"]],
        ["产品总数", str(meta["total"])],
        ["重疾险/医疗险/意外险", f"{meta['CI']}/{meta['MED']}/{meta['ACC']} 款"],
        ["", ""],
        ["置信度定义", "high=官方条款/费率表/官网原文核实; medium=权威测评或平台页交叉验证; low=仅搜索摘要,建议复核"],
        ["缺失值", "— 表示该指标公开渠道未能核实到,非产品没有该责任;购买决策前请以官方条款为准"],
        ["保费口径", "重疾险默认 30岁/50万保额/30年缴/保终身; 医疗险默认 30岁有社保首年; 意外险为主销档位年缴。口径不同时以'保费口径'列文字为准"],
        ["机器可读数据", "data/products.json (全量) 及 data/{critical_illness,medical,accident}.json"],
        ["免责声明", "本表仅供研究参考,不构成投保建议;保险产品更新换代频繁,以保险公司官方披露的条款与费率为准"],
    ]
    for r in notes:
        ws.append(r)
    ws["A1"].font = Font(name="微软雅黑", bold=True, size=12)
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = wrap
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90

    wb.save(out_path)


def main():
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = {"generated_at": now, "retrieved_at": "2026-09-08"}
    all_products, sheets, counts = [], [], {}

    ci_rows, med_rows, acc_rows = [], [], []

    for slug, label, prefix in CATS:
        products = load_and_normalize(slug, prefix)
        counts[prefix] = len(products)
        (ROOT / "data" / f"{slug}.json").write_text(
            json.dumps(products, ensure_ascii=False, indent=2), encoding="utf-8")
        for p in products:
            rec = dict(p)
            rec["category"] = label
            all_products.append(rec)
            cov = p.get("coverage") or {}
            pr = p.get("premium_reference") or {}
            common = base_row(p)
            if slug == "critical_illness":
                ci_rows.append({**common,
                    "形态": txt(p.get("sub_type")),
                    "保额档位": txt(p.get("sum_insured_options")),
                    "缴费期间": txt(p.get("payment_period_options")),
                    "重疾责任": txt(cov.get("severe_illness")),
                    "中症责任": txt(cov.get("moderate_illness")),
                    "轻症责任": txt(cov.get("mild_illness")),
                    "特疾额外赔": txt(cov.get("special_extra")),
                    "身故责任": txt(cov.get("death_benefit")),
                    "豁免": txt(cov.get("waiver")),
                    "男性年保费(元)": num(pr.get("male_annual")) or "",
                    "女性年保费(元)": num(pr.get("female_annual")) or "",
                    "保费口径": txt(pr.get("assumption")),
                    "亮点与短板": txt(cov.get("highlights")),
                    "来源": "; ".join(p.get("source_urls") or [])})
            elif slug == "medical":
                med_rows.append({**common,
                    "形态": txt(p.get("sub_type")),
                    "续保条件": txt(p.get("renewal")),
                    "一般医疗限额(万)": cov.get("general_medical_limit") or "",
                    "重疾医疗限额(万)": cov.get("critical_medical_limit") or "",
                    "免赔额": txt(cov.get("deductible")),
                    "报销比例": txt(cov.get("reimbursement_ratio")),
                    "特药/外购药": txt(cov.get("special_drug")),
                    "质子重离子": txt(cov.get("proton_therapy")),
                    "增值服务": txt(cov.get("value_added")),
                    "年保费(元)": num(pr.get("annual_premium")) or "",
                    "保费口径": txt(pr.get("assumption")),
                    "亮点与短板": txt(cov.get("highlights")),
                    "来源": "; ".join(p.get("source_urls") or [])})
            else:
                acc_rows.append({**common,
                    "形态": txt(p.get("sub_type")),
                    "档位说明": txt(p.get("sum_insured_options")),
                    "意外身故": txt(cov.get("accidental_death")),
                    "意外伤残": txt(cov.get("disability")),
                    "意外医疗": txt(cov.get("accidental_medical")),
                    "住院津贴": txt(cov.get("hospital_subsidy")),
                    "交通意外额外赔": txt(cov.get("transport_extra")),
                    "猝死责任": txt(cov.get("sudden_death")),
                    "年保费(元)": num(pr.get("annual_premium")) or "",
                    "保费口径": txt(pr.get("assumption")),
                    "亮点与短板": txt(cov.get("highlights")),
                    "来源": "; ".join(p.get("source_urls") or [])})

    meta.update(total=len(all_products), **counts)
    meta["last_verified"] = max(
        (p.get("last_verified") or meta["retrieved_at"] for p in all_products),
        default=meta["retrieved_at"])
    (ROOT / "data" / "products.json").write_text(
        json.dumps({"meta": meta, "products": all_products}, ensure_ascii=False, indent=2),
        encoding="utf-8")

    ci_cols = list(ci_rows[0].keys()) if ci_rows else []
    med_cols = list(med_rows[0].keys()) if med_rows else []
    acc_cols = list(acc_rows[0].keys()) if acc_rows else []
    write_csv(ROOT / "csv" / "critical_illness.csv", ci_rows, ci_cols)
    write_csv(ROOT / "csv" / "medical.csv", med_rows, med_cols)
    write_csv(ROOT / "csv" / "accident.csv", acc_rows, acc_cols)

    build_xlsx([
        ("重疾险", ci_rows, ci_cols),
        ("医疗险", med_rows, med_cols),
        ("意外险", acc_rows, acc_cols),
    ], ROOT / "保险产品对比总表.xlsx", meta)

    print(json.dumps(meta, ensure_ascii=False))
    for slug, label, _ in CATS:
        prods = [p for p in all_products if p["category"] == label]
        lows = [p["product_id"] for p in prods if p.get("confidence") == "low"]
        no_prem = [p["product_id"] for p in prods
                   if not any((p.get("premium_reference") or {}).get(k)
                              for k in ("male_annual", "female_annual", "annual_premium"))]
        avg_c = round(sum(p["data_completeness"] for p in prods) / len(prods), 2)
        print(f"{label}: {len(prods)}款 平均完整度{avg_c} 无保费:{no_prem or '无'} 低置信:{lows or '无'}")


if __name__ == "__main__":
    main()
