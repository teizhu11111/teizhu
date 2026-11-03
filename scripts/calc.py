#!/usr/bin/env python3
"""
SKU 折扣价计算器

用法:
    python3 calc.py <输入.xlsx> [输出.xlsx]

规则:
    头程费用 = 包装长宽高均>0 ? L*W*H/1e6*1800
             : 预估长宽高均>0 ? L*W*H/1e6*1800
             : 沙特-头程运费>0 ? 沙特-头程运费
             : 阿联酋-头程运费>0 ? 阿联酋-头程运费
             : 7
    采购费用 = 采购成本 in (0, 1, 空) ? 10 : 采购成本
    折扣价   = (头程费用 + 采购费用) / 2

实现要点见 SKILL.md「硬性技术约束」：
- <f> 内不写前导 =
- 不加 ROUND，存全精度，显示靠 numFmt 0.00
- 数值比较一律包 N()
- 行正则容忍 xml:space="preserve"
"""
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

# 每次运行按第 1 行表头解析实际列位；ERP 导出版本会让折扣价格在 AY/AZ 间移动。
FIELDS = ("SKU", "预估长", "预估宽", "预估高", "包装长", "包装宽", "包装高",
          "采购成本", "沙特-头程运费", "阿联酋-头程运费", "折扣价格")
COL = {}


def resolve_columns(df):
    missing = [name for name in FIELDS if name not in df.columns]
    duplicate = [name for name in FIELDS if sum(col == name for col in df.columns) != 1 and name not in missing]
    if missing or duplicate:
        raise ValueError("表头异常，缺失=%s，重复=%s；未写入文件" % (missing, duplicate))
    return {**{name: get_column_letter(df.columns.get_loc(name) + 1) for name in FIELDS},
            "目标列": get_column_letter(df.columns.get_loc("折扣价格") + 1)}


HEADER_ROW = 1
SHEET = "sheet1"
RATE = 1800          # 体积重换算系数
FALLBACK_HEAD = 7.0  # 头程兜底
FALLBACK_PUR = 10.0  # 采购兜底


# ---------------------------------------------------------------- 计算
def compute(df):
    def nz(name):
        return pd.to_numeric(df[name], errors="coerce").fillna(0.0)

    L, W, H = nz("包装长"), nz("包装宽"), nz("包装高")
    eL, eW, eH = nz("预估长"), nz("预估宽"), nz("预估高")
    SA, AE, PC = nz("沙特-头程运费"), nz("阿联酋-头程运费"), nz("采购成本")

    use_pack = (L > 0) & (W > 0) & (H > 0)
    use_est = (~use_pack) & (eL > 0) & (eW > 0) & (eH > 0)
    use_sa = (~use_pack) & (~use_est) & (SA > 0)
    use_ae = (~use_pack) & (~use_est) & (SA <= 0) & (AE > 0)
    use_def = (~use_pack) & (~use_est) & (SA <= 0) & (AE <= 0)

    head = np.select(
        [use_pack, use_est, use_sa, use_ae],
        [L * W * H / 1000000 * RATE, eL * eW * eH / 1000000 * RATE, SA, AE],
        default=FALLBACK_HEAD,
    )
    pur = np.where((PC == 0) | (PC == 1), FALLBACK_PUR, PC)
    return (head + pur) / 2, {
        "包装尺寸": int(use_pack.sum()),
        "预估尺寸": int(use_est.sum()),
        "沙特头程": int(use_sa.sum()),
        "阿联酋头程": int(use_ae.sum()),
        f"兜底{FALLBACK_HEAD:g}": int(use_def.sum()),
    }, int(((PC == 0) | (PC == 1)).sum())


# ---------------------------------------------------------------- 公式
def formula(r):
    """注意：<f> 内不含前导 '='；数值比较一律用 N() 包裹"""
    return (
        "(IF(AND(N({K}{r})>0,N({W}{r})>0,N({H}{r})>0),"
        "{K}{r}*{W}{r}*{H}{r}/1000000*{RATE:g},"
        "IF(AND(N({eK}{r})>0,N({eW}{r})>0,N({eH}{r})>0),"
        "{eK}{r}*{eW}{r}*{eH}{r}/1000000*{RATE:g},"
        "IF(N({SA}{r})>0,N({SA}{r}),IF(N({AE}{r})>0,N({AE}{r}),{FALLBACK_HEAD:g})))))"
        "+IF(OR(N({PC}{r})=0,N({PC}{r})=1),{FALLBACK_PUR:g},N({PC}{r})))/2)"
    ).format(
        r=r, RATE=RATE, FALLBACK_HEAD=FALLBACK_HEAD, FALLBACK_PUR=FALLBACK_PUR,
        K=COL["包装长"], W=COL["包装宽"], H=COL["包装高"],
        eK=COL["预估长"], eW=COL["预估宽"], eH=COL["预估高"],
        SA=COL["沙特-头程运费"], AE=COL["阿联酋-头程运费"], PC=COL["采购成本"],
    )


# ---------------------------------------------------------------- 校验
def verify(src_path, out_path, expected, sheet_member):
    """独立解析路径复核：直接从源 XML 读单元格重算，与输出缓存值比对"""
    def parse(path):
        with zipfile.ZipFile(path) as archive:
            xml = archive.read(sheet_member).decode("utf-8")
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                shared = ["".join(node.itertext()) for node in strings.findall("{*}si")]
        data = {}
        for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
            r = int(rm.group(1))
            cells = {}
            for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', rm.group(2), re.S):
                vm = re.search(r"<v>(.*?)</v>", cm.group(4) or "", re.S)
                raw = vm.group(1) if vm else ""
                kind = re.search(r'\bt="([^"]+)"', cm.group(3))
                if kind and kind.group(1) == "s" and raw.strip():
                    raw = shared[int(raw)]
                elif kind and kind.group(1) == "inlineStr":
                    raw = "".join(ET.fromstring(cm.group(0)).itertext())
                try:
                    cells[cm.group(1)] = float(raw) if raw.strip() != "" else 0.0
                except ValueError:
                    cells[cm.group(1)] = 0.0
            data[r] = cells
        return data

    src = parse(src_path)
    with zipfile.ZipFile(out_path) as archive:
        out_xml = archive.read(sheet_member).decode("utf-8")
    tgt = COL["目标列"]
    cached = {
        int(m.group(1)): float(m.group(3))
        for m in re.finditer(r'<c r="%s(\d+)"[^>]*><f>(.*?)</f><v>([^<]+)</v></c>' % tgt, out_xml, re.S)
    }

    def calc(c):
        L, W, H = c.get(COL["包装长"], 0), c.get(COL["包装宽"], 0), c.get(COL["包装高"], 0)
        eL, eW, eH = c.get(COL["预估长"], 0), c.get(COL["预估宽"], 0), c.get(COL["预估高"], 0)
        SA, AE, PC = c.get(COL["沙特-头程运费"], 0), c.get(COL["阿联酋-头程运费"], 0), c.get(COL["采购成本"], 0)
        if L > 0 and W > 0 and H > 0:
            head, tag = L * W * H / 1000000 * RATE, "包装尺寸"
        elif eL > 0 and eW > 0 and eH > 0:
            head, tag = eL * eW * eH / 1000000 * RATE, "预估尺寸"
        elif SA > 0:
            head, tag = SA, "沙特头程"
        elif AE > 0:
            head, tag = AE, "阿联酋头程"
        else:
            head, tag = FALLBACK_HEAD, f"兜底{FALLBACK_HEAD:g}"
        pur = FALLBACK_PUR if PC in (0.0, 1.0) else PC
        return (head + pur) / 2, head, pur, tag

    bad, samples = [], {}
    for r in range(HEADER_ROW + 1, HEADER_ROW + 1 + len(expected)):
        if r not in src:
            bad.append((r, "行缺失", None))
            continue
        exp, head, pur, tag = calc(src[r])
        got = cached.get(r)
        if got is None or abs(got - exp) > 1e-9 or abs(float(expected[r - HEADER_ROW - 1]) - exp) > 1e-9:
            bad.append((r, exp, got))
        samples.setdefault(tag, []).append((r, head, pur, cached.get(r)))

    print("\n--- 独立复核 ---")
    print("缓存值数量: %d" % len(cached))
    print("不一致行数: %d" % len(bad))
    for b in bad[:10]:
        print("   ", b)
    print("--- 分来源抽查 ---")
    for tag, rows in samples.items():
        for r, head, pur, val in rows[:2]:
            print("行%-5d 头程=%8.2f 采购=%7.2f 折扣价=%7s  [%s]"
                  % (r, head, pur, ("%.2f" % val) if val is not None else "NA", tag))
    print("结论:", "全部一致" if not bad else "存在不一致")
    return not bad


# ---------------------------------------------------------------- 主流程
def main():
    if len(sys.argv) < 2:
        sys.exit("用法: python3 calc.py <输入.xlsx> [输出.xlsx]")
    src = os.path.abspath(sys.argv[1])
    if len(sys.argv) >= 3:
        out = os.path.abspath(sys.argv[2])
    else:
        base, ext = os.path.splitext(src)
        out = base + "_折扣价" + ext

    if os.path.samefile(src, out) if os.path.exists(out) else src == out:
        raise ValueError("输入和输出路径相同；不会覆盖源文件")
    if os.path.exists(out):
        raise FileExistsError("输出文件已存在，不会覆盖：%s" % out)
    df = pd.read_excel(src, sheet_name=SHEET, header=HEADER_ROW - 1)
    global COL
    COL = resolve_columns(df)
    if df.empty or df["SKU"].isna().any() or df["SKU"].astype(str).str.strip().eq("").any():
        raise ValueError("SKU 列为空或存在空 SKU；未写入文件")
    result, dist, pur10 = compute(df)
    print("折扣价格列: %s" % COL["目标列"])
    print("SKU 行数: %d" % len(df))
    print("头程来源分布: " + " ".join("%s=%d" % kv for kv in dist.items()))
    print("采购费用取 %g 的行数: %d (采购成本为 0 或 1)" % (FALLBACK_PUR, pur10))
    print("折扣价 min/mean/max: %.2f / %.2f / %.2f" % (result.min(), result.mean(), result.max()))

    work = tempfile.mkdtemp(prefix="discount_")
    try:
        with zipfile.ZipFile(src) as z:
            z.extractall(work)

        # styles.xml 确保存在 0.00 数字格式，并取其 xf 索引
        sty_path = os.path.join(work, "xl/styles.xml")
        sty = open(sty_path, encoding="utf-8").read()
        xfs_block = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', sty, re.S)
        xfs = re.findall(r'<xf [^>]*/>', xfs_block.group(2))
        idx = next((i for i, x in enumerate(xfs) if re.search(r'numFmtId="2"', x)), None)
        if idx is None:
            new_xf = '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            sty = (sty[:xfs_block.start()]
                   + '<cellXfs count="%d">%s%s</cellXfs>' % (len(xfs) + 1, xfs_block.group(2), new_xf)
                   + sty[xfs_block.end():])
            open(sty_path, "w", encoding="utf-8").write(sty)
            idx = len(xfs)
        style_2dp = idx

        # 按工作簿关系定位目标工作表，兼容相对路径与 /xl/ 开头的路径。
        workbook = ET.parse(os.path.join(work, "xl/workbook.xml")).getroot()
        sheet = next((s for s in workbook.findall("{*}sheets/{*}sheet")
                      if s.get("name") == SHEET), None)
        if sheet is None:
            raise ValueError("工作表不存在：%s" % SHEET)
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        relationships = ET.parse(os.path.join(work, "xl/_rels/workbook.xml.rels")).getroot()
        relation = next((rel for rel in relationships if rel.get("Id") == rid), None)
        if relation is None:
            raise ValueError("工作表关系缺失：%s" % rid)
        target = relation.get("Target")
        sheet_member = posixpath.normpath(target.lstrip("/") if target.startswith("/")
                                          else posixpath.join("xl", target))
        if not sheet_member.startswith("xl/") or sheet_member.startswith("xl/../"):
            raise ValueError("非法工作表路径：%s" % target)
        sheet_path = os.path.join(work, *sheet_member.split("/"))

        xml = open(sheet_path, encoding="utf-8").read()
        tgt = COL["目标列"]
        n = 0
        for i in range(len(df)):
            r = i + HEADER_ROW + 1
            cell = '<c r="%s%d" s="%d"><f>%s</f><v>%s</v></c>' % (
                tgt, r, style_2dp, formula(r), repr(float(result[i])))
            pat = re.compile(r'<c r="%s%d"(?:\s[^>]*)?/>|<c r="%s%d"(?:\s[^>]*)?>.*?</c>' % (tgt, r, tgt, r))
            xml, cnt = pat.subn(lambda _: cell, xml, count=1)
            n += cnt
        assert n == len(df), "仅匹配到 %d/%d 行，请检查列位映射" % (n, len(df))
        open(sheet_path, "w", encoding="utf-8").write(xml)
        print("写入单元格数: %d" % n)

        order = ["[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
                 "xl/workbook.xml", "xl/_rels/workbook.xml.rels", sheet_member,
                 "xl/styles.xml", "xl/theme/theme1.xml"]
        candidate = os.path.join(work, "verified.xlsx")
        with zipfile.ZipFile(candidate, "w", zipfile.ZIP_DEFLATED) as z:
            for name in order:
                p = os.path.join(work, name)
                if os.path.exists(p):
                    z.write(p, name)
            for root, _, files in os.walk(work):
                for f in files:
                    full = os.path.join(root, f)
                    rel = os.path.relpath(full, work)
                    if rel not in order and not rel.endswith((".py", ".md", ".xlsx")):
                        z.write(full, rel)
        if not verify(src, candidate, result, sheet_member):
            raise ValueError("独立复核失败，未生成输出文件")
        # 仅允许目标单元格和必要的数字格式变动，保留其他工作簿成员原样。
        with zipfile.ZipFile(src) as original, zipfile.ZipFile(candidate) as generated:
            src_root = ET.fromstring(original.read(sheet_member))
            out_root = ET.fromstring(generated.read(sheet_member))
            def other_cells(root):
                return {cell.get("r"): ET.tostring(cell)
                        for cell in root.findall(".//{*}sheetData/{*}row/{*}c")
                        if not cell.get("r", "").startswith(COL["目标列"])}
            if other_cells(src_root) != other_cells(out_root):
                raise ValueError("非目标单元格发生变化，未生成输出文件")
            for member in original.namelist():
                if member.endswith("/"):
                    continue
                if member not in (sheet_member, "xl/styles.xml") and original.read(member) != generated.read(member):
                    raise ValueError("非目标工作簿成员发生变化：%s" % member)
        if os.path.exists(out):
            raise FileExistsError("输出文件已存在，不会覆盖：%s" % out)
        shutil.copyfile(candidate, out)
        print("输出: %s" % out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
