# -*- coding: utf-8 -*-
"""
Parse Factiva-like RTF export to CSV with columns:
标题, 作者, 时间, 语言, 正文, IPD, 关联话题, 关联地区, 发行公司, 文件号, 关联公司, 关联行业, 文章字数

Usage:
  python rtf2csv_factiva.py --input Factiva-20260125-1237.rtf --output out.csv
"""

import re
import csv
import argparse
import datetime
from pathlib import Path
import glob

# ---------- RTF -> TEXT ----------

def rtf_to_text(rtf: str) -> str:
    s = rtf.replace("\r\n", "\n").replace("\r", "\n")

    # decode unicode escapes: \uNNNN?
    def uni_repl(m):
        n = int(m.group(1))
        if n < 0:
            n += 65536
        try:
            return chr(n)
        except Exception:
            return ""

    s = re.sub(r"\\u(-?\d+)\??", uni_repl, s)

    # decode hex escapes: \'hh
    def hex_repl(m):
        b = bytes.fromhex(m.group(1))
        return b.decode("cp1252", errors="ignore")

    s = re.sub(r"\\'([0-9a-fA-F]{2})", hex_repl, s)

    # line breaks
    s = s.replace(r"\par", "\n").replace(r"\line", "\n")

    # remove control words
    s = re.sub(r"\\[a-zA-Z]+\d* ?", "", s)
    s = re.sub(r"\\[^a-zA-Z]", "", s)

    # remove braces
    s = s.replace("{", "").replace("}", "")

    # normalize newlines
    s = re.sub(r"\n{3,}", "\n\n", s)
    s = re.sub(r"[ \t]+\n", "\n", s)
    return s.strip()

def clean_factiva_text(t: str) -> str:
    # drop huge hex/binary-looking lines
    out_lines = []
    for line in t.splitlines():
        l = line.strip()
        if not l:
            out_lines.append("")
            continue
        hex_chars = sum(c in "0123456789ABCDEFabcdef" for c in l)
        ratio = hex_chars / max(1, len(l))
        if len(l) > 120 and ratio > 0.75:
            continue
        l2 = "".join(c if c.isprintable() else " " for c in l)
        out_lines.append(l2)
    cleaned = "\n".join(out_lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()

def preprocess(t: str) -> str:
    t = t.replace("\\'", "'")
    t = re.sub(r'HYPERLINK\s+toc\d+"', "\n", t)
    t = re.sub(r"\bd Page\b.*", "", t)
    t = re.sub(r"\bPAGE\d+\b", "", t)
    t = re.sub(r"\bNUMPAGES\d+\b", "", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n[ \t]+", "\n", t)
    return t.strip()

# ---------- Parsing helpers ----------

WORDCOUNT_RE = re.compile(r"^\s*([\d,]+)\s*字")
FILE_ID_PATTERN = re.compile(r"\b[A-Z]{2,10}\d{8,}[A-Za-z0-9]{3,}\b")

def parse_factiva_date(line: str):
    # tolerant: find yyyy, month, day, optional hh:mm even with junk chars between
    s = line.strip()
    m = re.search(r"(\d{4}).{0,10}?(\d{1,2}).{0,10}?(\d{1,2}).{0,10}?日(?:.{0,10}?(\d{1,2}):(\d{2}))?", s)
    if not m:
        return None
    y, mo, d, hh, mm = m.groups()
    return datetime.datetime(int(y), int(mo), int(d), int(hh or 0), int(mm or 0))

def detect_language(line: str) -> str:
    l = line.strip()
    if "英文" in l or ("英" in l and "文" in l) or re.search(r"\bEnglish\b", l, re.I):
        return "英文"
    if "中文" in l or ("中" in l and "文" in l) or re.search(r"\bChinese\b", l, re.I):
        return "中文"
    return ""

def clean_title(title: str) -> str:
    t = re.sub(r"^(?:toc\d+)+", "", title).strip()
    t = re.sub(r"^\W+", "", t).strip()
    return t

def body_sanitize(body: str) -> str:
    lines = []
    for ln in body.splitlines():
        l = ln.strip()
        if not l:
            lines.append("")
            continue
        hex_chars = sum(c in "0123456789ABCDEFabcdef" for c in l)
        if len(l) > 80 and hex_chars / len(l) > 0.8:
            break
        lines.append(ln.rstrip())
    out = "\n".join(lines).strip()
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out

def extract_file_id(text: str) -> str:
    # prefer near "文件" even if乱码：只要附近有形如 PRN00000... 的串
    m = re.search(r"文件\D{0,15}(" + FILE_ID_PATTERN.pattern + r")", text)
    if m:
        return m.group(1)
    allm = FILE_ID_PATTERN.findall(text)
    return allm[-1] if allm else ""

GEO = {
    "Asia","China","Hong Kong","Macau","Macao","Taiwan","United States","U.S.","US","UK","Europe",
    "Beijing","Shanghai","Japan","Korea","South Korea","North Korea","Germany","France","Netherlands",
    "Ireland","Switzerland","Latin America","Australia","Canada","Singapore","India","Russia","Africa","Middle East"
}
INDUSTRY_HINTS = [
    "Finance","Investment","Materials","Technology","Energy","Healthcare","Pharmaceutical",
    "Mining","Metals","Oil","Gas","Shipping","Logistics","Automotive","Real Estate","Telecom","AI","Chip"
]

COMPANY_RE = re.compile(
    r"\b([A-Z][A-Za-z0-9&.\-]*(?:\s+[A-Z][A-Za-z0-9&.\-]*){0,6}\s+"
    r"(?:Inc\.|Corp\.|Corporation|Ltd\.|Limited|Group|Company|Co\.|PLC|LLC|S\.A\.|AG))\b"
)

def split_keywords(kw: str):
    parts = [p.strip() for p in kw.split(",") if p.strip()]
    regions, topics, industries = [], [], []
    for p in parts:
        if p in GEO:
            regions.append(p)
        elif any(h.lower() in p.lower() for h in INDUSTRY_HINTS):
            industries.append(p)
        else:
            topics.append(p)
    return topics, regions, industries

def split_articles(text: str):
    lines = [ln.rstrip() for ln in text.splitlines()]
    idxs = [i for i, l in enumerate(lines) if WORDCOUNT_RE.search(l)]
    articles = []

    for wc_i in idxs:
        # title is nearest non-empty above wordcount
        j = wc_i - 1
        while j >= 0 and not lines[j].strip():
            j -= 1
        if j < 0:
            continue
        title = lines[j].strip()

        # date must exist right after wordcount (skip blanks)
        p = wc_i + 1
        while p < len(lines) and not lines[p].strip():
            p += 1
        dt = parse_factiva_date(lines[p]) if p < len(lines) else None
        if not dt:
            continue  # filter out non-article parts

        # author lines between title and wordcount (some articles have none)
        author_lines = [lines[k].strip() for k in range(j + 1, wc_i) if lines[k].strip()]
        author = " | ".join(author_lines) if author_lines else ""

        wc = int(WORDCOUNT_RE.search(lines[wc_i]).group(1).replace(",", ""))
        time_iso = dt.isoformat(sep=" ")

        # publisher line after date
        p += 1
        while p < len(lines) and not lines[p].strip():
            p += 1
        publisher = lines[p].strip() if p < len(lines) else ""
        p += 1

        # IPD code often next line (e.g., PRN/INVWK/AMM...)
        while p < len(lines) and not lines[p].strip():
            p += 1
        ipd_candidate = lines[p].strip() if p < len(lines) else ""
        ipd = ipd_candidate if re.fullmatch(r"[A-Z0-9]{2,10}", ipd_candidate) else ""
        if ipd:
            p += 1

        # find language within next lines
        lang = ""
        q = p
        for _ in range(25):
            if q >= len(lines):
                break
            lang = detect_language(lines[q])
            if lang:
                break
            q += 1

        # body starts after first blank line following the metadata block
        s = q
        while s < len(lines) and lines[s].strip():
            s += 1
        while s < len(lines) and not lines[s].strip():
            s += 1
        body_start = s if s < len(lines) else None

        articles.append({
            "title": title,
            "author": author,
            "time": time_iso,
            "language": lang,
            "publisher": publisher,
            "ipd": ipd,
            "word_count": wc,
            "body_start_idx": body_start,
            "title_line_idx": j
        })

    # attach bodies and extract keywords/file id
    for i, a in enumerate(articles):
        start = a["body_start_idx"]
        end = len(lines) if i + 1 == len(articles) else articles[i + 1]["title_line_idx"]
        body_text = "\n".join(lines[start:end]).strip() if start is not None else ""
        body_text = body_sanitize(body_text)

        # Keywords
        km = re.search(r"Keywords for this news article include:\s*(.+)", body_text)
        keywords = km.group(1).strip().rstrip(".") if km else ""

        # File id (robust)
        file_no = extract_file_id(body_text)

        # If author missing, try extract "By ..."
        if not a["author"]:
            m = re.search(r"--\s*By\s+([^-\n]+)", body_text)
            if m:
                a["author"] = m.group(1).strip()

        a["keywords"] = keywords
        a["file_no"] = file_no
        a["body"] = body_text

    return articles

def rtf_factiva_to_rows(rtf_path: Path):
    raw = rtf_path.read_text(errors="ignore")
    txt = rtf_to_text(raw)

    # keep from "Factiva RTF Display Format" onward if exists
    start_idx = txt.find("Factiva RTF Display Format")
    if start_idx != -1:
        txt = txt[start_idx:]

    txt = clean_factiva_text(txt)
    txt = preprocess(txt)

    articles = split_articles(txt)
    rows = []

    for a in articles:
        title = clean_title(a["title"])
        topics, regions, industries = split_keywords(a.get("keywords", ""))

        # companies from title + first 30 lines of body
        sample_text = title + "\n" + "\n".join(a["body"].splitlines()[:30])
        companies = sorted(set(m.group(1) for m in COMPANY_RE.finditer(sample_text)))

        rows.append({
            "标题": title,
            "作者": a.get("author", ""),
            "时间": a.get("time", ""),
            "语言": a.get("language", ""),
            "正文": a.get("body", ""),
            "IPD": a.get("ipd", ""),
            "关联话题": "; ".join(topics),
            "关联地区": "; ".join(regions),
            "发行公司": a.get("publisher", ""),
            "文件号": a.get("file_no", ""),
            "关联公司": "; ".join(companies),
            "关联行业": "; ".join(industries),
            "文章字数": a.get("word_count", ""),
        })

    return rows


# ---------- Main ----------

def main():
    parser = argparse.ArgumentParser(description='Convert Factiva RTF files to CSV')
    parser.add_argument('-i', '--input', help='Input RTF file or directory containing RTF files', required=True)
    parser.add_argument('-o', '--output', help='Output CSV file or directory (default: same as input with .csv extension)', required=False)
    parser.add_argument('-m', '--merge', help='Merge all results into a single CSV file', action='store_true', required=False)
    args = parser.parse_args()
    
    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None
    
    # Get all RTF files
    if input_path.is_dir():
        rtf_files = list(input_path.glob('*.rtf'))
        rtf_files.extend(input_path.glob('*.RTF'))  # Also check for uppercase extensions
    else:
        rtf_files = [input_path] if input_path.suffix.lower() == '.rtf' else []
    
    if not rtf_files:
        print(f"No RTF files found at {input_path}")
        return
    
    print(f"Found {len(rtf_files)} RTF file(s) to process")
    
    # Process files
    all_rows = []
    cols = ["标题","作者","时间","语言","正文","IPD","关联话题","关联地区","发行公司","文件号","关联公司","关联行业","文章字数"]
    
    for rtf_file in rtf_files:
        print(f"Processing {rtf_file}...")
        try:
            rows = rtf_factiva_to_rows(rtf_file)
            
            if args.merge:
                all_rows.extend(rows)
            else:
                # Generate output path for this file
                if output_path and output_path.is_dir():
                    out_file = output_path / f"{rtf_file.stem}.csv"
                elif output_path:
                    out_file = output_path
                else:
                    out_file = rtf_file.with_suffix('.csv')
                
                with out_file.open("w", newline="", encoding="utf-8-sig") as f:
                    w = csv.DictWriter(f, fieldnames=cols)
                    w.writeheader()
                    w.writerows(rows)
                
                print(f"  OK: wrote {len(rows)} rows -> {out_file}")
                
        except Exception as e:
            print(f"  Error processing {rtf_file}: {e}")
    
    # Handle merged output
    if args.merge:
        if output_path:
            merged_file = output_path
        else:
            merged_file = input_path.parent / f"merged_output.csv"
        
        with merged_file.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            w.writerows(all_rows)
        
        print(f"\nOK: merged {len(all_rows)} rows from {len(rtf_files)} files -> {merged_file}")


if __name__ == "__main__":
    main()
    # python factivartf2xlsx.py -i input.rtf //处理单个文件
    # python factivartf2xlsx.py -i rtf_files/ //处理目录中所有RTF文件：
    # python factivartf2xlsx.py -i rtf_files/ -o merged.csv -m // 处理目录并合并到单个文件：
    # python factivartf2xlsx.py -i rtf_files/ -o output_csv/ //处理目录并保存到指定输出目录：