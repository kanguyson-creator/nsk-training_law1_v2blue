#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
nsk_localize_law_pdfs.py
------------------------
ดาวน์โหลด PDF กฎหมายจากลิงก์เว็บราชการ (cad.go.th, cpd.go.th, soc.go.th ฯลฯ)
มาเก็บไว้ในโฟลเดอร์ pdf/ ของ repo เว็บคลังกฎหมาย แล้วแก้ลิงก์ในไฟล์ .html/.js/.json/.md
ให้ชี้ไปที่ไฟล์ในเว็บเราเอง (ลิงก์จะไม่เสียอีก)

วิธีใช้ (รันในโฟลเดอร์ repo nsk-training_law1_v2blue บนเครื่องบอส):
    pip install requests beautifulsoup4
    python nsk_localize_law_pdfs.py . --dry-run     # ดูก่อนว่าเจอลิงก์อะไรบ้าง
    python nsk_localize_law_pdfs.py .               # ดาวน์โหลด + แก้ลิงก์จริง
    git diff --stat                                 # ตรวจสิ่งที่เปลี่ยน

ผลลัพธ์:
    pdf/                      ไฟล์ PDF ที่ดาวน์โหลดมา
    pdf/index.json            ทะเบียนไฟล์: ลิงก์ต้นทาง -> ไฟล์ในเครื่อง, ขนาด, SHA-256, วันที่ดาวน์โหลด
    law_link_report.csv       รายงานทุกลิงก์ (OK / FAILED) เปิดด้วย Excel ได้

รันซ้ำได้: ลิงก์ที่แก้แล้วจะไม่ถูกดาวน์โหลดซ้ำ จะลองใหม่เฉพาะลิงก์ที่ยัง FAILED
"""
import argparse
import csv
import hashlib
import html
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

DEFAULT_DOMAINS = ["cad.go.th", "cpd.go.th", "soc.go.th"]
TEXT_EXT = {".html", ".htm", ".js", ".json", ".md"}
SKIP_DIRS = {".git", "node_modules", "pdf", ".github"}
URL_RE = re.compile(r"""https?://[^\s"'<>`)\]]+""")
TRAIL = ".,;:"
MAX_WARN_MB = 50  # GitHub จำกัดไฟล์ละ 100 MB
PDF_DIR = "pdf"
JUNK_RE = re.compile(r"user_?manual|qa_tod|นโยบาย|ปฏิเสธความรับผิด", re.I)


def host_match(url, domains):
    host = (urlparse(url).hostname or "").lower()
    return any(host == d or host.endswith("." + d) for d in domains)


def clean_match(raw):
    cleaned = raw.rstrip(TRAIL)
    return cleaned, raw[len(cleaned):]


def iter_text_files(root):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in filenames:
            p = Path(dirpath) / fn
            if p.suffix.lower() in TEXT_EXT:
                yield p


def find_pdf_candidates(page_url, text):
    soup = BeautifulSoup(text, "html.parser")
    found = []
    for tag, attr in (("a", "href"), ("iframe", "src"), ("embed", "src"),
                      ("object", "data"), ("source", "src")):
        for t in soup.find_all(tag):
            v = t.get(attr)
            if not v:
                continue
            full = urljoin(page_url, v)
            # กรณีเป็น PDF viewer เช่น viewer.html?file=xxx.pdf
            for vals in parse_qs(urlparse(full).query).values():
                for q in vals:
                    if ".pdf" in q.lower():
                        found.append(urljoin(page_url, q))
            if ".pdf" in unquote(urlparse(full).path).lower():
                found.append(full)
    uniq = []
    for f in found:
        # ตัด PDF ที่เป็นส่วนประกอบของเว็บ (คู่มือผู้ใช้, นโยบายเว็บไซต์, FAQ ท้ายหน้า) ออก
        if JUNK_RE.search(unquote(urlparse(f).path)):
            continue
        if f not in uniq:
            uniq.append(f)
    # ให้ไฟล์ในโฟลเดอร์ /uploads/ (เอกสารจริงที่อัปโหลด) มาก่อน
    uniq.sort(key=lambda u: 0 if "/uploads/" in urlparse(u).path else 1)
    return uniq


def fetch_pdf(sess, url, depth=0):
    """คืนค่า (content, final_url, note)"""
    r = sess.get(url, timeout=60, allow_redirects=True)
    r.raise_for_status()
    ctype = r.headers.get("Content-Type", "").lower()
    if "pdf" in ctype or r.content[:5] == b"%PDF-":
        return r.content, r.url, ""
    if depth >= 1:
        raise ValueError("ปลายทางไม่ใช่ไฟล์ PDF")
    # ส่งเป็น bytes ให้ BeautifulSoup ตรวจ charset เอง (กันภาษาไทยเพี้ยน)
    cands = find_pdf_candidates(r.url, r.content)
    if not cands:
        raise ValueError("เป็นหน้าเว็บ แต่ไม่พบไฟล์ PDF แนบ (อาจโหลดด้วย JavaScript)")
    content, final_url, _ = fetch_pdf(sess, cands[0], depth + 1)
    note = "ดึง PDF จากหน้าเว็บ"
    if len(cands) > 1:
        note += f" (หน้านี้มี PDF {len(cands)} ไฟล์ ใช้ไฟล์แรก — ควรตรวจ)"
    return content, final_url, note


def make_name(orig_url, final_url, content):
    base = Path(unquote(urlparse(final_url).path)).name
    if not base.lower().endswith(".pdf"):
        base = Path(unquote(urlparse(orig_url).path).rstrip("/")).name + ".pdf"
    stem = base[:-4]
    stem = re.sub(r'[<>:"/\\|?*#%&\x00-\x1f]', "_", stem)
    stem = re.sub(r"\s+", "_", stem).strip(" ._") or "doc"
    stem = stem[:80]
    h = hashlib.sha256(content).hexdigest()[:8]
    return f"{stem}-{h}.pdf"


def main():
    ap = argparse.ArgumentParser(description="ดาวน์โหลด PDF กฎหมายมาเก็บใน repo และแก้ลิงก์")
    ap.add_argument("repo", nargs="?", default=".", help="โฟลเดอร์ repo เว็บคลังกฎหมาย")
    ap.add_argument("--domains", nargs="+", default=DEFAULT_DOMAINS,
                    help="โดเมนที่จะดึงไฟล์มาเก็บ (default: cad.go.th cpd.go.th soc.go.th)")
    ap.add_argument("--dry-run", action="store_true", help="แสดงรายการลิงก์เท่านั้น ไม่ดาวน์โหลด/ไม่แก้ไฟล์")
    ap.add_argument("--delay", type=float, default=1.0, help="หน่วงเวลาระหว่างดาวน์โหลด (วินาที)")
    args = ap.parse_args()

    root = Path(args.repo).resolve()
    pdf_dir = root / PDF_DIR
    index_path = pdf_dir / "index.json"
    domains = [d.lower() for d in args.domains]

    # 1) สแกนหาลิงก์
    url_files = {}  # url (unescaped) -> set ของไฟล์ที่พบ
    for f in iter_text_files(root):
        text = f.read_text(encoding="utf-8", errors="replace")
        for m in URL_RE.finditer(text):
            cleaned, _ = clean_match(m.group(0))
            url = html.unescape(cleaned)
            if host_match(url, domains):
                url_files.setdefault(url, set()).add(str(f.relative_to(root)))

    print(f"พบลิงก์ภายนอกที่ต้องจัดการ {len(url_files)} ลิงก์ (ไม่ซ้ำ)")
    if args.dry_run:
        for u, fs in sorted(url_files.items()):
            print(f"  {u}\n      ใน: {', '.join(sorted(fs))}")
        return

    # 2) ดาวน์โหลด
    pdf_dir.mkdir(exist_ok=True)
    index = {}
    if index_path.exists():
        index = json.loads(index_path.read_text(encoding="utf-8"))

    sess = requests.Session()
    sess.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36")
    rows = []
    mapping = {}  # url -> path ใน repo (เช่น pdf/xxx.pdf)
    total = len(url_files)
    for i, (url, files) in enumerate(sorted(url_files.items()), 1):
        found_in = "; ".join(sorted(files))
        if url in index and (root / index[url]["file"]).exists():
            mapping[url] = index[url]["file"]
            rows.append(["OK (มีแล้ว)", url, index[url]["file"], found_in, ""])
            continue
        last_err = ""
        for attempt in range(2):
            try:
                content, final_url, note = fetch_pdf(sess, url)
                name = make_name(url, final_url, content)
                rel = f"{PDF_DIR}/{name}"
                (root / rel).write_bytes(content)
                size_mb = len(content) / 1024 / 1024
                if size_mb > MAX_WARN_MB:
                    note = (note + f" | ไฟล์ใหญ่ {size_mb:.1f} MB").strip(" |")
                index[url] = {
                    "file": rel,
                    "final_url": final_url,
                    "bytes": len(content),
                    "sha256": hashlib.sha256(content).hexdigest(),
                    "downloaded_at": datetime.now().isoformat(timespec="seconds"),
                }
                mapping[url] = rel
                rows.append(["OK", url, rel, found_in, note])
                print(f"[{i}/{total}] OK     {rel}")
                last_err = ""
                break
            except Exception as e:  # noqa: BLE001
                last_err = f"{type(e).__name__}: {e}"
                time.sleep(2)
        if last_err:
            rows.append(["FAILED", url, "", found_in, last_err])
            print(f"[{i}/{total}] FAILED {url}\n        {last_err}")
        time.sleep(args.delay)

    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")

    # 3) แก้ลิงก์ในไฟล์ (เปลี่ยนเฉพาะ URL ที่ดาวน์โหลดสำเร็จ รูปแบบไฟล์เดิมคงอยู่)
    changed = 0
    for f in iter_text_files(root):
        text = f.read_text(encoding="utf-8", errors="replace")
        # ไฟล์ .html อ้างอิงจากโฟลเดอร์ของตัวเอง, .js/.json/.md อ้างอิงจาก root ของเว็บ
        base_dir = f.parent if f.suffix.lower() in {".html", ".htm"} else root

        def repl(m):
            cleaned, tail = clean_match(m.group(0))
            url = html.unescape(cleaned)
            if url in mapping:
                return os.path.relpath(root / mapping[url], start=base_dir).replace("\\", "/") + tail
            return m.group(0)

        new_text = URL_RE.sub(repl, text)
        if new_text != text:
            f.write_text(new_text, encoding="utf-8")
            changed += 1

    # 4) รายงาน
    report = root / "law_link_report.csv"
    with report.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["สถานะ", "ลิงก์ต้นทาง", "ไฟล์ในเว็บเรา", "พบในไฟล์", "หมายเหตุ"])
        w.writerows(rows)

    ok = sum(1 for r in rows if r[0].startswith("OK"))
    print("\n=== สรุป ===")
    print(f"สำเร็จ {ok} / {total} ลิงก์ | ล้มเหลว {total - ok} ลิงก์")
    print(f"แก้ไขไฟล์ {changed} ไฟล์")
    print(f"รายงาน: {report.name}  (ลิงก์ FAILED ต้องหาไฟล์ฉบับใหม่ด้วยมือ)")


if __name__ == "__main__":
    sys.exit(main())
