#!/usr/bin/env python3
"""
Build VOTERS_LOCAL/: an offline browser of the full voter list of each booth,
with the families found by up_family.py and a GenZ (age range) selector.

  ./build_voters_local.py 186_Amethi
  ./build_voters_local.py 186_Amethi -o VOTERS_LOCAL

Reads   <AC dir>_xlsx/AC*_Part*.xlsx        voters (up_extract.py)
        FAMILY/<AC dir>_Families.json       families (up_family.py)
        BOOTHNAMES/, CSV/                   booth names, village / panchayat
        site/data/names_hi.json             Hindi AC / district names (optional)
Writes  VOTERS_LOCAL/index.html             the browser (open it from disk)
        VOTERS_LOCAL/inc-logo.png
        VOTERS_LOCAL/data/acs.js            the ACs built so far
        VOTERS_LOCAL/data/<AC>/booths.js    per-booth counts and age histograms
        VOTERS_LOCAL/data/<AC>/pNNN.js      one booth's voters

VOTERS_LOCAL holds every voter's name, age, relation, house and EPIC: it is
for this machine (or a team's own copies) only -- never push it anywhere.
"""

from __future__ import annotations

import argparse
import csv
import glob
import json
import os
import re
import shutil
import sys

import up_family
from booth_names import load_booth_names

HERE = os.path.dirname(os.path.abspath(__file__))
RT = {"father": "F", "mother": "M", "husband": "H", "other": "O"}
SEX = {"Male": "M", "Female": "F", "Third Gender": "T", "Other": "T"}
AGE_MIN, AGE_MAX = 18, 110          # age histogram range


def js_write(path, prefix, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(prefix)
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
        fh.write(");\n")


def cover(ac_dir_name):
    path = os.path.join(HERE, "CSV", f"{ac_dir_name}.csv")
    out = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                if str(r.get("बूथ", "")).isdigit():
                    out[int(r["बूथ"])] = (r.get("मुख्य कस्बा/ग्राम", ""), r.get("पंचायत", ""))
    return out


def build(ac_dir, out_dir):
    base = os.path.basename(os.path.normpath(ac_dir))
    m = re.match(r"^(\d+)_(.+)$", base)
    if not m:
        sys.exit(f"{ac_dir}: expected <ACNo>_<Name>")
    ac_no, ac_name = int(m.group(1)), m.group(2).replace("_", " ")
    xlsx_dir = os.path.join(HERE, f"{base}_xlsx")
    fam_path = os.path.join(HERE, "FAMILY", f"{base}_Families.json")
    if not os.path.exists(fam_path):
        sys.exit(f"{fam_path} missing: run ./up_family.py {base}_xlsx first")
    with open(fam_path, encoding="utf-8") as fh:
        families = {int(k): v for k, v in json.load(fh).items()}
    names = load_booth_names(os.path.join(HERE, "BOOTHNAMES"), ac_no)
    places = cover(base)

    booths = []
    files = sorted(glob.glob(os.path.join(xlsx_dir, "*.xlsx")), key=up_family.part_no)
    for k, path in enumerate(files, 1):
        part = up_family.part_no(path)
        voters = up_family.read_part(path)
        fam_of, links_of = {}, {}
        for no, serials, links in families.get(part, []):
            for sl, ln in zip(serials, links):
                fam_of[sl], links_of[sl] = no, ln
        rows, live, deleted = [], 0, 0
        hist_m = [0] * (AGE_MAX - AGE_MIN + 1)
        hist_f = [0] * (AGE_MAX - AGE_MIN + 1)
        hist_t = [0] * (AGE_MAX - AGE_MIN + 1)
        for v in voters:
            sex = SEX.get(v["gender"], "")
            rows.append([v["sl"], v["epic"], v["name"], RT.get(v["rtype"], v["rtype"][:1].upper()),
                         v["rname"], v["house"], v["age"], sex, v["section"],
                         "" if v["list"] in ("", "Main") else v["list"].replace("Addition ", "A"),
                         v["mark"], fam_of.get(v["sl"], 0), links_of.get(v["sl"], [])])
            if v["deleted"]:
                deleted += 1
                continue
            live += 1
            if v["age"] is not None and AGE_MIN <= v["age"] <= AGE_MAX:
                h = hist_m if sex == "M" else hist_f if sex == "F" else hist_t
                h[v["age"] - AGE_MIN] += 1
        fams = families.get(part, [])
        # age histogram of voters in a family, for "GenZ in families"
        hist_fam = [0] * (AGE_MAX - AGE_MIN + 1)
        in_fam = set(fam_of)
        for v in voters:
            if not v["deleted"] and v["sl"] in in_fam and v["age"] is not None \
                    and AGE_MIN <= v["age"] <= AGE_MAX:
                hist_fam[v["age"] - AGE_MIN] += 1
        village, gp = places.get(part, ("", ""))
        booths.append([part, names.get(part, ""), village, gp, live, deleted, len(fams),
                       sum(len(f[1]) for f in fams), hist_m, hist_f, hist_t, hist_fam])
        js_write(os.path.join(out_dir, "data", str(ac_no), f"p{part:03d}.js"),
                 f"window.__booth({ac_no},{part},", rows)
        if k % 50 == 0:
            print(f"  {k}/{len(files)} booths", file=sys.stderr, flush=True)

    hi = {}
    hi_path = os.path.join(HERE, "site", "data", "names_hi.json")
    if os.path.exists(hi_path):
        with open(hi_path, encoding="utf-8") as fh:
            hi = json.load(fh)
    meta = {"no": ac_no, "name": ac_name, "nameHi": hi.get("acs", {}).get(str(ac_no), ""),
            "ageMin": AGE_MIN, "booths": booths}
    js_write(os.path.join(out_dir, "data", str(ac_no), "booths.js"), f"window.__ac({ac_no},", meta)

    # index of every AC built so far
    acs = []
    for path in sorted(glob.glob(os.path.join(out_dir, "data", "*", "booths.js"))):
        text = open(path, encoding="utf-8").read()
        d = json.loads(text[text.index(",") + 1: text.rindex(")")])
        acs.append([d["no"], d["name"], d["nameHi"], len(d["booths"]),
                    sum(b[4] for b in d["booths"])])
    js_write(os.path.join(out_dir, "data", "acs.js"), "window.__acs(", acs)
    shutil.copy(os.path.join(HERE, "voters_local.html"), os.path.join(out_dir, "index.html"))
    shutil.copy(os.path.join(HERE, "site", "assets", "inc-logo.png"), os.path.join(out_dir, "inc-logo.png"))
    print(f"{out_dir}: AC {ac_no} {ac_name}, {len(booths)} booths, "
          f"{sum(b[4] for b in booths)} voters -> {out_dir}/index.html")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[1])
    ap.add_argument("dirs", nargs="+", help="AC directories, e.g. 186_Amethi")
    ap.add_argument("-o", "--out", default=os.path.join(HERE, "VOTERS_LOCAL"))
    args = ap.parse_args(argv)
    for d in args.dirs:
        build(d, args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
