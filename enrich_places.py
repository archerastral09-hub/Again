#!/usr/bin/env python3
"""Enrichit agences_prospection.xlsx via Searlo GET /search/places.

Une ligne source = une requête (q = nom, location = adresse si dispo, limit = 5).
On garde le résultat dont nom (+ adresse si dispo) est identique après normalisation ; sinon la ligne est ignorée (raison tracée dans l'onglet "Ignorées").
Plusieurs clés API : rotation (la moins récemment utilisée d'abord), cooldown sur
429, clé retirée du pool sur 401/402/403.
"""
import argparse
import json
import os
import re
import sys
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = os.environ.get("SEARLO_BASE_URL", "https://api.searlo.tech/api/v1")
# Limite Searlo Maps/Places : 240 req/min par compte -> marge de sécurité
RATE_PER_MIN = int(os.environ.get("SEARLO_RATE_PER_MIN", "220"))
LIMIT = 5
MAX_ATTEMPTS = 8
CHUNK = 200
FIELDS = ["title", "address", "rating", "ratingCount", "category",
          "phoneNumber", "website", "domain"]

# ----------------------------------------------------------------- normalisation
LEGAL = {"sarl", "sas", "sasu", "eurl", "sa", "sci"}
ABBR = {
    "av": "avenue", "ave": "avenue", "bd": "boulevard", "bld": "boulevard",
    "blvd": "boulevard", "chem": "chemin", "pl": "place", "imp": "impasse",
    "all": "allee", "rte": "route", "fg": "faubourg", "fbg": "faubourg",
    "st": "saint", "ste": "sainte", "crs": "cours", "sq": "square",
    "zi": "zone industrielle", "za": "zone artisanale",
}


def _tokens(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = s.replace("&", " et ").replace("œ", "oe")
    return re.findall(r"[a-z0-9]+", s)


def norm_name(s):
    return " ".join(t for t in _tokens(s) if t not in LEGAL)


def norm_addr(s, title=None):
    """Adresse normalisée. Google préfixe souvent l'adresse avec le nom de la
    fiche ("100% Immobilier, 37 Pl. du Marché, ...") : ce préfixe est retiré."""
    toks = _tokens(s)
    if title:
        t = _tokens(title)
        if t and toks[:len(t)] == t:
            toks = toks[len(t):]
    out = []
    for t in toks:
        if t == "france":
            continue
        out.extend(ABBR.get(t, t).split())
    return " ".join(out)


def pick(name, addr, places):
    """Retourne (place, raison). place=None => ligne ignorée."""
    if not places:
        return None, "aucun_resultat"
    if len(places) == 1:  # un seul résultat : gardé sans vérification
        return places[0], "ok"
    n = norm_name(name)
    cands = [p for p in places if norm_name(p.get("title")) == n]
    if not cands:
        return None, "nom_different"
    if addr:
        a = norm_addr(addr)
        cands = [p for p in cands if norm_addr(p.get("address"), p.get("title")) == a]
        if not cands:
            return None, "adresse_differente"
    uniq = {(p.get("placeId") or p.get("cid") or id(p)): p for p in cands}
    if len(uniq) != 1:
        return None, "ambigu"
    return next(iter(uniq.values())), "ok"


# ---------------------------------------------------------------- rotation de clés
class AllKeysDead(Exception):
    pass


class ApiError(Exception):
    pass


class KeyPool:
    def __init__(self, keys, per_min):
        self.lock = threading.Lock()
        self.interval = 60.0 / per_min
        self.keys = {k: {"next": 0.0, "dead": None} for k in keys}

    @staticmethod
    def mask(k):
        return "…" + k[-4:]

    def acquire(self):
        with self.lock:
            alive = [(v["next"], k) for k, v in self.keys.items() if not v["dead"]]
            if not alive:
                return None
            t, k = min(alive)
            now = time.monotonic()
            slot = max(t, now)
            self.keys[k]["next"] = slot + self.interval
        if slot > now:
            time.sleep(slot - now)
        return k

    def cooldown(self, k, sec):
        with self.lock:
            self.keys[k]["next"] = max(self.keys[k]["next"], time.monotonic() + sec)

    def kill(self, k, status):
        with self.lock:
            if not self.keys[k]["dead"]:
                self.keys[k]["dead"] = status
                print(f"[clé {self.mask(k)}] retirée du pool (HTTP {status})", flush=True)

    def n_alive(self):
        with self.lock:
            return sum(1 for v in self.keys.values() if not v["dead"])


def search(session, pool, name, addr):
    params = {"q": name, "limit": LIMIT, "gl": "fr", "hl": "fr"}
    if addr:
        params["location"] = addr[:200]
    last = "erreur_api"
    for attempt in range(MAX_ATTEMPTS):
        key = pool.acquire()
        if key is None:
            raise AllKeysDead()
        try:
            r = session.get(f"{BASE_URL}/search/places", params=params,
                            headers={"x-api-key": key}, timeout=45)
        except requests.RequestException as e:
            last = f"reseau:{type(e).__name__}"
            time.sleep(min(2 ** attempt, 30))
            continue
        s = r.status_code
        if s == 200:
            return r.json().get("places") or []
        if s == 429:
            try:
                wait = float(r.headers.get("Retry-After"))
            except (TypeError, ValueError):
                wait = min(10 * (attempt + 1), 60)
            pool.cooldown(key, wait)
            last = "429"
            continue
        if s in (401, 402, 403):
            pool.kill(key, s)
            continue
        if s in (500, 502, 503, 504):
            last = str(s)
            time.sleep(min(2 ** attempt, 30))
            continue
        raise ApiError(f"HTTP {s}: {r.text[:120]}")
    raise ApiError(last)


# --------------------------------------------------------------------- E/S fichiers
def read_rows(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Agences"] if "Agences" in wb.sheetnames else wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    header = [str(h or "").strip().lower() for h in next(it)]

    def col(key, default):
        for i, h in enumerate(header):
            if key in h:
                return i
        return default

    i_name, i_addr = col("nom", 0), col("adresse", 1)
    rows = []
    for n, r in enumerate(it, start=2):
        name = str(r[i_name]).strip() if r[i_name] is not None else ""
        addr = str(r[i_addr]).strip() if len(r) > i_addr and r[i_addr] is not None else ""
        if name or addr:
            rows.append((n, name, addr))
    return rows


def _style(ws, widths):
    font = Font(name="Arial", size=10)
    head = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F3864")
    for row in ws.iter_rows():
        for c in row:
            c.font = font
            c.alignment = Alignment(vertical="top")
    for c in ws[1]:
        c.font, c.fill = head, fill
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_xlsx(path, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats"
    ws.append(["ligne_source"] + FIELDS)
    ign = wb.create_sheet("Ignorées")
    ign.append(["ligne_source", "nom", "adresse", "raison"])
    for rec in sorted(records, key=lambda r: r["row"]):
        if rec["status"] == "ok":
            d = rec["data"]
            ws.append([rec["row"]] + [d.get(f) for f in FIELDS])
        else:
            ign.append([rec["row"], rec["name"], rec["addr"], rec["reason"]])
    _style(ws, [12, 38, 52, 8, 11, 26, 18, 40, 28])
    _style(ign, [12, 40, 52, 22])
    wb.save(path)


# ------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="agences_prospection.xlsx")
    ap.add_argument("--output-dir", default="output")
    ap.add_argument("--start", type=int, default=0, help="index de départ (0 = 1re agence)")
    ap.add_argument("--max-rows", type=int, default=0, help="0 = tout")
    args = ap.parse_args()

    keys = [k for k in re.split(r"[,\s]+", os.environ.get("SEARLO_API_KEYS", "")) if k]
    if not keys:
        sys.exit("SEARLO_API_KEYS vide (secret GitHub : une clé par ligne ou séparées par des virgules)")
    keys = list(dict.fromkeys(keys))

    rows = read_rows(args.input)
    start = max(args.start, 0)
    end = len(rows) if not args.max_rows else min(len(rows), start + args.max_rows)
    todo = list(enumerate(rows))[start:end]
    print(f"{len(rows)} agences au total ; traitement [{start}:{end}] = {len(todo)} ; "
          f"{len(keys)} clé(s)", flush=True)

    os.makedirs(args.output_dir, exist_ok=True)
    tag = f"{start}-{end}"
    jsonl_path = os.path.join(args.output_dir, f"resultats_{tag}.jsonl")
    xlsx_path = os.path.join(args.output_dir, f"agences_places_{tag}.xlsx")

    pool = KeyPool(keys, RATE_PER_MIN)
    session = requests.Session()
    lock = threading.Lock()
    records, done = [], set()
    t0 = time.time()

    def work(item):
        idx, (row, name, addr) = item
        rec = {"idx": idx, "row": row, "name": name, "addr": addr}
        if not name:
            rec.update(status="skip", reason="nom_vide")
            return rec
        try:
            places = search(session, pool, name, addr)
        except AllKeysDead:
            return None
        except ApiError as e:
            rec.update(status="skip", reason=f"erreur_api:{e}"[:80])
            return rec
        place, reason = pick(name, addr, places)
        if place is None:
            rec.update(status="skip", reason=reason)
        else:
            rec.update(status="ok", reason="ok",
                       data={f: place.get(f) for f in FIELDS})
        return rec

    aborted = False
    with open(jsonl_path, "w", encoding="utf-8") as fh, \
            ThreadPoolExecutor(max_workers=min(max(2, len(keys) * 3), 20)) as ex:
        for c in range(0, len(todo), CHUNK):
            for rec in ex.map(work, todo[c:c + CHUNK]):
                if rec is None:
                    aborted = True
                    continue
                with lock:
                    records.append(rec)
                    done.add(rec["idx"])
                    fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fh.flush()
            ok = sum(r["status"] == "ok" for r in records)
            print(f"[{len(records)}/{len(todo)}] ok={ok} ignorées={len(records) - ok} "
                  f"clés actives={pool.n_alive()} ({time.time() - t0:.0f}s)", flush=True)
            if aborted:
                break

    write_xlsx(xlsx_path, records)
    ok = sum(r["status"] == "ok" for r in records)
    reasons = {}
    for r in records:
        if r["status"] != "ok":
            reasons[r["reason"].split(":")[0]] = reasons.get(r["reason"].split(":")[0], 0) + 1
    lines = [f"**Searlo places** — lignes [{start}:{end}]",
             f"- traitées : {len(records)}/{len(todo)}",
             f"- retenues (match identique) : {ok}",
             f"- ignorées : {len(records) - ok} {reasons}"]
    if aborted:
        nxt = min((i for i, _ in todo if i not in done), default=end)
        lines.append(f"- ⚠️ toutes les clés sont épuisées/invalides. Reprendre avec start_row = {nxt}")
    print("\n".join(lines))
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    sys.exit(1 if aborted else 0)


if __name__ == "__main__":
    main()
          
