#!/usr/bin/env python3
"""Enrichit agences_prospection.xlsx via Searlo GET /search/places.

Une ligne source = une requête (q = nom, location = adresse si dispo, limit = 1).
Le résultat renvoyé est gardé tel quel (aucune vérification nom/adresse) ; une
ligne sans résultat est ignorée (raison tracée dans l'onglet "Ignorées").
Plusieurs clés API : rotation (la moins récemment utilisée d'abord), cooldown sur
429 (plafonné), clé retirée du pool sur 401/402/403.
Sauvegarde tous les 100 résultats (xlsx partiel + checkpoint) ; une relance reprend
automatiquement là où le run précédent s'est arrêté (lignes en erreur retentées).
Robustesse : pas de barrière entre lots (une requête lente ne bloque pas le reste),
heartbeat toutes les 30 s dans les logs, arrêt propre (xlsx écrit) sur annulation
du run ou si plus aucune ligne n'avance pendant STALL_TIMEOUT secondes.
"""
import argparse
import json
import os
import re
import signal
import sys
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor
from concurrent.futures import wait as fwait

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = os.environ.get("SEARLO_BASE_URL", "https://api.searlo.tech/api/v1")
# Limite Searlo Maps/Places : 240 req/min par compte -> marge de sécurité
RATE_PER_MIN = int(os.environ.get("SEARLO_RATE_PER_MIN", "220"))
LIMIT = 1
MAX_ATTEMPTS = 8            # tentatives hors 429 (réseau, 5xx)
CONNECT_TIMEOUT = 10
READ_TIMEOUT = int(os.environ.get("SEARLO_READ_TIMEOUT", "30"))
MAX_COOLDOWN = 120          # plafond (s) du Retry-After sur 429
STALL_TIMEOUT = int(os.environ.get("STALL_TIMEOUT", "600"))   # s sans aucune ligne terminée
HEARTBEAT = int(os.environ.get("HEARTBEAT", "30"))
FIELDS = ["title", "address", "rating", "ratingCount", "category",
          "phoneNumber", "website", "domain"]


def pick(name, addr, places):
    """Garde le résultat renvoyé tel quel. Retourne (place, raison)."""
    if not places:
        return None, "aucun_resultat"
    return places[0], "ok"


# ---------------------------------------------------------------- rotation de clés
class AllKeysDead(Exception):
    pass


class Stopped(Exception):
    pass


class ApiError(Exception):
    pass


class KeyPool:
    def __init__(self, keys, per_min, stop):
        self.lock = threading.Lock()
        self.stop = stop
        self.interval = 60.0 / per_min
        self.keys = {k: {"next": 0.0, "dead": None} for k in keys}
        self.n429 = 0

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
            self.stop.wait(slot - now)
        return k

    def cooldown(self, k, sec, raw):
        with self.lock:
            self.keys[k]["next"] = max(self.keys[k]["next"], time.monotonic() + sec)
            self.n429 += 1
            n = self.n429
        if n <= 10 or n % 100 == 0:
            print(f"[429 #{n}] clé {self.mask(k)} Retry-After={raw!r} -> pause {sec:.0f}s",
                  flush=True)

    def kill(self, k, status):
        with self.lock:
            if not self.keys[k]["dead"]:
                self.keys[k]["dead"] = status
                print(f"[clé {self.mask(k)}] retirée du pool (HTTP {status})", flush=True)

    def n_alive(self):
        with self.lock:
            return sum(1 for v in self.keys.values() if not v["dead"])

    def status(self):
        now = time.monotonic()
        with self.lock:
            return " ".join(
                f"{self.mask(k)}:{'morte' if v['dead'] else str(max(0, round(v['next'] - now))) + 's'}"
                for k, v in self.keys.items())


def search(session, pool, stop, name, addr):
    params = {"q": name, "limit": LIMIT, "gl": "fr", "hl": "fr"}
    if addr:
        params["location"] = addr[:200]
    last, attempt = "erreur_api", 0
    while attempt < MAX_ATTEMPTS:
        if stop.is_set():
            raise Stopped()
        key = pool.acquire()
        if key is None:
            raise AllKeysDead()
        if stop.is_set():
            raise Stopped()
        try:
            r = session.get(f"{BASE_URL}/search/places", params=params,
                            headers={"x-api-key": key},
                            timeout=(CONNECT_TIMEOUT, READ_TIMEOUT))
        except requests.RequestException as e:
            last = f"reseau:{type(e).__name__}"
            attempt += 1
            stop.wait(min(2 ** attempt, 30))
            continue
        s = r.status_code
        if s == 200:
            try:
                return r.json().get("places") or []
            except ValueError:
                last = "json_invalide"
                attempt += 1
                continue
        if s == 429:   # n'entame pas MAX_ATTEMPTS : on attend que la clé se libère
            raw = r.headers.get("Retry-After")
            try:
                sec = float(raw)
            except (TypeError, ValueError):
                sec = 15.0
            pool.cooldown(key, min(sec, MAX_COOLDOWN), raw)
            last = "429"
            continue
        if s in (401, 402, 403):
            pool.kill(key, s)
            continue
        if s in (500, 502, 503, 504):
            last = str(s)
            attempt += 1
            stop.wait(min(2 ** attempt, 30))
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


def write_xlsx(path, records, styled=True):
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
    w1, w2 = [12, 38, 52, 8, 11, 26, 18, 40, 28], [12, 40, 52, 22]
    if styled:
        _style(ws, w1)
        _style(ign, w2)
    else:  # sauvegarde intermédiaire : rapide, sans mise en forme cellule par cellule
        for sh, w in ((ws, w1), (ign, w2)):
            for i, x in enumerate(w, start=1):
                sh.column_dimensions[get_column_letter(i)].width = x
            sh.freeze_panes = "A2"
    tmp = path[:-5] + ".tmp.xlsx"
    wb.save(tmp)
    os.replace(tmp, path)   # écriture atomique : jamais de xlsx à moitié écrit


# ------------------------------------------------------------------------- main
SAVE_EVERY = 100


def load_checkpoint(path, rows):
    """Relit le jsonl cumulatif. Garde uniquement les lignes qui correspondent
    encore au fichier source ; les lignes en erreur seront retentées."""
    prev, retry = {}, 0
    if not os.path.exists(path):
        return prev, retry
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            try:
                rec = json.loads(line)
            except ValueError:      # dernière ligne tronquée (arrêt brutal)
                continue
            i = rec.get("idx")
            if not isinstance(i, int) or not 0 <= i < len(rows):
                continue
            if (rec.get("row"), rec.get("name"), rec.get("addr")) != rows[i]:
                continue            # le fichier source a changé : on ignore
            prev[i] = rec           # la dernière version d'une ligne l'emporte
    for i in [i for i, r in prev.items() if str(r.get("reason", "")).startswith("erreur")]:
        del prev[i]
        retry += 1
    return prev, retry


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="agences_prospection.xlsx")
    ap.add_argument("--output-dir", default="output")
    ap.add_argument("--checkpoint-dir", default="checkpoint")
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

    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(args.checkpoint_dir, exist_ok=True)
    ck_jsonl = os.path.join(args.checkpoint_dir, "resultats.jsonl")
    ck_meta = os.path.join(args.checkpoint_dir, "checkpoint.json")
    xlsx_path = os.path.join(args.output_dir, "agences_places.xlsx")

    # --- reprise depuis le checkpoint
    records, retried = load_checkpoint(ck_jsonl, rows)
    if os.path.exists(ck_jsonl) and os.path.getsize(ck_jsonl):
        with open(ck_jsonl, "rb+") as f:      # dernière ligne tronquée : on la termine
            f.seek(-1, os.SEEK_END)
            if f.read(1) != b"\n":
                f.write(b"\n")
    reprises = sum(1 for i in records if start <= i < end)
    todo = [(i, rows[i]) for i in range(start, end) if i not in records]
    plage_total = end - start
    n_plage = reprises
    print(f"{len(rows)} agences au total ; plage [{start}:{end}] = {plage_total} ; "
          f"reprises du checkpoint : {reprises} ; lignes retentées (erreur) : {retried} ; "
          f"à traiter : {len(todo)} ; {len(keys)} clé(s)", flush=True)

    stop = threading.Event()
    signaled = []

    def on_signal(signum, frame):
        signaled.append(signum)
        print(f"[signal {signum}] arrêt demandé : sauvegarde des résultats…", flush=True)
        stop.set()

    signal.signal(signal.SIGINT, on_signal)
    signal.signal(signal.SIGTERM, on_signal)

    pool = KeyPool(keys, RATE_PER_MIN, stop)
    session = requests.Session()
    lock = threading.Lock()
    inflight = {}
    t0 = time.time()

    def work(item):
        idx, (row, name, addr) = item
        rec = {"idx": idx, "row": row, "name": name, "addr": addr}
        if stop.is_set():
            return None
        if not name:
            rec.update(status="skip", reason="nom_vide")
            return rec
        with lock:
            inflight[idx] = (time.monotonic(), name)
        try:
            places = search(session, pool, stop, name, addr)
        except (AllKeysDead, Stopped):
            return None
        except ApiError as e:
            rec.update(status="skip", reason=f"erreur_api:{e}"[:80])
            return rec
        except Exception as e:  # filet de sécurité : une ligne ne doit jamais tout bloquer
            rec.update(status="skip", reason=f"erreur:{type(e).__name__}")
            return rec
        finally:
            with lock:
                inflight.pop(idx, None)
        place, reason = pick(name, addr, places)
        if place is None:
            rec.update(status="skip", reason=reason)
        else:
            rec.update(status="ok", reason="ok", data={f: place.get(f) for f in FIELDS})
        return rec

    def heartbeat():
        with lock:
            n = len(inflight)
            oldest = max(((time.monotonic() - t, nm) for t, nm in inflight.values()),
                         default=(0, ""))
        print(f"[hb] {n_plage}/{plage_total} | en cours: {n} (plus ancienne: "
              f"{oldest[0]:.0f}s « {oldest[1][:30]} ») | clés: {pool.status()}", flush=True)

    def save(final=False):
        """Sauvegarde : jsonl cumulatif + xlsx + checkpoint.json (écritures atomiques)."""
        fh.flush()
        write_xlsx(xlsx_path, list(records.values()), styled=final)
        missing = next((i for i in range(len(rows)) if i not in records), None)
        meta = {"mis_a_jour": time.strftime("%Y-%m-%d %H:%M:%S"), "input": args.input,
                "total_agences": len(rows), "lignes_faites": len(records),
                "prochaine_ligne_manquante": missing}
        with open(ck_meta + ".tmp", "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=1)
        os.replace(ck_meta + ".tmp", ck_meta)

    workers = min(max(2, len(keys) * 3), 20)
    ex = ThreadPoolExecutor(max_workers=workers)
    pending = {ex.submit(work, it) for it in todo}
    aborted = None
    last_progress = last_hb = time.monotonic()
    new_count, next_save = 0, SAVE_EVERY
    fh = open(ck_jsonl, "a", encoding="utf-8")
    try:
        while pending and not stop.is_set():
            done_f, pending = fwait(pending, timeout=2, return_when=FIRST_COMPLETED)
            for f in done_f:
                rec = f.result()
                if rec is None:
                    if not stop.is_set():
                        aborted = "cles"
                        stop.set()
                    continue
                records[rec["idx"]] = rec
                new_count += 1
                n_plage += 1
                fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
                last_progress = time.monotonic()
            now = time.monotonic()
            if new_count >= next_save:
                while next_save <= new_count:
                    next_save += SAVE_EVERY
                save()
                ok = sum(r["status"] == "ok" for r in records.values())
                print(f"[{n_plage}/{plage_total}] ok={ok} ignorées={len(records) - ok} "
                      f"clés actives={pool.n_alive()} ({time.time() - t0:.0f}s) — checkpoint sauvé",
                      flush=True)
            if now - last_hb >= HEARTBEAT:
                heartbeat()
                last_hb = now
            if now - last_progress > STALL_TIMEOUT:
                aborted = "blocage"
                print(f"[blocage] aucune ligne terminée depuis {STALL_TIMEOUT}s : arrêt propre",
                      flush=True)
                heartbeat()
                break
    finally:
        stop.set()
        ex.shutdown(wait=False, cancel_futures=True)
        save(final=True)
        fh.close()

    recs = list(records.values())
    ok = sum(r["status"] == "ok" for r in recs)
    reasons = {}
    for r in recs:
        if r["status"] != "ok":
            k = r["reason"].split(":")[0]
            reasons[k] = reasons.get(k, 0) + 1
    lines = [f"**Searlo places** — plage [{start}:{end}]",
             f"- plage traitée : {n_plage}/{plage_total} (dont {reprises} reprises du checkpoint)",
             f"- total cumulé : {len(recs)}/{len(rows)} | avec fiche : {ok}",
             f"- ignorées : {len(recs) - ok} {reasons}"]
    incomplete = n_plage < plage_total
    if incomplete:
        why = {"cles": "toutes les clés sont épuisées/invalides",
               "blocage": f"plus de progression depuis {STALL_TIMEOUT}s (429 prolongé ?)"
               }.get(aborted, "run interrompu" if signaled else "arrêt")
        lines.append(f"- ⚠️ {why}. Relance le workflow : reprise automatique du checkpoint "
                     f"({len(recs)} lignes déjà faites)")
    print("\n".join(lines), flush=True)
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    sys.stdout.flush()
    os._exit(1 if incomplete else 0)   # évite d'attendre d'éventuels threads bloqués


if __name__ == "__main__":
    main()
