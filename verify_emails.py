#!/usr/bin/env python3
"""
Vérification des e-mails probables des gérants via l'API Searlo (1 ou plusieurs clés).

Pour chaque gérant : on teste les e-mails probables 1 par 1.
  - premier e-mail au statut "valid" (et rien d'autre) -> on le garde, gérant suivant
  - aucun "valid" -> "no email"
Un fichier de sortie .xlsx par agence. Reprise automatique si le fichier de sortie existe déjà.
"""
import argparse
import os
import re
import sys
import time
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE_URL = os.environ.get("SEARLO_BASE_URL", "https://api.searlo.tech/api/v1")
VERIFY_URL = f"{BASE_URL}/email-verifier/verify"
NO_EMAIL = "no email"
SAVE_EVERY = 10  # gérants entre deux sauvegardes

FONT = "Arial"


class FatalError(Exception):
    """Erreur qui doit stopper le run (clé invalide, plus de crédits...)."""


class KeyPool:
    """Plusieurs clés API : on reste sur la courante, on bascule quand elle est morte (402/401/403)."""

    def __init__(self, keys, start=0):
        self.keys = keys
        self.dead = set()
        self.i = start % len(keys)

    def current(self):
        return self.keys[self.i]

    def burn(self, reason):
        print(f"  ! clé #{self.i + 1}/{len(self.keys)} retirée ({reason})", flush=True)
        self.dead.add(self.i)
        for step in range(1, len(self.keys) + 1):
            j = (self.i + step) % len(self.keys)
            if j not in self.dead:
                self.i = j
                print(f"  -> bascule sur la clé #{j + 1}", flush=True)
                return
        raise FatalError("Toutes les clés API sont épuisées ou refusées.")


def load_keys():
    raw = os.environ.get("SEARLO_API_KEYS", "") + "," + os.environ.get("SEARLO_API_KEY", "")
    keys = []
    for k in re.split(r"[,;\s]+", raw):
        if k and k not in keys:
            keys.append(k)
    return keys


def verify_email(session, email, pool, cache):
    """Retourne le statut Searlo ('valid', 'invalid', 'accept_all', 'unknown', 'error', ...)."""
    key = email.lower()
    if key in cache:
        return cache[key]

    busy_retries = 0
    server_retries = 0
    while True:
        try:
            resp = session.post(
                VERIFY_URL,
                headers={"x-api-key": pool.current(), "Content-Type": "application/json"},
                json={"email": email},
                timeout=60,
            )
        except requests.RequestException as exc:
            server_retries += 1
            if server_retries > 4:
                print(f"  ! réseau KO pour {email}: {exc}", flush=True)
                return "error"
            time.sleep(5 * server_retries)
            continue

        code = resp.status_code
        if code == 429:  # VERIFIER_BUSY ou rate limit : non facturé, on attend
            busy_retries += 1
            if busy_retries > 30:
                raise FatalError("Trop de 429 consécutifs, on arrête (relance plus tard).")
            try:
                wait = int(resp.headers.get("Retry-After", 30))
            except ValueError:
                wait = 30
            time.sleep(max(wait, 1))
            continue
        if code in (401, 403):
            pool.burn(f"refusée, HTTP {code}")  # lève FatalError s'il n'en reste plus
            continue
        if code == 402:
            pool.burn("plus de crédits, HTTP 402")
            continue
        if code == 400:
            cache[key] = "invalid_format"
            return "invalid_format"
        if code >= 500:
            server_retries += 1
            if server_retries > 4:
                print(f"  ! HTTP {code} persistant pour {email}", flush=True)
                return "error"
            time.sleep(5 * server_retries)
            continue
        if code != 200:
            print(f"  ! HTTP {code} inattendu pour {email}", flush=True)
            return "error"

        try:
            status = resp.json()["data"]["status"]
        except (ValueError, KeyError, TypeError):
            print(f"  ! réponse illisible pour {email}", flush=True)
            return "error"
        cache[key] = status
        return status


def read_input(path):
    ws = load_workbook(path, read_only=True).active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = (row[0] or "").strip() if isinstance(row[0], str) else row[0]
        if not name:
            continue
        emails = [str(c).strip() for c in row[1:6] if c and str(c).strip()]
        rows.append((str(name), emails))
    return rows


def read_existing(path):
    if not path.exists():
        return []
    ws = load_workbook(path, read_only=True).active
    return [
        (r[0], r[1], r[2] if len(r) > 2 else "")
        for r in ws.iter_rows(min_row=2, values_only=True)
        if r[0]
    ]


def write_output(path, results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats"
    headers = ["Nom prénom du gérant", "E-mail validé", "Détail"]
    ws.append(headers)
    head_fill = PatternFill("solid", start_color="1F3864")
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(vertical="center")
    miss_fill = PatternFill("solid", start_color="FDE9E7")
    for name, email, detail in results:
        ws.append([name, email, detail])
        r = ws.max_row
        for c in ws[r]:
            c.font = Font(name=FONT)
            if email == NO_EMAIL:
                c.fill = miss_fill
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 70
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{max(ws.max_row, 2)}"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(path)


def out_name(in_path):
    stem = in_path.stem.replace("_emails_probables", "_emails_valides")
    return f"{stem}.xlsx"


def process_file(in_path, out_dir, session, pool, deadline):
    out_path = out_dir / out_name(in_path)
    gerants = read_input(in_path)

    results = read_existing(out_path)
    # reprise seulement si les noms déjà traités correspondent à l'entrée
    if results and [r[0] for r in results] != [g[0] for g in gerants[: len(results)]]:
        print("  sortie existante incohérente avec l'entrée -> on repart de zéro", flush=True)
        results = []
    start = len(results)
    print(f"\n=== {in_path.name} : {len(gerants)} gérants (reprise à {start}) ===", flush=True)

    cache = {}
    checks = 0
    complete = True
    try:
        for i in range(start, len(gerants)):
            if time.time() > deadline:
                print("  budget temps atteint -> sauvegarde et arrêt propre", flush=True)
                complete = False
                break
            name, emails = gerants[i]
            found, trace = NO_EMAIL, []
            for n, email in enumerate(emails, 1):
                status = verify_email(session, email, pool, cache)
                checks += 1
                trace.append(f"{n}:{status}")
                if status == "valid":  # uniquement "valid", rien d'autre
                    found = email
                    break
            results.append((name, found, " | ".join(trace)))
            print(f"  [{i + 1}/{len(gerants)}] {name} -> {found}", flush=True)
            if len(results) % SAVE_EVERY == 0:
                write_output(out_path, results)
    except FatalError as exc:
        write_output(out_path, results)
        print(f"::error::{exc} (progression sauvegardée : {len(results)}/{len(gerants)})")
        raise
    write_output(out_path, results)
    ok = sum(1 for r in results if r[1] != NO_EMAIL)
    print(f"  -> {out_path} : {ok} e-mails valides / {len(results)} gérants, {checks} appels", flush=True)
    return complete


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+", help="fichiers xlsx d'entrée")
    ap.add_argument("--out-dir", default="output")
    ap.add_argument("--max-minutes", type=float, default=330, help="budget temps avant arrêt propre")
    args = ap.parse_args()

    keys = load_keys()
    if not keys:
        sys.exit("Aucune clé : définis SEARLO_API_KEYS (séparées par virgule ou retour à la ligne)")
    # chaque job de la matrice démarre sur une clé différente pour ne pas tous taper la même
    pool = KeyPool(keys, start=int(os.environ.get("KEY_OFFSET", "0")))
    print(f"{len(keys)} clé(s) chargée(s), départ sur la clé #{pool.i + 1}", flush=True)

    deadline = time.time() + args.max_minutes * 60
    session = requests.Session()
    all_done = True
    try:
        for p in args.inputs:
            all_done &= process_file(Path(p), Path(args.out_dir), session, pool, deadline)
    except FatalError:
        sys.exit(1)
    if not all_done:
        print("::warning::Run incomplet (limite de temps). Relance le workflow : il reprend là où il s'est arrêté.")


if __name__ == "__main__":
    main()
