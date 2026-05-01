"""
FantasyCalc SF dynasty values fetcher.

Pulls current rookie & overall dynasty SF values, filters to 2026 rookies,
outputs sources/fc.csv.

API docs: https://github.com/zachatrocity/fantasycalc-api (community-documented public API)
Endpoint: https://api.fantasycalc.com/values/current
Params: isDynasty (bool), numQbs (1 or 2), numTeams (int), ppr (0 / 0.5 / 1)
Note: FantasyCalc has no public TEP toggle — values are non-TEP. KTC remains the TEP source.

Usage: python fc_fetcher.py
Requires: pip install requests
"""
from __future__ import annotations
import csv
import json
import os
import sys
from typing import Optional

try:
    import requests
except ImportError:
    print("ERROR: pip install requests")
    sys.exit(1)

OUT_DIR = "sources"
OUT_CSV = os.path.join(OUT_DIR, "fc.csv")

URL = "https://api.fantasycalc.com/values/current"
PARAMS = {"isDynasty": "true", "numQbs": 2, "numTeams": 12, "ppr": 0.5}


def main():
    print(f"Fetching {URL} ...")
    try:
        r = requests.get(URL, params=PARAMS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
    data = r.json()
    print(f"  -> {len(data)} entries")

    # FC schema (verified Apr 2026): player.maybeYoe == 0 is the rookie signal.
    # The old player.rookie boolean was removed.
    rookies = []
    for entry in data:
        p = entry.get("player", {})
        if p.get("maybeYoe") != 0:
            continue
        rookies.append({
            "player": p.get("name"),
            "pos": p.get("position"),
            "team": p.get("maybeTeam") or "FA",
            "value": entry.get("value"),
            "overall_rank": entry.get("overallRank"),
            "position_rank": entry.get("positionRank"),
            "sleeper_id": p.get("sleeperId"),
        })
    rookies.sort(key=lambda x: -x["value"])
    for i, r in enumerate(rookies, 1):
        r["source_rank"] = i

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_CSV, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["player", "pos", "team", "source_rank", "value", "overall_rank", "position_rank", "sleeper_id"])
        for r in rookies:
            w.writerow([r["player"], r["pos"], r["team"], r["source_rank"],
                        r["value"], r["overall_rank"], r["position_rank"], r.get("sleeper_id") or ""])
    print(f"\nWrote {len(rookies)} 2026 rookies to {OUT_CSV}")
    print("Top 15:")
    for r in rookies[:15]:
        print(f"  {r['source_rank']:>2}. {r['player']:25s} {r['pos']:2s} {r['team']:4s}  val={r['value']}")


if __name__ == "__main__":
    main()
