"""
Sleeper post-NFL-draft rookie ADP scraper for dynasty SF TEP leagues.

Usage:
    1. Edit DRAFT_IDS below — paste Sleeper mock-draft / league-draft IDs.
       To find IDs: open the draft on sleeper.com and copy the long number from the URL,
       e.g. https://sleeper.com/draft/nfl/1234567890123456789 → 1234567890123456789
    2. (Optional) Edit DRAFT_FILTERS below to only keep drafts that match your settings.
    3. Run:    python sleeper_scraper.py
    4. Outputs: sources/sleeper_adp.csv (with columns: player, source_rank, n_drafts, mean_pick, earliest, latest)

Requires: pip install requests

Sleeper API docs: https://docs.sleeper.app
No auth needed for public drafts.
"""
from __future__ import annotations
import csv
import json
import os
import sys
import time
from collections import defaultdict
from typing import Dict, List, Optional

try:
    import requests
except ImportError:
    print("ERROR: pip install requests")
    sys.exit(1)

# ============= CONFIG =============

# Paste your Sleeper draft IDs here. Each ID = 1 mock or league rookie draft.
# Find more at sleeper.com → Mock Draft Lobby → join one → copy URL → extract ID
# Or search r/dynastyff for shared mock draft links
DRAFT_IDS: List[str] = [
    # The Phoenix League rookie draft (~May 4, 2026). Becomes useful once it's
    # in_progress / complete. For pre-draft, paste shared mock IDs here.
    "1335118685605498880",
]

# Optional filters — only include drafts where settings match.
# Set to None to disable a filter. Defaults match Colin's league.
DRAFT_FILTERS = {
    "season": "2026",
    "season_type": "off",   # off-season = rookie draft
    "type": "draft_only",   # draft_only = startup or rookie-only
    "rookie_draft": True,   # heuristic: 4-5 rounds, all 2026 rookies
    "superflex": True,
    "tep": True,            # TE premium scoring
    "min_teams": 10,        # ignore tiny mocks
}

# Output paths
OUT_DIR = "sources"
OUT_CSV = os.path.join(OUT_DIR, "sleeper_adp.csv")
OUT_RAW = os.path.join(OUT_DIR, "sleeper_drafts_raw.json")  # for audit

# Sleeper API base
API = "https://api.sleeper.app/v1"
SESSION = requests.Session()
SESSION.headers["User-Agent"] = "dynasty-board-builder/1.0"


# ============= HELPERS =============

def get(path: str, retries: int = 3) -> Optional[dict | list]:
    """GET with retry. Returns None on terminal failure."""
    url = f"{API}/{path.lstrip('/')}"
    for i in range(retries):
        try:
            r = SESSION.get(url, timeout=15)
            if r.status_code == 200:
                return r.json()
            if r.status_code == 404:
                return None
            print(f"  [warn] {url} → HTTP {r.status_code}, retry {i+1}/{retries}")
        except Exception as e:
            print(f"  [warn] {url} → {e}, retry {i+1}/{retries}")
        time.sleep(1 + i)
    return None


_PLAYER_CACHE: Optional[Dict[str, dict]] = None

def players_index() -> Dict[str, dict]:
    """Fetch the full Sleeper player index. Cache in memory; ~5MB."""
    global _PLAYER_CACHE
    if _PLAYER_CACHE is not None:
        return _PLAYER_CACHE
    print("Fetching Sleeper player index (one-time, ~5MB)...")
    data = get("players/nfl")
    if not data:
        print("ERROR: failed to fetch player index")
        sys.exit(1)
    _PLAYER_CACHE = data
    print(f"  → {len(data)} players cached")
    return data


def fetch_draft(draft_id: str) -> Optional[dict]:
    """Fetch draft metadata + picks."""
    print(f"Fetching draft {draft_id}...")
    meta = get(f"draft/{draft_id}")
    if not meta:
        print(f"  [skip] draft {draft_id} not found")
        return None
    picks = get(f"draft/{draft_id}/picks")
    if picks is None:
        picks = []
    return {"meta": meta, "picks": picks}


def is_rookie_draft(meta: dict, picks: list) -> bool:
    """Heuristic: a rookie draft has all picks belonging to players with rookie_year == season."""
    season = meta.get("season")
    if not season:
        return False
    if not picks:
        return False
    rookie_year = season  # "2026"
    pidx = players_index()
    rookie_count = 0
    total = 0
    for pk in picks:
        pid = pk.get("player_id")
        if not pid:
            continue
        total += 1
        p = pidx.get(pid, {})
        years_exp = p.get("years_exp")
        # rookies have years_exp == 0 currently or rookie_year matches
        if years_exp == 0 or str(p.get("metadata", {}).get("rookie_year")) == rookie_year:
            rookie_count += 1
    if total == 0:
        return False
    return rookie_count / total >= 0.85


def passes_filters(meta: dict, picks: list) -> tuple[bool, str]:
    """Return (passes, reason)."""
    settings = meta.get("settings", {})
    scoring = meta.get("scoring_settings") or meta.get("scoring", {})
    if DRAFT_FILTERS.get("season") and meta.get("season") != DRAFT_FILTERS["season"]:
        return False, f"wrong season ({meta.get('season')})"
    if DRAFT_FILTERS.get("min_teams"):
        teams = settings.get("teams") or len(meta.get("draft_order") or {})
        if teams and teams < DRAFT_FILTERS["min_teams"]:
            return False, f"too few teams ({teams})"
    if DRAFT_FILTERS.get("rookie_draft") and not is_rookie_draft(meta, picks):
        return False, "not a rookie draft (mixed exp years)"
    # Superflex: at least 2 QB starting slots OR roster has SF flag
    if DRAFT_FILTERS.get("superflex"):
        roster_pos = settings.get("roster_positions") or meta.get("roster_positions") or []
        slots_qb = sum(1 for s in roster_pos if s == "QB")
        has_sf = any(s in ("SUPER_FLEX", "SF") for s in roster_pos)
        if slots_qb < 2 and not has_sf:
            return False, f"not superflex (roster: {roster_pos})"
    # TEP: rec_te > rec or bonus_rec_te > 0
    if DRAFT_FILTERS.get("tep") and isinstance(scoring, dict):
        rec_te = scoring.get("rec_te") or 0
        rec = scoring.get("rec") or 0
        bonus_te = scoring.get("bonus_rec_te") or 0
        if rec_te <= rec and bonus_te <= 0:
            return False, f"not TEP (rec_te={rec_te}, rec={rec})"
    return True, "OK"


# ============= AGGREGATION =============

def aggregate_picks(drafts: List[dict]) -> Dict[str, dict]:
    """Across all kept drafts, compute per-player ADP."""
    pidx = players_index()
    by_player: Dict[str, List[int]] = defaultdict(list)
    for d in drafts:
        n_teams = d["meta"].get("settings", {}).get("teams", 12)
        for pk in d["picks"]:
            pid = pk.get("player_id")
            if not pid:
                continue
            overall = pk.get("pick_no") or pk.get("pick", {}).get("pick_no")
            if not overall:
                continue
            by_player[pid].append(overall)

    out = {}
    for pid, picks in by_player.items():
        p = pidx.get(pid, {})
        name = (p.get("full_name") or
                f"{p.get('first_name','')} {p.get('last_name','')}").strip()
        pos = p.get("position", "")
        team = p.get("team", "FA") or "FA"
        n = len(picks)
        mean_pk = sum(picks) / n
        out[name] = {
            "player": name,
            "pos": pos,
            "team": team,
            "n_drafts": n,
            "mean_pick": round(mean_pk, 2),
            "earliest": min(picks),
            "latest": max(picks),
        }
    # Rank by mean pick
    ordered = sorted(out.values(), key=lambda x: x["mean_pick"])
    for i, r in enumerate(ordered, 1):
        r["source_rank"] = i
    return {r["player"]: r for r in ordered}


# ============= MAIN =============

def main():
    if not DRAFT_IDS:
        print("ERROR: DRAFT_IDS list is empty. Edit this file and paste Sleeper draft IDs.")
        print("To find IDs: sleeper.com → join/find a mock draft → URL contains the ID.")
        print("Look for: 12-team, SF (superflex), TEP, 2026 rookie-only, post-NFL-draft.")
        sys.exit(1)

    os.makedirs(OUT_DIR, exist_ok=True)
    drafts_raw = []
    drafts_kept = []
    for did in DRAFT_IDS:
        d = fetch_draft(did)
        if not d:
            continue
        drafts_raw.append(d)
        ok, reason = passes_filters(d["meta"], d["picks"])
        print(f"  {did}: {len(d['picks'])} picks → {'KEEP' if ok else 'SKIP'} ({reason})")
        if ok:
            drafts_kept.append(d)
        time.sleep(0.2)  # be polite

    if not drafts_kept:
        print("\nNo drafts passed filters. Check DRAFT_FILTERS or paste different IDs.")
        sys.exit(1)

    with open(OUT_RAW, "w") as f:
        json.dump(drafts_raw, f, default=str)
    print(f"\nKept {len(drafts_kept)}/{len(drafts_raw)} drafts. Aggregating...")

    adp = aggregate_picks(drafts_kept)

    with open(OUT_CSV, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["player", "pos", "team", "source_rank", "n_drafts", "mean_pick", "earliest", "latest"])
        for r in adp.values():
            w.writerow([r["player"], r["pos"], r["team"], r["source_rank"],
                        r["n_drafts"], r["mean_pick"], r["earliest"], r["latest"]])
    print(f"\nWrote {len(adp)} players to {OUT_CSV}")
    print(f"Top 15 by ADP:")
    for r in list(adp.values())[:15]:
        print(f"  {r['source_rank']:>2}. {r['player']:25s} {r['pos']:2s} {r['team']:4s}  "
              f"mean={r['mean_pick']:5.1f}  n={r['n_drafts']}")


if __name__ == "__main__":
    main()
