"""
One-shot Sleeper league snapshot dump for the static webpage.

Hits the public Sleeper API (no auth) and writes docs/league.json:
    - users: dropdown list of league members
    - rosters: per-roster current players + roster_id -> owner_id mapping
    - players: subset of /players/nfl trimmed to rostered players only
                (full file is ~5 MB, the rostered slice is ~30 KB)
    - draft_picks: live picks (only populates once draft is in_progress / done)

If anything fails (network, schema change, rate limit), this script
prints the error and exits 1. Don't paper over — surface it.

Re-run any time you want to refresh the snapshot:
    python bootstrap_league.py
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone

try:
    import requests
except ImportError:
    print("ERROR: pip install requests", file=sys.stderr)
    sys.exit(1)

LEAGUE_ID = "1335118685588701184"
DRAFT_ID = "1335118685605498880"
OWNER_USER_ID = "1250202714625822720"
DOCS_DIR = "docs"
OUT_PATH = os.path.join(DOCS_DIR, "league.json")

API = "https://api.sleeper.app/v1"
TIMEOUT = 30


def get(url: str):
    r = requests.get(url, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def main() -> int:
    try:
        print(f"GET {API}/league/{LEAGUE_ID}")
        league = get(f"{API}/league/{LEAGUE_ID}")
        print(f"  {league.get('name')!r} — {league.get('total_rosters')} teams, "
              f"season {league.get('season')}, status {league.get('status')}")

        print(f"GET {API}/league/{LEAGUE_ID}/users")
        users = get(f"{API}/league/{LEAGUE_ID}/users")
        print(f"  {len(users)} users")

        print(f"GET {API}/league/{LEAGUE_ID}/rosters")
        rosters = get(f"{API}/league/{LEAGUE_ID}/rosters")
        print(f"  {len(rosters)} rosters")

        print(f"GET {API}/draft/{DRAFT_ID}")
        draft = get(f"{API}/draft/{DRAFT_ID}")
        print(f"  draft status={draft.get('status')!r}")

        draft_picks = []
        if draft.get("status") in ("in_progress", "complete", "paused"):
            print(f"GET {API}/draft/{DRAFT_ID}/picks")
            draft_picks = get(f"{API}/draft/{DRAFT_ID}/picks")
            print(f"  {len(draft_picks)} picks made")
        else:
            print("  draft not started yet — skipping picks fetch")

        print(f"GET {API}/league/{LEAGUE_ID}/traded_picks")
        traded_picks = get(f"{API}/league/{LEAGUE_ID}/traded_picks")
        # Filter to current season only
        season = league.get("season")
        traded_picks_season = [
            tp for tp in traded_picks if str(tp.get("season")) == str(season)
        ]
        print(f"  {len(traded_picks_season)} traded picks for season {season} "
              f"(of {len(traded_picks)} total in league history)")

        print(f"GET {API}/players/nfl  (large, ~5 MB)")
        all_players = get(f"{API}/players/nfl")
        print(f"  {len(all_players)} NFL players in master index")
    except requests.RequestException as e:
        print(f"\nERROR: Sleeper API request failed: {e}", file=sys.stderr)
        return 1
    except (ValueError, KeyError) as e:
        print(f"\nERROR: Sleeper API returned unexpected payload: {e}", file=sys.stderr)
        return 1

    rostered_ids = set()
    for r in rosters:
        for pid in (r.get("players") or []):
            rostered_ids.add(pid)
    for p in draft_picks:
        pid = p.get("player_id")
        if pid:
            rostered_ids.add(pid)

    keep_fields = ("first_name", "last_name", "full_name", "position",
                   "team", "age", "years_exp", "college", "fantasy_positions",
                   "depth_chart_position", "status", "injury_status")
    players_subset = {}
    for pid in rostered_ids:
        p = all_players.get(pid)
        if not p:
            continue
        slim = {k: p.get(k) for k in keep_fields if p.get(k) is not None}
        if "full_name" not in slim:
            slim["full_name"] = (
                f"{slim.get('first_name','')} {slim.get('last_name','')}".strip()
            )
        players_subset[pid] = slim

    users_slim = []
    for u in users:
        users_slim.append({
            "user_id": u.get("user_id"),
            "display_name": u.get("display_name"),
            "team_name": (u.get("metadata") or {}).get("team_name"),
            "avatar": u.get("avatar"),
            "is_owner": u.get("is_owner", False),
        })

    rosters_slim = []
    for r in rosters:
        s = r.get("settings") or {}
        rosters_slim.append({
            "roster_id": r.get("roster_id"),
            "owner_id": r.get("owner_id"),
            "co_owners": r.get("co_owners") or [],
            "players": r.get("players") or [],
            "starters": r.get("starters") or [],
            "taxi": r.get("taxi") or [],
            "reserve": r.get("reserve") or [],
            "wins": s.get("wins"),
            "losses": s.get("losses"),
            "ties": s.get("ties"),
            "fpts": s.get("fpts"),
            "fpts_against": s.get("fpts_against"),
        })

    out = {
        "fetched_at": datetime.now(timezone.utc).isoformat(),
        "league": {
            "league_id": league.get("league_id"),
            "name": league.get("name"),
            "season": league.get("season"),
            "status": league.get("status"),
            "total_rosters": league.get("total_rosters"),
            "scoring_settings": league.get("scoring_settings") or {},
            "roster_positions": league.get("roster_positions") or [],
        },
        "draft": {
            "draft_id": draft.get("draft_id"),
            "status": draft.get("status"),
            "type": draft.get("type"),
            "start_time": draft.get("start_time"),
            "settings": draft.get("settings") or {},
            # Slot -> roster_id mapping (so we can compute each user's picks
            # correctly accounting for draft slot != roster_id).
            "slot_to_roster_id": draft.get("slot_to_roster_id") or {},
        },
        "owner_user_id": OWNER_USER_ID,
        "users": users_slim,
        "rosters": rosters_slim,
        "draft_picks": draft_picks,
        "traded_picks": traded_picks_season,
        "players": players_subset,
    }

    os.makedirs(DOCS_DIR, exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    size_kb = os.path.getsize(OUT_PATH) / 1024
    print(f"\nWrote {OUT_PATH}  ({size_kb:.1f} KB, {len(players_subset)} rostered players)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
