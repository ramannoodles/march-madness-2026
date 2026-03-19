"""
app.py  —  2026 Office March Madness Web App
Run locally:  python app.py
Deploy:       Railway / Render (see README)
"""

from flask import Flask, render_template, jsonify
from openpyxl import load_workbook
import os, glob
from datetime import datetime

app = Flask(__name__)

PLAYERS = ["Raman", "Porter", "Chuck", "Shelly", "Kim", "Tucker"]

ROUND_POINTS = {
    "Round of 64":  1,
    "Round of 32":  2,
    "Sweet 16":     4,
    "Elite 8":      8,
    "Final 4":      16,
    "Championship": 32,
}

ROUND_ORDER = ["Round of 64", "Round of 32", "Sweet 16", "Elite 8", "Final 4", "Championship"]


def find_excel_file():
    """Find the bracket Excel file regardless of exact filename casing."""
    env_file = os.environ.get("EXCEL_FILE")
    if env_file and os.path.exists(env_file):
        return env_file
    search_dir = os.path.dirname(os.path.abspath(__file__))
    for pattern in ["*Bracket*.xlsx", "*bracket*.xlsx", "*Office*.xlsx", "*.xlsx"]:
        matches = glob.glob(os.path.join(search_dir, pattern))
        if matches:
            return matches[0]
    raise FileNotFoundError(
        "No Excel bracket file found. Place your .xlsx file in the same folder as app.py"
    )


def load_data():
    excel_path = find_excel_file()
    wb = load_workbook(excel_path, data_only=True)

    # ── Master Results ────────────────────────────────────────────────────────
    wm = wb["Master_Results"]
    results = {}
    for row in wm.iter_rows(min_row=1, max_row=wm.max_row, values_only=True):
        gid = row[1] if len(row) > 1 else None
        if not gid or not isinstance(gid, str):
            continue
        if not gid.startswith(("R64", "Rou", "Swe", "Eli", "Fin", "Cha")):
            continue
        results[gid] = {
            "round":   row[0],
            "matchup": row[2],
            "winner":  row[3],
        }

    # ── Player Picks ──────────────────────────────────────────────────────────
    # Scores computed directly from master results — not from cached Excel formulas
    players = {}
    for p in PLAYERS:
        if p not in wb.sheetnames:
            continue
        ws = wb[p]
        picks = []

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            gid = row[1] if len(row) > 1 else None
            if not gid or not isinstance(gid, str):
                continue
            if not gid.startswith(("R64", "Rou", "Swe", "Eli", "Fin", "Cha")):
                continue

            rnd     = row[0]
            matchup = row[2]
            pick    = row[3]
            pts_val = row[4] if row[4] else ROUND_POINTS.get(rnd, 1)

            master = results.get(gid, {})
            actual = master.get("winner")
            if actual is None:
                result = None
            elif pick and str(pick).strip().lower() == str(actual).strip().lower():
                result = 1
            else:
                result = 0

            picks.append({
                "game_id":   gid,
                "round":     rnd,
                "matchup":   matchup,
                "pick":      pick,
                "pts_value": pts_val,
                "result":    result,
                "score":     pts_val if result == 1 else 0,
            })

        total         = sum(pk["score"] for pk in picks)
        pts_remaining = sum(pk["pts_value"] for pk in picks if pk["result"] is None)
        players[p] = {
            "picks":         picks,
            "total":         total,
            "pts_remaining": pts_remaining,
            "max_possible":  total + pts_remaining,
        }

    # ── Leaderboard ───────────────────────────────────────────────────────────
    leaderboard = sorted(
        [{"name": name, **data} for name, data in players.items()],
        key=lambda x: (-x["total"], x["name"])
    )
    prev_score = None
    rank_counter = 1
    for i, entry in enumerate(leaderboard):
        if entry["total"] != prev_score:
            rank_counter = i + 1
        entry["rank"] = rank_counter
        prev_score = entry["total"]
        entry["champion_pick"] = next(
            (pk["pick"] for pk in players[entry["name"]]["picks"] if pk["game_id"] == "Cha_G1"),
            "—"
        )

    completed   = sum(1 for v in results.values() if v["winner"])
    total_games = len(results)

    return {
        "leaderboard":     leaderboard,
        "players":         players,
        "results":         results,
        "completed_games": completed,
        "total_games":     total_games,
        "last_updated":    datetime.now().strftime("%b %d, %Y %I:%M %p"),
    }


@app.route("/")
def index():
    data = load_data()
    return render_template("index.html", data=data, players=PLAYERS)


@app.route("/api/data")
def api_data():
    data = load_data()
    return jsonify({
        "leaderboard": [
            {k: v for k, v in e.items() if k != "picks"} for e in data["leaderboard"]
        ],
        "completed_games": data["completed_games"],
        "total_games":     data["total_games"],
        "last_updated":    data["last_updated"],
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
