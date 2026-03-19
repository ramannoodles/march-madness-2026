"""
app.py  —  2026 Office March Madness Web App
Run locally:  python app.py
Deploy:       Railway / Render (see README)
"""

from flask import Flask, render_template, jsonify
from openpyxl import load_workbook
import os, re
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = os.environ.get("EXCEL_FILE", "2026_Office_Bracket_UPGRADED.xlsx")
PLAYERS    = ["Raman", "Porter", "Chuck", "Shelly", "Kim", "Tucker"]

ROUND_ORDER = ["Round of 64", "Round of 32", "Sweet 16", "Elite 8", "Final 4", "Championship"]

ROUND_POINTS = {
    "Round of 64": 1,
    "Round of 32": 2,
    "Sweet 16":    4,
    "Elite 8":     8,
    "Final 4":     16,
    "Championship": 32,
}


def load_data():
    wb = load_workbook(EXCEL_FILE, data_only=True)

    # ── Master Results ────────────────────────────────────────────────────────
    wm = wb["Master_Results"]
    results = {}   # game_id → { round, matchup, winner }
    for row in wm.iter_rows(min_row=1, max_row=wm.max_row, values_only=True):
        if not row[1] or not isinstance(row[1], str) or not row[1].startswith(("R64", "Rou", "Swe", "Eli", "Fin", "Cha")):
            continue
        results[row[1]] = {
            "round":   row[0],
            "matchup": row[2],
            "winner":  row[3],
        }

    # ── Player Picks ──────────────────────────────────────────────────────────
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
            if not gid.startswith(("R64","Rou","Swe","Eli","Fin","Cha")):
                continue
            rnd    = row[0]
            matchup= row[2]
            pick   = row[3]
            pts_w  = row[4] or ROUND_POINTS.get(rnd, 1)
            result = row[5]   # 1, 0, or None
            score  = row[6]   # pts earned or None
            picks.append({
                "game_id":  gid,
                "round":    rnd,
                "matchup":  matchup,
                "pick":     pick,
                "pts_value": pts_w,
                "result":   result,   # 1=correct, 0=wrong, None=pending
                "score":    score or 0,
            })
        # Total score = sum of earned points
        total = sum(p["score"] or 0 for p in picks)
        pts_remaining = sum(p["pts_value"] for p in picks if p["result"] is None)
        players[p] = {
            "picks":          picks,
            "total":          total,
            "pts_remaining":  pts_remaining,
            "max_possible":   total + pts_remaining,
        }

    # ── Leaderboard (sorted) ──────────────────────────────────────────────────
    leaderboard = sorted(
        [{"name": name, **data} for name, data in players.items()],
        key=lambda x: (-x["total"], x["name"])
    )
    # Assign ranks (handle ties)
    rank = 1
    for i, entry in enumerate(leaderboard):
        if i > 0 and entry["total"] == leaderboard[i-1]["total"]:
            entry["rank"] = leaderboard[i-1]["rank"]
        else:
            entry["rank"] = rank
        rank += 1

    # Champion pick per player
    for entry in leaderboard:
        champ = next(
            (p["pick"] for p in players[entry["name"]]["picks"] if p["game_id"] == "Cha_G1"),
            "—"
        )
        entry["champion_pick"] = champ

    # ── Games completed count ─────────────────────────────────────────────────
    completed = sum(1 for v in results.values() if v["winner"])
    total_games = len(results)

    return {
        "leaderboard":    leaderboard,
        "players":        players,
        "results":        results,
        "completed_games": completed,
        "total_games":    total_games,
        "last_updated":   datetime.now().strftime("%b %d, %Y %I:%M %p"),
    }


@app.route("/")
def index():
    data = load_data()
    return render_template("index.html", data=data, players=PLAYERS)


@app.route("/player/<name>")
def player(name):
    if name not in PLAYERS:
        return "Player not found", 404
    data = load_data()
    return render_template("player.html", data=data, player_name=name,
                           player=data["players"][name])


@app.route("/api/data")
def api_data():
    """JSON endpoint — useful for auto-refresh."""
    data = load_data()
    # Make serializable
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
