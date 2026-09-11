"""
LeagueBoss v2 — Multi-League Soccer Manager
============================================
Hierarchy: League ▸ Stream ▸ Team ▸ Results

New in v2:
- Leagues sit above streams. You can manage multiple leagues from one app.
- Edit any past result (date, teams, score) — standings recompute automatically.
- Multiple users with per-league permissions. Super-admin manages everything.
- Dashboard with key stats and Chart.js visualisations per league.
- Flexible Excel export: all data, single league, single stream, or single team.

Run with: python3 app.py
Open:     http://localhost:5000

Default super-admin login:
  Username: admin
  Password: safa2026   (override with ADMIN_PASSWORD env var)
"""

from flask import Flask, render_template, request, jsonify, send_file, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date, datetime
from collections import defaultdict
import json
import os
import io
import re
import uuid

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'leagueboss-dev-key')
DATA_FILE = os.environ.get('DATA_FILE', 'league_data.json')

DEFAULT_SUPER_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'safa2026')


# ════════════════════════════════════════════════════════════════════════════
#  DATA PERSISTENCE + MIGRATION FROM V1
# ════════════════════════════════════════════════════════════════════════════
def empty_data():
    """A fresh data store with the default super-admin."""
    return {
        'leagues': [],
        'results': [],
        'players': [],         # league-scoped roster: {id, leagueId, name, position, number, dob}
        'registrations': [],   # {id, playerId, teamName, leagueId, startDate, endDate?}
        'match_events': [],    # {id, resultId, type, ...} type = goal|card|sub|lineup|keeper
        'users': {
            'admin': {
                'password': DEFAULT_SUPER_PASSWORD,
                'role': 'super',
                'leagues': 'all',
                'name': 'Super Admin',
            }
        },
    }


def migrate(data):
    """Migrate older data files forward."""
    # v1 → v2: wrap streams into a default league
    if 'leagues' not in data:
        print("⚙️  Migrating data from v1 to v2 (wrapping streams in 'Default League')…")
        league_id = f"l_{uuid.uuid4().hex[:8]}_default"
        migrated = empty_data()
        migrated['leagues'].append({
            'id': league_id,
            'name': 'Default League',
            'streams': data.get('streams', []),
        })
        for r in data.get('results', []):
            r['leagueId'] = league_id
            migrated['results'].append(r)
        data = migrated

    # v2 → v2.1: add players/registrations/match_events lists if absent
    for key in ('players', 'registrations', 'match_events'):
        if key not in data:
            data[key] = []
    return data


def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f:
            data = json.load(f)
        data = migrate(data)
        # make sure 'users' exists even after migration
        if 'users' not in data:
            data['users'] = empty_data()['users']
        return data
    return empty_data()


def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2)


# ════════════════════════════════════════════════════════════════════════════
#  AUTH HELPERS
# ════════════════════════════════════════════════════════════════════════════
def current_user(data=None):
    """Return the user record for whoever is logged in, or None."""
    username = session.get('user')
    if not username:
        return None
    data = data or load_data()
    return data['users'].get(username)


def is_super():
    u = current_user()
    return bool(u and u.get('role') == 'super')


def can_edit_league(league_id, data=None):
    """True if the logged-in user may edit the given league."""
    u = current_user(data)
    if not u:
        return False
    if u.get('role') == 'super':
        return True
    leagues = u.get('leagues', [])
    return leagues == 'all' or league_id in leagues


def find_league(data, league_id):
    return next((l for l in data['leagues'] if l['id'] == league_id), None)


def find_stream(league, stream_id):
    if not league:
        return None
    return next((s for s in league['streams'] if s['id'] == stream_id), None)


# ════════════════════════════════════════════════════════════════════════════
#  STANDINGS + DASHBOARD CALCULATIONS
# ════════════════════════════════════════════════════════════════════════════
def calc_standings(league_id, stream_id, data):
    league = find_league(data, league_id)
    stream = find_stream(league, stream_id)
    if not stream:
        return []

    results = [r for r in data['results']
               if r['leagueId'] == league_id and r['streamId'] == stream_id]

    table = {team: {'team': team, 'p': 0, 'w': 0, 'd': 0, 'l': 0,
                    'gf': 0, 'ga': 0, 'gd': 0, 'pts': 0, 'form': []}
             for team in stream['teams']}

    for r in sorted(results, key=lambda r: r['date']):
        home, away = r['home'], r['away']
        if home not in table or away not in table:
            continue
        gh, ga = r['gh'], r['ga']
        h, a = table[home], table[away]
        h['p'] += 1; a['p'] += 1
        h['gf'] += gh; h['ga'] += ga
        a['gf'] += ga; a['ga'] += gh
        if gh > ga:
            h['w'] += 1; h['pts'] += 3; h['form'].append('W')
            a['l'] += 1; a['form'].append('L')
        elif gh < ga:
            a['w'] += 1; a['pts'] += 3; a['form'].append('W')
            h['l'] += 1; h['form'].append('L')
        else:
            h['d'] += 1; h['pts'] += 1; h['form'].append('D')
            a['d'] += 1; a['pts'] += 1; a['form'].append('D')

    for t in table.values():
        t['gd'] = t['gf'] - t['ga']
        t['form'] = t['form'][-5:]  # keep last 5

    return sorted(table.values(),
                  key=lambda x: (-x['pts'], -x['gd'], -x['gf'], x['team']))


def calc_team_summary(league_id, team_name, data):
    """Aggregate a single team across all streams in a league."""
    results = [r for r in data['results']
               if r['leagueId'] == league_id
               and (r['home'] == team_name or r['away'] == team_name)]
    summary = {'team': team_name, 'p': 0, 'w': 0, 'd': 0, 'l': 0,
               'gf': 0, 'ga': 0, 'gd': 0, 'pts': 0}
    matches = []
    for r in sorted(results, key=lambda r: r['date']):
        is_home = r['home'] == team_name
        own = r['gh'] if is_home else r['ga']
        opp = r['ga'] if is_home else r['gh']
        outcome = 'W' if own > opp else ('D' if own == opp else 'L')
        summary['p'] += 1
        summary['gf'] += own
        summary['ga'] += opp
        if outcome == 'W':
            summary['w'] += 1; summary['pts'] += 3
        elif outcome == 'D':
            summary['d'] += 1; summary['pts'] += 1
        else:
            summary['l'] += 1
        matches.append({
            'date': r['date'],
            'opponent': r['away'] if is_home else r['home'],
            'venue': 'H' if is_home else 'A',
            'score': f"{own}-{opp}",
            'outcome': outcome,
        })
    summary['gd'] = summary['gf'] - summary['ga']
    return summary, matches


def _standings_from_results(results, teams):
    """Compute standings for an arbitrary list of results + teams. Used by both
    the main standings endpoint and the dashboard's per-stream stats."""
    table = {t: {'team': t, 'p': 0, 'w': 0, 'd': 0, 'l': 0,
                 'gf': 0, 'ga': 0, 'gd': 0, 'pts': 0, 'form': []}
             for t in teams}
    for r in sorted(results, key=lambda r: r['date']):
        h, a = table.get(r['home']), table.get(r['away'])
        if not h or not a: continue
        gh, ga = r['gh'], r['ga']
        h['p'] += 1; a['p'] += 1
        h['gf'] += gh; h['ga'] += ga
        a['gf'] += ga; a['ga'] += gh
        if gh > ga:
            h['w'] += 1; h['pts'] += 3; h['form'].append('W')
            a['l'] += 1; a['form'].append('L')
        elif gh < ga:
            a['w'] += 1; a['pts'] += 3; a['form'].append('W')
            h['l'] += 1; h['form'].append('L')
        else:
            h['d'] += 1; h['pts'] += 1; h['form'].append('D')
            a['d'] += 1; a['pts'] += 1; a['form'].append('D')
    for t in table.values():
        t['gd'] = t['gf'] - t['ga']
    return sorted(table.values(),
                  key=lambda x: (-x['pts'], -x['gd'], -x['gf'], x['team']))


def _filter_by_date(results, date_from=None, date_to=None):
    """Filter a list of results by an inclusive date range (YYYY-MM-DD)."""
    if date_from:
        results = [r for r in results if r['date'] >= date_from]
    if date_to:
        results = [r for r in results if r['date'] <= date_to]
    return results


def calc_dashboard(league_id, data, date_from=None, date_to=None):
    """Comprehensive league dashboard stats. Optional date filtering."""
    league = find_league(data, league_id)
    if not league:
        return {}

    all_league_results = [r for r in data['results'] if r['leagueId'] == league_id]
    results = _filter_by_date(all_league_results, date_from, date_to)
    streams = league['streams']
    all_teams = sorted({t for s in streams for t in s['teams']})

    # Available date range (across ALL results, ignoring filter)
    all_dates = sorted({r['date'] for r in all_league_results})
    date_range = {
        'from': date_from,
        'to': date_to,
        'min_available': all_dates[0] if all_dates else None,
        'max_available': all_dates[-1] if all_dates else None,
    }

    # ─── BASIC TOTALS ───────────────────────────────────────────────
    total_matches = len(results)
    total_goals = sum(r['gh'] + r['ga'] for r in results)
    home_wins = sum(1 for r in results if r['gh'] > r['ga'])
    away_wins = sum(1 for r in results if r['ga'] > r['gh'])
    draws = sum(1 for r in results if r['gh'] == r['ga'])

    biggest_win = None
    if results:
        biggest = max(results, key=lambda r: abs(r['gh'] - r['ga']))
        if abs(biggest['gh'] - biggest['ga']) > 0:
            biggest_win = {
                'date': biggest['date'],
                'home': biggest['home'],
                'away': biggest['away'],
                'score': f"{biggest['gh']}-{biggest['ga']}",
                'margin': abs(biggest['gh'] - biggest['ga']),
            }

    # ─── GOALS FOR/AGAINST + CLEAN SHEETS ──────────────────────────
    goals_for = defaultdict(int)
    goals_against = defaultdict(int)
    clean_sheets = defaultdict(int)
    matches_played = defaultdict(int)
    for r in results:
        goals_for[r['home']] += r['gh']
        goals_for[r['away']] += r['ga']
        goals_against[r['home']] += r['ga']
        goals_against[r['away']] += r['gh']
        matches_played[r['home']] += 1
        matches_played[r['away']] += 1
        if r['ga'] == 0:
            clean_sheets[r['home']] += 1
        if r['gh'] == 0:
            clean_sheets[r['away']] += 1

    top_scorers = sorted(goals_for.items(), key=lambda x: -x[1])[:10]
    best_defence = sorted(
        [(t, goals_against[t]) for t in all_teams if matches_played[t] > 0],
        key=lambda x: x[1]
    )[:10]
    most_clean_sheets = sorted(
        [(t, clean_sheets[t]) for t in all_teams if clean_sheets[t] > 0],
        key=lambda x: -x[1]
    )[:10]
    worst_conceding = sorted(
        [(t, goals_against[t]) for t in all_teams if matches_played[t] > 0],
        key=lambda x: -x[1]
    )[:10]

    # ─── TIMELINE ──────────────────────────────────────────────────
    goals_by_date = defaultdict(int)
    matches_by_date = defaultdict(int)
    for r in results:
        goals_by_date[r['date']] += r['gh'] + r['ga']
        matches_by_date[r['date']] += 1
    timeline_dates = sorted(goals_by_date.keys())

    # ─── TITLE RACE + STREAM COMPARISON ───────────────────────────
    title_race = []
    stream_comparison = []
    for s in streams:
        s_results = [r for r in results if r['streamId'] == s['id']]
        s_standings = _standings_from_results(s_results, s['teams'])
        s_total_goals = sum(r['gh'] + r['ga'] for r in s_results)

        title_race.append({
            'stream': s['name'],
            'leader': s_standings[0]['team'] if s_standings else None,
            'leader_pts': s_standings[0]['pts'] if s_standings else 0,
            'second': s_standings[1]['team'] if len(s_standings) > 1 else None,
            'second_pts': s_standings[1]['pts'] if len(s_standings) > 1 else 0,
            'gap': (s_standings[0]['pts'] - s_standings[1]['pts'])
                   if len(s_standings) > 1 else 0,
            'teams_count': len(s['teams']),
            'matches_played': len(s_results),
        })
        stream_comparison.append({
            'stream': s['name'],
            'teams': len(s['teams']),
            'matches': len(s_results),
            'goals': s_total_goals,
            'avg_goals': round(s_total_goals / len(s_results), 2) if s_results else 0,
            'leader': s_standings[0]['team'] if s_standings else '—',
            'gap_to_2nd': (s_standings[0]['pts'] - s_standings[1]['pts'])
                          if len(s_standings) > 1 else 0,
        })
    # sort title race by closest race (smallest gap first)
    title_race.sort(key=lambda x: (x['gap'], -x['matches_played']))
    stream_comparison.sort(key=lambda x: -x['matches'])

    # ─── FORM TABLE (last 5 matches per team, across all streams) ─
    team_recent = defaultdict(list)
    for r in sorted(results, key=lambda r: r['date']):
        for team, gf, ga in ((r['home'], r['gh'], r['ga']),
                              (r['away'], r['ga'], r['gh'])):
            outcome = 'W' if gf > ga else ('D' if gf == ga else 'L')
            team_recent[team].append({'date': r['date'], 'outcome': outcome,
                                       'gf': gf, 'ga': ga})

    form_table = []
    for team in all_teams:
        last5 = team_recent[team][-5:] if team_recent.get(team) else []
        if not last5:
            continue
        pts = sum(3 if m['outcome']=='W' else (1 if m['outcome']=='D' else 0)
                  for m in last5)
        form_table.append({
            'team': team,
            'matches': len(last5),
            'points': pts,
            'form': [m['outcome'] for m in last5],
            'gf': sum(m['gf'] for m in last5),
            'ga': sum(m['ga'] for m in last5),
        })
    form_table.sort(key=lambda x: (-x['points'], -(x['gf']-x['ga']), -x['gf']))
    form_table = form_table[:10]

    # ─── SCORELINE DISTRIBUTION ───────────────────────────────────
    scoreline_counts = defaultdict(int)
    for r in results:
        # normalise so 2-1 and 1-2 group together as "the 2-1 scoreline"
        high, low = max(r['gh'], r['ga']), min(r['gh'], r['ga'])
        scoreline_counts[f"{high}-{low}"] += 1
    scorelines = sorted(scoreline_counts.items(), key=lambda x: -x[1])[:10]

    return {
        'league_name': league['name'],
        'date_range': date_range,
        'totals': {
            'streams': len(streams),
            'teams': len(all_teams),
            'matches': total_matches,
            'goals': total_goals,
            'avg_goals': round(total_goals / total_matches, 2) if total_matches else 0,
        },
        'outcomes': {
            'home_wins': home_wins,
            'away_wins': away_wins,
            'draws': draws,
        },
        'biggest_win': biggest_win,
        'top_scorers': top_scorers,
        'best_defence': best_defence,
        'most_clean_sheets': most_clean_sheets,
        'worst_conceding': worst_conceding,
        'timeline': {
            'dates': timeline_dates,
            'goals': [goals_by_date[d] for d in timeline_dates],
            'matches': [matches_by_date[d] for d in timeline_dates],
        },
        'title_race': title_race,
        'stream_comparison': stream_comparison,
        'form_table': form_table,
        'scorelines': scorelines,
    }


def calc_team_dashboard(league_id, team, data, date_from=None, date_to=None):
    """Deep stats for a single team within a league."""
    league = find_league(data, league_id)
    if not league:
        return None

    # Which stream does this team belong to?
    team_stream = next((s for s in league['streams'] if team in s['teams']), None)

    all_results = [r for r in data['results']
                   if r['leagueId'] == league_id
                   and (r['home'] == team or r['away'] == team)]
    results = _filter_by_date(all_results, date_from, date_to)

    all_dates = sorted({r['date'] for r in all_results})
    date_range = {
        'from': date_from,
        'to': date_to,
        'min_available': all_dates[0] if all_dates else None,
        'max_available': all_dates[-1] if all_dates else None,
    }

    # Overall record
    summary = {'p': 0, 'w': 0, 'd': 0, 'l': 0, 'gf': 0, 'ga': 0, 'gd': 0, 'pts': 0,
               'home_w': 0, 'home_d': 0, 'home_l': 0,
               'away_w': 0, 'away_d': 0, 'away_l': 0,
               'home_gf': 0, 'home_ga': 0, 'away_gf': 0, 'away_ga': 0,
               'clean_sheets': 0, 'failed_to_score': 0,
               'biggest_win': None, 'worst_loss': None}

    matches = []  # full match history
    opponents = defaultdict(lambda: {'p': 0, 'w': 0, 'd': 0, 'l': 0,
                                       'gf': 0, 'ga': 0})
    timeline_gf = defaultdict(int)
    timeline_ga = defaultdict(int)
    timeline_pts = defaultdict(int)

    biggest_win_margin = 0
    worst_loss_margin = 0

    for r in sorted(results, key=lambda r: r['date']):
        is_home = r['home'] == team
        own = r['gh'] if is_home else r['ga']
        opp = r['ga'] if is_home else r['gh']
        opponent = r['away'] if is_home else r['home']
        outcome = 'W' if own > opp else ('D' if own == opp else 'L')

        summary['p'] += 1
        summary['gf'] += own
        summary['ga'] += opp
        if own == 0: summary['failed_to_score'] += 1
        if opp == 0: summary['clean_sheets'] += 1

        if is_home:
            summary['home_gf'] += own; summary['home_ga'] += opp
            if outcome == 'W': summary['home_w'] += 1
            elif outcome == 'D': summary['home_d'] += 1
            else: summary['home_l'] += 1
        else:
            summary['away_gf'] += own; summary['away_ga'] += opp
            if outcome == 'W': summary['away_w'] += 1
            elif outcome == 'D': summary['away_d'] += 1
            else: summary['away_l'] += 1

        if outcome == 'W':
            summary['w'] += 1; summary['pts'] += 3
            if own - opp > biggest_win_margin:
                biggest_win_margin = own - opp
                summary['biggest_win'] = {
                    'opponent': opponent, 'venue': 'H' if is_home else 'A',
                    'score': f"{own}-{opp}", 'date': r['date'],
                }
        elif outcome == 'D':
            summary['d'] += 1; summary['pts'] += 1
        else:
            summary['l'] += 1
            if opp - own > worst_loss_margin:
                worst_loss_margin = opp - own
                summary['worst_loss'] = {
                    'opponent': opponent, 'venue': 'H' if is_home else 'A',
                    'score': f"{own}-{opp}", 'date': r['date'],
                }

        matches.append({
            'date': r['date'],
            'opponent': opponent,
            'venue': 'H' if is_home else 'A',
            'score': f"{own}-{opp}",
            'gf': own, 'ga': opp,
            'outcome': outcome,
        })

        # Opponent breakdown
        op = opponents[opponent]
        op['p'] += 1; op['gf'] += own; op['ga'] += opp
        if outcome == 'W': op['w'] += 1
        elif outcome == 'D': op['d'] += 1
        else: op['l'] += 1

        # Timeline
        timeline_gf[r['date']] += own
        timeline_ga[r['date']] += opp
        timeline_pts[r['date']] += 3 if outcome == 'W' else (1 if outcome == 'D' else 0)

    summary['gd'] = summary['gf'] - summary['ga']
    summary['avg_gf'] = round(summary['gf'] / summary['p'], 2) if summary['p'] else 0
    summary['avg_ga'] = round(summary['ga'] / summary['p'], 2) if summary['p'] else 0
    summary['win_rate'] = round(100 * summary['w'] / summary['p'], 1) if summary['p'] else 0

    # Form last 5
    form = [m['outcome'] for m in matches[-5:]]

    # Standings position in their stream
    position = None
    if team_stream:
        stream_results = _filter_by_date(
            [r for r in data['results']
             if r['leagueId'] == league_id and r['streamId'] == team_stream['id']],
            date_from, date_to
        )
        standings = _standings_from_results(stream_results, team_stream['teams'])
        for i, row in enumerate(standings, 1):
            if row['team'] == team:
                position = i
                break

    # Top opponents by record
    opponents_list = []
    for op_name, stats in opponents.items():
        opponents_list.append({
            'opponent': op_name, **stats,
            'pts': stats['w'] * 3 + stats['d'],
        })
    opponents_list.sort(key=lambda x: (-x['pts'], -(x['gf']-x['ga'])))

    # Timeline arrays
    timeline_dates = sorted(timeline_gf.keys())
    cumulative_pts = []
    running = 0
    for d in timeline_dates:
        running += timeline_pts[d]
        cumulative_pts.append(running)

    return {
        'team': team,
        'league_name': league['name'],
        'stream': team_stream['name'] if team_stream else None,
        'position': position,
        'date_range': date_range,
        'summary': summary,
        'form': form,
        'matches': matches,
        'opponents': opponents_list,
        'timeline': {
            'dates': timeline_dates,
            'goals_for': [timeline_gf[d] for d in timeline_dates],
            'goals_against': [timeline_ga[d] for d in timeline_dates],
            'cumulative_pts': cumulative_pts,
        },
    }


# ════════════════════════════════════════════════════════════════════════════
#  AUTH ROUTES
# ════════════════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/login', methods=['POST'])
def login():
    body = request.json or {}
    username = (body.get('username') or '').strip().lower()
    password = body.get('password') or ''
    data = load_data()
    user = data['users'].get(username)
    if user and user['password'] == password:
        session['user'] = username
        return jsonify({
            'ok': True,
            'username': username,
            'name': user.get('name', username),
            'role': user['role'],
            'leagues': user.get('leagues', []),
        })
    return jsonify({'error': 'Incorrect username or password'}), 401


@app.route('/api/logout', methods=['POST'])
def logout():
    session.pop('user', None)
    return jsonify({'ok': True})


@app.route('/api/me')
def me():
    u = current_user()
    if not u:
        return jsonify({'logged_in': False})
    return jsonify({
        'logged_in': True,
        'username': session.get('user'),
        'name': u.get('name'),
        'role': u['role'],
        'leagues': u.get('leagues', []),
    })


# ════════════════════════════════════════════════════════════════════════════
#  USER MANAGEMENT (super-admin only)
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/users', methods=['GET'])
def list_users():
    if not is_super():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    safe = [{'username': u, 'name': info.get('name', u),
             'role': info['role'], 'leagues': info.get('leagues', [])}
            for u, info in data['users'].items()]
    return jsonify(safe)


@app.route('/api/user', methods=['POST'])
def add_user():
    if not is_super():
        return jsonify({'error': 'Only the super-admin can add users'}), 403
    body = request.json or {}
    username = (body.get('username') or '').strip().lower()
    password = body.get('password') or ''
    name = body.get('name') or username
    leagues = body.get('leagues', [])
    if not username or not password:
        return jsonify({'error': 'Username and password required'}), 400
    data = load_data()
    if username in data['users']:
        return jsonify({'error': 'User already exists'}), 400
    data['users'][username] = {
        'password': password,
        'role': 'league_admin',
        'leagues': leagues,
        'name': name,
    }
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/user/<username>', methods=['DELETE'])
def delete_user(username):
    if not is_super():
        return jsonify({'error': 'Unauthorised'}), 403
    if username == 'admin':
        return jsonify({'error': "Can't delete the default super-admin"}), 400
    data = load_data()
    data['users'].pop(username, None)
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/user/<username>', methods=['PUT'])
def update_user(username):
    if not is_super():
        return jsonify({'error': 'Unauthorised'}), 403
    body = request.json or {}
    data = load_data()
    if username not in data['users']:
        return jsonify({'error': 'User not found'}), 404
    u = data['users'][username]
    if 'password' in body and body['password']:
        u['password'] = body['password']
    if 'leagues' in body:
        u['leagues'] = body['leagues']
    if 'name' in body:
        u['name'] = body['name']
    save_data(data)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════════════════════
#  DATA ENDPOINT
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/data')
def get_data():
    """Return the full data set, filtered to leagues the user can access."""
    data = load_data()
    u = current_user(data)
    visible = data['leagues']
    if u and u['role'] != 'super':
        allowed = u.get('leagues', [])
        if allowed != 'all':
            visible = [l for l in data['leagues'] if l['id'] in allowed]
    return jsonify({
        'leagues': visible,
        'results': data['results'],
        'me': {
            'logged_in': bool(u),
            'username': session.get('user'),
            'name': u.get('name') if u else None,
            'role': u['role'] if u else None,
            'leagues': u.get('leagues', []) if u else [],
        }
    })


# ════════════════════════════════════════════════════════════════════════════
#  LEAGUE CRUD  (super-admin)
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/league', methods=['POST'])
def add_league():
    if not is_super():
        return jsonify({'error': 'Only the super-admin can create leagues'}), 403
    body = request.json or {}
    name = (body.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400
    data = load_data()
    league_id = f"l_{uuid.uuid4().hex[:8]}_{re.sub(r'[^a-z0-9]+', '_', name.lower())[:24]}"
    data['leagues'].append({'id': league_id, 'name': name, 'streams': []})
    save_data(data)
    return jsonify({'ok': True, 'id': league_id})


@app.route('/api/league/<league_id>', methods=['PUT'])
def rename_league(league_id):
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    body = request.json or {}
    name = (body.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400
    data = load_data()
    league = find_league(data, league_id)
    if not league:
        return jsonify({'error': 'Not found'}), 404
    league['name'] = name
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/league/<league_id>', methods=['DELETE'])
def delete_league(league_id):
    if not is_super():
        return jsonify({'error': 'Only the super-admin can delete leagues'}), 403
    data = load_data()
    data['leagues'] = [l for l in data['leagues'] if l['id'] != league_id]
    data['results'] = [r for r in data['results'] if r.get('leagueId') != league_id]
    save_data(data)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════════════════════
#  STREAM CRUD
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/stream', methods=['POST'])
def add_stream():
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    name = (body.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400
    data = load_data()
    league = find_league(data, league_id)
    if not league:
        return jsonify({'error': 'League not found'}), 404
    stream_id = f"s_{uuid.uuid4().hex[:8]}_{re.sub(r'[^a-z0-9]+', '_', name.lower())[:24]}"
    league['streams'].append({'id': stream_id, 'name': name, 'teams': []})
    save_data(data)
    return jsonify({'ok': True, 'id': stream_id})


@app.route('/api/stream/<league_id>/<stream_id>', methods=['DELETE'])
def delete_stream(league_id, stream_id):
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    league = find_league(data, league_id)
    if not league:
        return jsonify({'error': 'Not found'}), 404
    league['streams'] = [s for s in league['streams'] if s['id'] != stream_id]
    data['results'] = [r for r in data['results']
                       if not (r['leagueId'] == league_id and r['streamId'] == stream_id)]
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/stream/<league_id>/<stream_id>', methods=['PUT'])
def rename_stream(league_id, stream_id):
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    body = request.json or {}
    name = (body.get('name') or '').strip()
    data = load_data()
    league = find_league(data, league_id)
    stream = find_stream(league, stream_id)
    if not stream:
        return jsonify({'error': 'Not found'}), 404
    stream['name'] = name
    save_data(data)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════════════════════
#  TEAM CRUD
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/team', methods=['POST'])
def add_team():
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    league = find_league(data, league_id)
    stream = find_stream(league, body.get('streamId'))
    name = (body.get('name') or '').strip()
    if not stream:
        return jsonify({'error': 'Stream not found'}), 404
    if name in stream['teams']:
        return jsonify({'error': 'Team already in stream'}), 400
    stream['teams'].append(name)
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/team', methods=['DELETE'])
def delete_team():
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    league = find_league(data, league_id)
    stream = find_stream(league, body.get('streamId'))
    name = body.get('name')
    if not stream:
        return jsonify({'error': 'Stream not found'}), 404
    stream['teams'] = [t for t in stream['teams'] if t != name]
    data['results'] = [r for r in data['results']
                       if not (r['leagueId'] == league_id
                               and r['streamId'] == body.get('streamId')
                               and (r['home'] == name or r['away'] == name))]
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/team/rename', methods=['POST'])
def rename_team():
    """Rename a team — also updates every result referencing it."""
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    league = find_league(data, league_id)
    stream = find_stream(league, body.get('streamId'))
    old, new = body.get('oldName'), (body.get('newName') or '').strip()
    if not stream or not new:
        return jsonify({'error': 'Bad request'}), 400
    stream['teams'] = [new if t == old else t for t in stream['teams']]
    for r in data['results']:
        if r['leagueId'] == league_id and r['streamId'] == body.get('streamId'):
            if r['home'] == old: r['home'] = new
            if r['away'] == old: r['away'] = new
    save_data(data)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════════════════════
#  RESULT CRUD  (create / edit / delete)
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/result', methods=['POST'])
def add_result():
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    result = {
        'id': str(uuid.uuid4())[:8],
        'leagueId': league_id,
        'streamId': body['streamId'],
        'date': body['date'],
        'home': body['home'],
        'away': body['away'],
        'gh': int(body['gh']),
        'ga': int(body['ga']),
    }
    data['results'].append(result)
    save_data(data)
    return jsonify({'ok': True, 'id': result['id']})


@app.route('/api/result/<result_id>', methods=['PUT'])
def edit_result(result_id):
    """Edit any field of any past result — standings recompute from data."""
    body = request.json or {}
    data = load_data()
    r = next((x for x in data['results'] if x['id'] == result_id), None)
    if not r:
        return jsonify({'error': 'Result not found'}), 404
    if not can_edit_league(r['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    for field in ('date', 'home', 'away'):
        if field in body:
            r[field] = body[field]
    if 'gh' in body:
        r['gh'] = int(body['gh'])
    if 'ga' in body:
        r['ga'] = int(body['ga'])
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/result/<result_id>', methods=['DELETE'])
def delete_result(result_id):
    data = load_data()
    r = next((x for x in data['results'] if x['id'] == result_id), None)
    if not r:
        return jsonify({'error': 'Result not found'}), 404
    if not can_edit_league(r['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    data['results'] = [x for x in data['results'] if x['id'] != result_id]
    save_data(data)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════════════════════
#  STANDINGS + DASHBOARD ENDPOINTS
# ════════════════════════════════════════════════════════════════════════════
@app.route('/api/standings/<league_id>/<stream_id>')
def get_standings(league_id, stream_id):
    return jsonify(calc_standings(league_id, stream_id, load_data()))


@app.route('/api/dashboard/<league_id>')
def dashboard(league_id):
    date_from = request.args.get('from') or None
    date_to = request.args.get('to') or None
    return jsonify(calc_dashboard(league_id, load_data(), date_from, date_to))


@app.route('/api/team-summary/<league_id>/<team_name>')
def team_summary(league_id, team_name):
    summary, matches = calc_team_summary(league_id, team_name, load_data())
    return jsonify({'summary': summary, 'matches': matches})


@app.route('/api/team-dashboard/<league_id>/<team_name>')
def team_dashboard(league_id, team_name):
    date_from = request.args.get('from') or None
    date_to = request.args.get('to') or None
    result = calc_team_dashboard(league_id, team_name, load_data(), date_from, date_to)
    if result is None:
        return jsonify({'error': 'League not found'}), 404
    return jsonify(result)


# ════════════════════════════════════════════════════════════════════════════
#  PLAYERS, REGISTRATIONS, MATCH EVENTS
# ════════════════════════════════════════════════════════════════════════════
POSITIONS = ('GK', 'DEF', 'MID', 'FWD')


def _active_team_for(player_id, data, on_date=None):
    """Returns the team name the player is currently registered with (or on a specific date)."""
    on_date = on_date or date.today().isoformat()
    regs = [r for r in data['registrations'] if r['playerId'] == player_id]
    for r in regs:
        if r['startDate'] <= on_date and (not r.get('endDate') or r['endDate'] >= on_date):
            return r['teamName']
    return None


def _player_with_team(player, data):
    """Enrich a player dict with currentTeam."""
    p = dict(player)
    p['currentTeam'] = _active_team_for(player['id'], data)
    return p


# ──────────── PLAYERS CRUD ──────────────────────────────────────────────────
@app.route('/api/players/<league_id>')
def list_players(league_id):
    data = load_data()
    players = [p for p in data['players'] if p['leagueId'] == league_id]
    return jsonify([_player_with_team(p, data) for p in players])


@app.route('/api/player', methods=['POST'])
def add_player():
    body = request.json or {}
    league_id = body.get('leagueId')
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    name = (body.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400
    data = load_data()
    player = {
        'id': f"p_{uuid.uuid4().hex[:10]}",
        'leagueId': league_id,
        'name': name,
        'position': body.get('position') if body.get('position') in POSITIONS else None,
        'number': body.get('number'),
        'dob': body.get('dob') or None,
    }
    data['players'].append(player)

    # Optional immediate registration
    team_name = (body.get('teamName') or '').strip()
    start_date = body.get('startDate') or date.today().isoformat()
    if team_name:
        # Enforce one-team-per-league rule
        active = _active_team_for(player['id'], data, start_date)
        if active:
            return jsonify({'error': f'Player already at {active}'}), 400
        data['registrations'].append({
            'id': f"reg_{uuid.uuid4().hex[:8]}",
            'playerId': player['id'],
            'leagueId': league_id,
            'teamName': team_name,
            'startDate': start_date,
            'endDate': None,
        })
    save_data(data)
    return jsonify({'ok': True, 'id': player['id']})


@app.route('/api/player/<player_id>', methods=['PUT'])
def update_player(player_id):
    data = load_data()
    p = next((x for x in data['players'] if x['id'] == player_id), None)
    if not p: return jsonify({'error': 'Not found'}), 404
    if not can_edit_league(p['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    body = request.json or {}
    for f in ('name', 'number', 'dob'):
        if f in body: p[f] = body[f] or None
    if 'position' in body:
        p['position'] = body['position'] if body['position'] in POSITIONS else None
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/player/<player_id>', methods=['DELETE'])
def delete_player(player_id):
    data = load_data()
    p = next((x for x in data['players'] if x['id'] == player_id), None)
    if not p: return jsonify({'error': 'Not found'}), 404
    if not can_edit_league(p['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    data['players'] = [x for x in data['players'] if x['id'] != player_id]
    data['registrations'] = [r for r in data['registrations'] if r['playerId'] != player_id]
    # Also remove any match events referencing this player
    data['match_events'] = [e for e in data['match_events']
                             if e.get('playerId') != player_id
                             and e.get('assistPlayerId') != player_id]
    save_data(data)
    return jsonify({'ok': True})


# ──────────── REGISTRATIONS (transfers) ─────────────────────────────────────
@app.route('/api/player/<player_id>/registrations')
def player_registrations(player_id):
    data = load_data()
    regs = [r for r in data['registrations'] if r['playerId'] == player_id]
    regs.sort(key=lambda r: r['startDate'])
    return jsonify(regs)


@app.route('/api/registration', methods=['POST'])
def add_registration():
    """Transfer a player to a (new) team — closes any open registration first."""
    body = request.json or {}
    player_id = body.get('playerId')
    team_name = (body.get('teamName') or '').strip()
    start = body.get('startDate') or date.today().isoformat()
    data = load_data()
    p = next((x for x in data['players'] if x['id'] == player_id), None)
    if not p: return jsonify({'error': 'Player not found'}), 404
    if not can_edit_league(p['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    if not team_name:
        return jsonify({'error': 'Team required'}), 400

    # Close any open registration before start date
    for r in data['registrations']:
        if r['playerId'] == player_id and not r.get('endDate'):
            # End it the day before the new start
            from datetime import datetime as _dt, timedelta as _td
            try:
                d_end = (_dt.strptime(start, '%Y-%m-%d') - _td(days=1)).strftime('%Y-%m-%d')
            except ValueError:
                d_end = start
            r['endDate'] = d_end

    data['registrations'].append({
        'id': f"reg_{uuid.uuid4().hex[:8]}",
        'playerId': player_id,
        'leagueId': p['leagueId'],
        'teamName': team_name,
        'startDate': start,
        'endDate': None,
    })
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/registration/<reg_id>', methods=['DELETE'])
def delete_registration(reg_id):
    data = load_data()
    r = next((x for x in data['registrations'] if x['id'] == reg_id), None)
    if not r: return jsonify({'error': 'Not found'}), 404
    if not can_edit_league(r['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403
    data['registrations'] = [x for x in data['registrations'] if x['id'] != reg_id]
    save_data(data)
    return jsonify({'ok': True})


# ──────────── ROSTER (active players for a team) ────────────────────────────
@app.route('/api/roster/<league_id>/<team_name>')
def get_roster(league_id, team_name):
    """Return all players active for this team. Optional ?date=YYYY-MM-DD."""
    on_date = request.args.get('date') or None
    data = load_data()
    out = []
    for p in data['players']:
        if p['leagueId'] != league_id: continue
        active_team = _active_team_for(p['id'], data, on_date)
        if active_team == team_name:
            out.append(_player_with_team(p, data))
    out.sort(key=lambda x: x['name'])
    return jsonify(out)


# ──────────── MATCH EVENTS  (per result) ────────────────────────────────────
@app.route('/api/match-events/<result_id>')
def list_match_events(result_id):
    data = load_data()
    events = [e for e in data['match_events'] if e['resultId'] == result_id]
    # Enrich with player names
    players_by_id = {p['id']: p['name'] for p in data['players']}
    for e in events:
        if e.get('playerId'):
            e['playerName'] = players_by_id.get(e['playerId'], '(unknown)')
        if e.get('assistPlayerId'):
            e['assistPlayerName'] = players_by_id.get(e['assistPlayerId'], '(unknown)')
    return jsonify(events)


@app.route('/api/match-events/<result_id>', methods=['PUT'])
def save_match_events(result_id):
    """Replace the entire set of events for one result. Body: {events: [...]}.

    Each event has a 'type' field: goal | card | sub | lineup | keeper.
    - goal:    {type, team, playerId, assistPlayerId?, minute?}
    - card:    {type, team, playerId, color: 'Y'|'R', minute?}
    - sub:     {type, team, offPlayerId, onPlayerId, minute?}
    - lineup:  {type, team, playerId, role: 'start'|'bench'}
    - keeper:  {type, team, playerId}  (override default GK detection)
    """
    body = request.json or {}
    data = load_data()
    r = next((x for x in data['results'] if x['id'] == result_id), None)
    if not r: return jsonify({'error': 'Result not found'}), 404
    if not can_edit_league(r['leagueId']):
        return jsonify({'error': 'Unauthorised'}), 403

    new_events = []
    for e in body.get('events', []):
        ev = {
            'id': e.get('id') or f"ev_{uuid.uuid4().hex[:10]}",
            'resultId': result_id,
            'leagueId': r['leagueId'],
            'type': e.get('type'),
            'team': e.get('team'),
        }
        for f in ('playerId', 'assistPlayerId', 'offPlayerId', 'onPlayerId',
                  'color', 'role', 'minute'):
            if f in e:
                ev[f] = e[f]
        new_events.append(ev)

    # Remove old events for this result, then append new
    data['match_events'] = [e for e in data['match_events'] if e['resultId'] != result_id]
    data['match_events'].extend(new_events)
    save_data(data)
    return jsonify({'ok': True, 'count': len(new_events)})


# ──────────── PLAYER STATS + RANKINGS ───────────────────────────────────────
def calc_player_stats(league_id, data, date_from=None, date_to=None):
    """Aggregate all player stats across the league within optional date range."""
    league = find_league(data, league_id)
    if not league: return {}

    # Result lookup
    results_in_league = [r for r in data['results'] if r['leagueId'] == league_id]
    results_in_league = _filter_by_date(results_in_league, date_from, date_to)
    result_by_id = {r['id']: r for r in results_in_league}

    # Initialise stats per player
    stats = {p['id']: {
        'playerId': p['id'], 'name': p['name'], 'position': p['position'],
        'number': p['number'], 'team': _active_team_for(p['id'], data),
        'goals': 0, 'assists': 0, 'g_plus_a': 0,
        'yellows': 0, 'reds': 0,
        'minutes': 0, 'matches': 0,
        'clean_sheets': 0,
    } for p in data['players'] if p['leagueId'] == league_id}

    # Walk events
    keeper_by_match_team = {}  # (resultId, team) -> playerId (override)
    lineups_by_match = defaultdict(lambda: defaultdict(list))  # resultId -> team -> [(playerId,role)]
    subs_by_match = defaultdict(list)  # resultId -> [sub events]

    for e in data['match_events']:
        if e['resultId'] not in result_by_id: continue
        if e['type'] == 'goal' and e.get('playerId') in stats:
            stats[e['playerId']]['goals'] += 1
            if e.get('assistPlayerId') in stats:
                stats[e['assistPlayerId']]['assists'] += 1
        elif e['type'] == 'card' and e.get('playerId') in stats:
            if e.get('color') == 'Y': stats[e['playerId']]['yellows'] += 1
            elif e.get('color') == 'R': stats[e['playerId']]['reds'] += 1
        elif e['type'] == 'lineup':
            lineups_by_match[e['resultId']][e.get('team')].append((e.get('playerId'), e.get('role')))
        elif e['type'] == 'sub':
            subs_by_match[e['resultId']].append(e)
        elif e['type'] == 'keeper':
            keeper_by_match_team[(e['resultId'], e.get('team'))] = e.get('playerId')

    # For each match where we have lineups, compute MP + minutes
    for result_id, by_team in lineups_by_match.items():
        r = result_by_id.get(result_id)
        if not r: continue
        for team_name, entries in by_team.items():
            starters = {pid for pid, role in entries if role == 'start' and pid in stats}
            bench = {pid for pid, role in entries if role == 'bench' and pid in stats}

            # Minutes: starters default 90, subs default 0; adjust by sub events
            minutes = {pid: 90 for pid in starters}
            minutes.update({pid: 0 for pid in bench})
            for sub in subs_by_match.get(result_id, []):
                if sub.get('team') != team_name: continue
                minute = sub.get('minute')
                try:
                    minute = int(minute) if minute is not None else None
                except (TypeError, ValueError):
                    minute = None
                if minute is None: continue
                off_id, on_id = sub.get('offPlayerId'), sub.get('onPlayerId')
                if off_id in minutes:
                    minutes[off_id] = min(minutes[off_id], minute)
                if on_id in minutes:
                    minutes[on_id] = max(0, 90 - minute)

            for pid in starters | bench:
                if pid in stats and minutes.get(pid, 0) > 0:
                    stats[pid]['matches'] += 1
                    stats[pid]['minutes'] += minutes[pid]

            # Clean sheet: pick keeper (override → GK by position → first starter)
            keeper_id = keeper_by_match_team.get((result_id, team_name))
            if not keeper_id:
                gks = [pid for pid in starters
                       if pid in stats and stats[pid]['position'] == 'GK']
                keeper_id = gks[0] if gks else next(iter(starters), None)
            if keeper_id and keeper_id in stats:
                if r['home'] == team_name and r['ga'] == 0:
                    stats[keeper_id]['clean_sheets'] += 1
                elif r['away'] == team_name and r['gh'] == 0:
                    stats[keeper_id]['clean_sheets'] += 1

    # Final compute
    for s in stats.values():
        s['g_plus_a'] = s['goals'] + s['assists']

    return list(stats.values())


@app.route('/api/player-stats/<league_id>')
def player_stats(league_id):
    date_from = request.args.get('from') or None
    date_to = request.args.get('to') or None
    return jsonify(calc_player_stats(league_id, load_data(), date_from, date_to))


@app.route('/api/player-dashboard/<player_id>')
def player_dashboard(player_id):
    """Detailed stats for one player: career, per-match log, rankings."""
    data = load_data()
    p = next((x for x in data['players'] if x['id'] == player_id), None)
    if not p: return jsonify({'error': 'Not found'}), 404
    league_id = p['leagueId']
    league = find_league(data, league_id)

    all_stats = calc_player_stats(league_id, data)
    me = next((s for s in all_stats if s['playerId'] == player_id), None)

    # Rankings: where does this player sit among all in the league?
    def rank_in(metric):
        sorted_list = sorted(all_stats, key=lambda x: -x[metric])
        for i, s in enumerate(sorted_list, 1):
            if s['playerId'] == player_id:
                return {'rank': i, 'total': len(sorted_list), 'value': s[metric]}
        return None

    rankings = {
        'goals': rank_in('goals'),
        'assists': rank_in('assists'),
        'g_plus_a': rank_in('g_plus_a'),
        'minutes': rank_in('minutes'),
        'matches': rank_in('matches'),
        'clean_sheets': rank_in('clean_sheets'),
    }

    # Match log
    results_in_league = {r['id']: r for r in data['results'] if r['leagueId'] == league_id}
    match_log = []
    events_for_player = [e for e in data['match_events']
                          if e['resultId'] in results_in_league
                          and (e.get('playerId') == player_id
                                or e.get('assistPlayerId') == player_id
                                or e.get('offPlayerId') == player_id
                                or e.get('onPlayerId') == player_id)]
    # Group by match
    by_match = defaultdict(list)
    for e in events_for_player:
        by_match[e['resultId']].append(e)
    for result_id, evs in by_match.items():
        r = results_in_league[result_id]
        goals_scored = sum(1 for e in evs if e['type'] == 'goal' and e.get('playerId') == player_id)
        assists = sum(1 for e in evs if e['type'] == 'goal' and e.get('assistPlayerId') == player_id)
        yellows = sum(1 for e in evs if e['type'] == 'card' and e.get('playerId') == player_id and e.get('color') == 'Y')
        reds = sum(1 for e in evs if e['type'] == 'card' and e.get('playerId') == player_id and e.get('color') == 'R')
        match_log.append({
            'date': r['date'],
            'home': r['home'], 'away': r['away'],
            'score': f"{r['gh']}-{r['ga']}",
            'goals': goals_scored, 'assists': assists,
            'yellows': yellows, 'reds': reds,
        })
    match_log.sort(key=lambda x: x['date'], reverse=True)

    # Teammates comparison: current team only
    team = me['team'] if me else None
    teammates = [s for s in all_stats if s['team'] == team and s['playerId'] != player_id] if team else []
    teammates.sort(key=lambda x: -x['g_plus_a'])

    return jsonify({
        'player': p,
        'team': team,
        'league_name': league['name'] if league else None,
        'stats': me,
        'rankings': rankings,
        'match_log': match_log,
        'teammates': teammates[:10],
        'registrations': sorted(
            [r for r in data['registrations'] if r['playerId'] == player_id],
            key=lambda r: r['startDate']
        ),
    })


# ════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT — shared styles + helpers
# ════════════════════════════════════════════════════════════════════════════
def _styles():
    return {
        'header_font':  Font(name='Calibri', bold=True, size=11, color='FFFFFF'),
        'header_fill':  PatternFill('solid', fgColor='1a472a'),
        'title_font':   Font(name='Calibri', bold=True, size=14, color='1a472a'),
        'subhead_font': Font(name='Calibri', bold=True, size=10, color='FFFFFF'),
        'subhead_fill': PatternFill('solid', fgColor='2d6a4f'),
        'center':       Alignment(horizontal='center', vertical='center'),
        'left':         Alignment(horizontal='left', vertical='center'),
        'border': Border(
            left=Side(style='thin', color='cccccc'),
            right=Side(style='thin', color='cccccc'),
            top=Side(style='thin', color='cccccc'),
            bottom=Side(style='thin', color='cccccc'),
        ),
        'alt_fill': PatternFill('solid', fgColor='f0fdf4'),
        'white_fill': PatternFill('solid', fgColor='FFFFFF'),
    }


def _write_stream_sheet(ws, league_name, stream, data, st):
    """Write one stream's standings + results to a worksheet."""
    standings = calc_standings(stream['_league_id'], stream['id'], data)
    results = sorted([r for r in data['results']
                      if r['leagueId'] == stream['_league_id']
                      and r['streamId'] == stream['id']],
                     key=lambda r: r['date'])

    row = 1
    ws.merge_cells(f'A{row}:K{row}')
    ws[f'A{row}'] = f"{league_name} — {stream['name']}"
    ws[f'A{row}'].font = st['title_font']
    ws[f'A{row}'].alignment = st['center']
    ws.row_dimensions[row].height = 30
    row += 1

    ws.merge_cells(f'A{row}:K{row}')
    ws[f'A{row}'] = f"Generated: {date.today().strftime('%d %B %Y')}"
    ws[f'A{row}'].font = Font(name='Calibri', size=9, color='888888')
    ws[f'A{row}'].alignment = st['center']
    row += 2

    headers = ['Pos', 'Team', 'P', 'W', 'D', 'L', 'GF', 'GA', 'GD', 'Pts']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = st['header_font']; cell.fill = st['header_fill']
        cell.alignment = st['center']; cell.border = st['border']
    ws.row_dimensions[row].height = 20
    row += 1

    for i, t in enumerate(standings):
        fill = st['alt_fill'] if i % 2 == 0 else st['white_fill']
        vals = [i+1, t['team'], t['p'], t['w'], t['d'], t['l'],
                t['gf'], t['ga'], t['gd'], t['pts']]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = st['center'] if col != 2 else st['left']
            cell.border = st['border']; cell.fill = fill
            cell.font = Font(name='Calibri', bold=(col == 10), size=10,
                             color='1a472a' if col == 10 else '000000')
        row += 1

    row += 2
    ws.merge_cells(f'A{row}:D{row}')
    ws[f'A{row}'] = 'Match Results'
    ws[f'A{row}'].font = st['title_font']
    row += 1

    for col, h in enumerate(['Date', 'Home Team', 'Score', 'Away Team'], 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = st['subhead_font']; cell.fill = st['subhead_fill']
        cell.alignment = st['center']; cell.border = st['border']
    row += 1

    for i, r in enumerate(results):
        fill = st['alt_fill'] if i % 2 == 0 else st['white_fill']
        for col, val in enumerate(
            [r['date'], r['home'], f"{r['gh']} - {r['ga']}", r['away']], 1
        ):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = st['center']; cell.border = st['border']
            cell.fill = fill; cell.font = Font(name='Calibri', size=10)
        row += 1

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[c].width = 7


def _safe_sheet_name(name):
    """Excel sheet names must be ≤31 chars and free of /\\?*[]:"""
    clean = re.sub(r'[\\/?*\[\]:]', ' ', str(name))[:31].strip()
    return clean or 'Sheet'


def _send_workbook(wb, filename):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output, download_name=filename, as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ────────────── EXPORT: ALL LEAGUES ──────────────────────────────────────────
@app.route('/api/export/all')
def export_all():
    data = load_data()
    wb = Workbook(); wb.remove(wb.active)
    st = _styles()
    used_names = set()
    for league in data['leagues']:
        for stream in league['streams']:
            stream['_league_id'] = league['id']
            sheet_name = _safe_sheet_name(f"{league['name'][:12]}-{stream['name'][:16]}")
            base = sheet_name; n = 1
            while sheet_name in used_names:
                sheet_name = _safe_sheet_name(f"{base[:28]}_{n}"); n += 1
            used_names.add(sheet_name)
            ws = wb.create_sheet(title=sheet_name)
            _write_stream_sheet(ws, league['name'], stream, data, st)
    if not wb.worksheets:
        wb.create_sheet('Empty')
    return _send_workbook(wb, f"LeagueBoss_All_{date.today().strftime('%Y-%m-%d')}.xlsx")


# ────────────── EXPORT: SINGLE LEAGUE ────────────────────────────────────────
@app.route('/api/export/league/<league_id>')
def export_league(league_id):
    data = load_data()
    league = find_league(data, league_id)
    if not league:
        return jsonify({'error': 'Not found'}), 404
    wb = Workbook(); wb.remove(wb.active)
    st = _styles()
    used = set()
    for stream in league['streams']:
        stream['_league_id'] = league_id
        name = _safe_sheet_name(stream['name'])
        base = name; n = 1
        while name in used:
            name = _safe_sheet_name(f"{base[:28]}_{n}"); n += 1
        used.add(name)
        ws = wb.create_sheet(title=name)
        _write_stream_sheet(ws, league['name'], stream, data, st)
    if not wb.worksheets:
        wb.create_sheet('Empty')
    safe = re.sub(r'[^A-Za-z0-9]+', '_', league['name'])[:40]
    return _send_workbook(wb, f"{safe}_{date.today().strftime('%Y-%m-%d')}.xlsx")


# ────────────── EXPORT: SINGLE STREAM ────────────────────────────────────────
@app.route('/api/export/stream/<league_id>/<stream_id>')
def export_stream(league_id, stream_id):
    data = load_data()
    league = find_league(data, league_id)
    stream = find_stream(league, stream_id)
    if not stream:
        return jsonify({'error': 'Not found'}), 404
    wb = Workbook(); wb.remove(wb.active)
    st = _styles()
    stream['_league_id'] = league_id
    ws = wb.create_sheet(title=_safe_sheet_name(stream['name']))
    _write_stream_sheet(ws, league['name'], stream, data, st)
    safe = re.sub(r'[^A-Za-z0-9]+', '_', f"{league['name']}_{stream['name']}")[:50]
    return _send_workbook(wb, f"{safe}_{date.today().strftime('%Y-%m-%d')}.xlsx")


# ────────────── EXPORT: SINGLE TEAM ──────────────────────────────────────────
@app.route('/api/export/team/<league_id>/<team_name>')
def export_team(league_id, team_name):
    data = load_data()
    league = find_league(data, league_id)
    if not league:
        return jsonify({'error': 'Not found'}), 404
    summary, matches = calc_team_summary(league_id, team_name, data)

    wb = Workbook(); wb.remove(wb.active)
    st = _styles()
    ws = wb.create_sheet(title=_safe_sheet_name(team_name))

    row = 1
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = f"{team_name} — {league['name']}"
    ws[f'A{row}'].font = st['title_font']; ws[f'A{row}'].alignment = st['center']
    ws.row_dimensions[row].height = 30
    row += 2

    # Summary block
    ws.merge_cells(f'A{row}:F{row}'); ws[f'A{row}'] = 'Season Summary'
    ws[f'A{row}'].font = st['title_font']; row += 1
    headers = ['P', 'W', 'D', 'L', 'GF', 'GA', 'GD', 'Pts']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = st['header_font']; cell.fill = st['header_fill']
        cell.alignment = st['center']; cell.border = st['border']
    row += 1
    for col, key in enumerate(['p', 'w', 'd', 'l', 'gf', 'ga', 'gd', 'pts'], 1):
        cell = ws.cell(row=row, column=col, value=summary[key])
        cell.alignment = st['center']; cell.border = st['border']
        cell.font = Font(name='Calibri', size=11, bold=(key == 'pts'))
    row += 2

    # Matches block
    ws.merge_cells(f'A{row}:F{row}'); ws[f'A{row}'] = 'All Matches'
    ws[f'A{row}'].font = st['title_font']; row += 1
    for col, h in enumerate(['Date', 'Venue', 'Opponent', 'Score', 'Result'], 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = st['subhead_font']; cell.fill = st['subhead_fill']
        cell.alignment = st['center']; cell.border = st['border']
    row += 1
    for i, m in enumerate(matches):
        fill = st['alt_fill'] if i % 2 == 0 else st['white_fill']
        for col, val in enumerate(
            [m['date'], m['venue'], m['opponent'], m['score'], m['outcome']], 1
        ):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = st['center'] if col != 3 else st['left']
            cell.border = st['border']; cell.fill = fill
            cell.font = Font(name='Calibri', size=10)
        row += 1

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 28
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10

    safe = re.sub(r'[^A-Za-z0-9]+', '_', f"{team_name}")[:40]
    return _send_workbook(wb, f"{safe}_{date.today().strftime('%Y-%m-%d')}.xlsx")


# ════════════════════════════════════════════════════════════════════════════
#  IMPORT (kept from v1, with league context)
# ════════════════════════════════════════════════════════════════════════════
def parse_result(s):
    if not s: return None
    s = str(s).strip().replace('\xa0', '').replace('\t', '')
    if s.lower() in ('no games this week', ''): return None
    m = re.match(r'^(.+?)\s+(\d+)\s*[-\u2013]\s*(\d+)\s+(.+)$', s)
    if m:
        return {'home': m.group(1).strip(), 'gh': int(m.group(2)),
                'ga': int(m.group(3)), 'away': m.group(4).strip()}
    return None


def parse_date(val):
    if not val: return None
    s = str(val).strip().replace('\t', '')
    for fmt in ['%A, %d %B %Y', '%d/%m/%Y', '%Y-%m-%d']:
        try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except: pass
    return None


@app.route('/api/import/<league_id>', methods=['POST'])
def import_excel(league_id):
    if not can_edit_league(league_id):
        return jsonify({'error': 'Unauthorised'}), 403
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file'}), 400

    try:
        wb = load_workbook(f, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header = next((r for r in rows if any(r)), None)
        stream_names = [str(c).strip() for c in header[1:] if c]

        data = load_data()
        league = find_league(data, league_id)
        if not league:
            return jsonify({'error': 'League not found'}), 404

        mapping = {}
        for name in stream_names:
            existing = next((s for s in league['streams'] if s['name'] == name), None)
            if existing:
                mapping[name] = existing['id']
            else:
                sid = f"s_{uuid.uuid4().hex[:8]}_{re.sub(r'[^a-z0-9]+', '_', name.lower())[:24]}"
                league['streams'].append({'id': sid, 'name': name, 'teams': []})
                mapping[name] = sid

        current_date, imported, skipped = None, 0, 0
        existing_keys = {
            f"{r['leagueId']}_{r['streamId']}_{r['date']}_{r['home']}_{r['away']}"
            for r in data['results']
        }
        for row in rows[1:]:
            if not any(row): continue
            d = parse_date(row[0])
            if d:
                current_date = d; continue
            if not current_date: continue
            for col_idx, stream_name in enumerate(stream_names):
                cell = row[col_idx + 1] if col_idx + 1 < len(row) else None
                parsed = parse_result(cell)
                if not parsed: continue
                sid = mapping[stream_name]
                key = f"{league_id}_{sid}_{current_date}_{parsed['home']}_{parsed['away']}"
                if key in existing_keys:
                    skipped += 1; continue
                stream = next(s for s in league['streams'] if s['id'] == sid)
                for tn in (parsed['home'], parsed['away']):
                    if tn not in stream['teams']:
                        stream['teams'].append(tn)
                data['results'].append({
                    'id': str(uuid.uuid4())[:8],
                    'leagueId': league_id, 'streamId': sid,
                    'date': current_date,
                    'home': parsed['home'], 'away': parsed['away'],
                    'gh': parsed['gh'], 'ga': parsed['ga'],
                })
                existing_keys.add(key); imported += 1

        save_data(data)
        return jsonify({'ok': True, 'imported': imported, 'skipped': skipped})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    print("\n⚽ LeagueBoss v2 is running!")
    print("   Open your browser: http://localhost:5000")
    print(f"   Default login → username: admin   password: {DEFAULT_SUPER_PASSWORD}\n")
    app.run(debug=True, port=5000)
