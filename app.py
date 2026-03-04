"""
LeagueBoss - Soccer League Manager
A Flask web app to manage league streams, log results and export standings to Excel.

Run with: python3 app.py
Then open: http://localhost:5000
"""

from flask import Flask, render_template, request, jsonify, send_file, session
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import os
import io
import re
from datetime import date, datetime

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'leagueboss-dev-key')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'safa2026')

def is_admin():
    return session.get('admin') == True


# ── DATA FILE ─────────────────────────────────────────────────────────────────
DATA_FILE = 'league_data.json'

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f:
            return json.load(f)
    return {'streams': [], 'results': []}

def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2)

# ── STANDINGS CALCULATION ─────────────────────────────────────────────────────
def calc_standings(stream_id, data):
    stream = next((s for s in data['streams'] if s['id'] == stream_id), None)
    if not stream:
        return []

    results = [r for r in data['results'] if r['streamId'] == stream_id]

    table = {}
    for team in stream['teams']:
        table[team] = {'team': team, 'p': 0, 'w': 0, 'd': 0, 'l': 0,
                       'gf': 0, 'ga': 0, 'gd': 0, 'pts': 0, 'form': []}

    sorted_results = sorted(results, key=lambda r: r['date'])

    for r in sorted_results:
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

    return sorted(table.values(),
                  key=lambda x: (-x['pts'], -x['gd'], -x['gf'], x['team']))

# ── ROUTES ────────────────────────────────────────────────────────────────────
# ── AUTH ──────────────────────────────────────────────────────────────────────
@app.route('/api/login', methods=['POST'])
def login():
    body = request.json
    if body.get('password') == ADMIN_PASSWORD:
        session['admin'] = True
        return jsonify({'ok': True})
    return jsonify({'error': 'Incorrect password'}), 401

@app.route('/api/logout', methods=['POST'])
def logout():
    session.pop('admin', None)
    return jsonify({'ok': True})

@app.route('/api/me')
def me():
    return jsonify({'admin': is_admin()})


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    return jsonify(load_data())

@app.route('/api/stream', methods=['POST'])
def add_stream():
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    body = request.json
    name = body.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400
    stream_id = f"s{len(data['streams'])}_{name.replace(' ','_').lower()}"
    data['streams'].append({'id': stream_id, 'name': name, 'teams': []})
    save_data(data)
    return jsonify({'ok': True, 'id': stream_id})

@app.route('/api/stream/<stream_id>', methods=['DELETE'])
def delete_stream(stream_id):
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    data['streams'] = [s for s in data['streams'] if s['id'] != stream_id]
    data['results'] = [r for r in data['results'] if r['streamId'] != stream_id]
    save_data(data)
    return jsonify({'ok': True})

@app.route('/api/team', methods=['POST'])
def add_team():
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    body = request.json
    stream_id = body.get('streamId')
    team_name = body.get('name', '').strip()
    stream = next((s for s in data['streams'] if s['id'] == stream_id), None)
    if not stream:
        return jsonify({'error': 'Stream not found'}), 404
    if team_name in stream['teams']:
        return jsonify({'error': 'Team already exists'}), 400
    stream['teams'].append(team_name)
    save_data(data)
    return jsonify({'ok': True})

@app.route('/api/team', methods=['DELETE'])
def delete_team():
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    body = request.json
    stream_id = body.get('streamId')
    team_name = body.get('name')
    stream = next((s for s in data['streams'] if s['id'] == stream_id), None)
    if stream:
        stream['teams'] = [t for t in stream['teams'] if t != team_name]
        data['results'] = [r for r in data['results']
                           if not (r['streamId'] == stream_id and
                                   (r['home'] == team_name or r['away'] == team_name))]
    save_data(data)
    return jsonify({'ok': True})

@app.route('/api/result', methods=['POST'])
def add_result():
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    body = request.json
    import uuid
    result = {
        'id': str(uuid.uuid4())[:8],
        'streamId': body['streamId'],
        'date': body['date'],
        'home': body['home'],
        'away': body['away'],
        'gh': int(body['gh']),
        'ga': int(body['ga']),
    }
    data['results'].append(result)
    save_data(data)
    return jsonify({'ok': True})

@app.route('/api/result/<result_id>', methods=['DELETE'])
def delete_result(result_id):
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    data = load_data()
    data['results'] = [r for r in data['results'] if r['id'] != result_id]
    save_data(data)
    return jsonify({'ok': True})

@app.route('/api/standings/<stream_id>')
def get_standings(stream_id):
    data = load_data()
    return jsonify(calc_standings(stream_id, data))


# ── IMPORT EXCEL ──────────────────────────────────────────────────────────────
def parse_result(result_str):
    """Parse 'Home Team X - X Away Team' into components"""
    if not result_str:
        return None
    s = str(result_str).strip().replace('\xa0', '').replace('\t', '')
    if s.lower() in ['no games this week', '']:
        return None
    pattern = r'^(.+?)\s+(\d+)\s*[-\u2013]\s*(\d+)\s+(.+)$'
    m = re.match(pattern, s)
    if m:
        return {
            'home': m.group(1).strip(),
            'gh': int(m.group(2)),
            'ga': int(m.group(3)),
            'away': m.group(4).strip(),
        }
    return None

def parse_date(val):
    """Parse date value from Excel cell"""
    if not val:
        return None
    s = str(val).strip().replace('\t', '')
    # Try common formats
    for fmt in ['%A, %d %B %Y', '%d/%m/%Y', '%Y-%m-%d']:
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except:
            pass
    return None

@app.route('/api/import', methods=['POST'])
def import_excel():
    if not is_admin():
        return jsonify({'error': 'Unauthorised'}), 403
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx)'}), 400

    try:
        from openpyxl import load_workbook
        wb = load_workbook(file, data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

        # Get stream names from header row (row 1, columns B onwards)
        header_row = next((r for r in rows if any(r)), None)
        stream_names = [str(c).strip() for c in header_row[1:] if c]

        data = load_data()
        existing_stream_names = [s['name'] for s in data['streams']]

        # Create streams that don't exist yet
        new_streams = {}
        for name in stream_names:
            existing = next((s for s in data['streams'] if s['name'] == name), None)
            if existing:
                new_streams[name] = existing['id']
            else:
                sid = f"s{len(data['streams'])}_{name.replace(' ','_').lower()}"
                data['streams'].append({'id': sid, 'name': name, 'teams': []})
                new_streams[name] = sid

        # Parse results
        current_date = None
        imported = 0
        skipped = 0
        existing_results = set(
            f"{r['streamId']}_{r['date']}_{r['home']}_{r['away']}"
            for r in data['results']
        )

        import uuid

        for row in rows[1:]:  # Skip header
            if not any(row):
                continue

            # Check if this is a date row
            first_cell = row[0]
            parsed_date = parse_date(first_cell)
            if parsed_date:
                current_date = parsed_date
                continue

            if not current_date:
                continue

            # Process each stream column
            for col_idx, stream_name in enumerate(stream_names):
                cell_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                result = parse_result(cell_val)
                if not result:
                    continue

                stream_id = new_streams.get(stream_name)
                if not stream_id:
                    continue

                # Check for duplicate
                key = f"{stream_id}_{current_date}_{result['home']}_{result['away']}"
                if key in existing_results:
                    skipped += 1
                    continue

                # Add teams if not already in stream
                stream = next(s for s in data['streams'] if s['id'] == stream_id)
                if result['home'] not in stream['teams']:
                    stream['teams'].append(result['home'])
                if result['away'] not in stream['teams']:
                    stream['teams'].append(result['away'])

                data['results'].append({
                    'id': str(uuid.uuid4())[:8],
                    'streamId': stream_id,
                    'date': current_date,
                    'home': result['home'],
                    'away': result['away'],
                    'gh': result['gh'],
                    'ga': result['ga'],
                })
                existing_results.add(key)
                imported += 1

        save_data(data)
        return jsonify({
            'ok': True,
            'imported': imported,
            'skipped': skipped,
            'streams': len(stream_names),
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── EXCEL EXPORT ──────────────────────────────────────────────────────────────
@app.route('/api/export')
def export_excel():
    data = load_data()
    wb = Workbook()
    wb.remove(wb.active)

    # Styles
    header_font   = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
    header_fill   = PatternFill('solid', fgColor='1a472a')  # Dark green
    title_font    = Font(name='Calibri', bold=True, size=14, color='1a472a')
    subhead_font  = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
    subhead_fill  = PatternFill('solid', fgColor='2d6a4f')
    accent_fill   = PatternFill('solid', fgColor='d8f3dc')
    center        = Alignment(horizontal='center', vertical='center')
    thin_border   = Border(
        left=Side(style='thin', color='cccccc'),
        right=Side(style='thin', color='cccccc'),
        top=Side(style='thin', color='cccccc'),
        bottom=Side(style='thin', color='cccccc'),
    )

    for stream in data['streams']:
        ws = wb.create_sheet(title=stream['name'][:31])
        standings = calc_standings(stream['id'], data)
        results = [r for r in data['results'] if r['streamId'] == stream['id']]
        results_sorted = sorted(results, key=lambda r: r['date'])

        row = 1

        # Title
        ws.merge_cells(f'A{row}:K{row}')
        ws[f'A{row}'] = f"{stream['name']} — League Standings"
        ws[f'A{row}'].font = title_font
        ws[f'A{row}'].alignment = center
        ws.row_dimensions[row].height = 30
        row += 1

        ws.merge_cells(f'A{row}:K{row}')
        ws[f'A{row}'] = f"Generated: {date.today().strftime('%d %B %Y')}"
        ws[f'A{row}'].font = Font(name='Calibri', size=9, color='888888')
        ws[f'A{row}'].alignment = center
        row += 2

        # Standings header
        headers = ['Pos', 'Team', 'P', 'W', 'D', 'L', 'GF', 'GA', 'GD', 'Pts']
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = thin_border
        ws.row_dimensions[row].height = 20
        row += 1

        # Standings rows
        for i, team in enumerate(standings):
            fill = PatternFill('solid', fgColor='f0fdf4') if i % 2 == 0 else PatternFill('solid', fgColor='FFFFFF')
            vals = [i+1, team['team'], team['p'], team['w'], team['d'],
                    team['l'], team['gf'], team['ga'], team['gd'], team['pts']]
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.alignment = center if col != 2 else Alignment(horizontal='left', vertical='center')
                cell.border = thin_border
                cell.fill = fill
                if col == 10:  # Points column
                    cell.font = Font(name='Calibri', bold=True, size=11, color='1a472a')
                else:
                    cell.font = Font(name='Calibri', size=10)
            row += 1

        row += 2

        # Results section
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = 'Match Results'
        ws[f'A{row}'].font = title_font
        row += 1

        result_headers = ['Date', 'Home Team', 'Score', 'Away Team']
        for col, h in enumerate(result_headers, 1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = subhead_font
            cell.fill = subhead_fill
            cell.alignment = center
            cell.border = thin_border
        row += 1

        for i, r in enumerate(results_sorted):
            fill = PatternFill('solid', fgColor='f0fdf4') if i % 2 == 0 else PatternFill('solid', fgColor='FFFFFF')
            for col, val in enumerate([r['date'], r['home'], f"{r['gh']} - {r['ga']}", r['away']], 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.alignment = center
                cell.border = thin_border
                cell.fill = fill
                cell.font = Font(name='Calibri', size=10)
            row += 1

        # Column widths
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 28
        for col in ['C','D','E','F','G','H','I','J']:
            ws.column_dimensions[col].width = 7

    # Save to memory and send
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"League_Standings_{date.today().strftime('%Y-%m-%d')}.xlsx"
    return send_file(output, download_name=filename,
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ── RUN ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("\n⚽ LeagueBoss is running!")
    print("   Open your browser and go to: http://localhost:5000\n")
    app.run(debug=True, port=5000)
