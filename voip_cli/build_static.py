#!/usr/bin/env python3
"""Gera index.html estático a partir da planilha Excel para GitHub Pages."""
import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from voip_cli.database import VoipDatabase

EXCEL_PATH = Path(__file__).parent.parent / "Telefones VOIPS atualizados 15 04 2026.xlsx"
OUTPUT = Path(__file__).parent.parent / "index.html"

db = VoipDatabase(str(EXCEL_PATH))
data = db.list_all()
s = db.stats()

modelos = sorted(set(r.get("Nome do Produto", "") for r in data if r.get("Nome do Produto")))

HTML = r"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Inventário VoIP</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
    --bg: #0a0e14; --surface: #121820; --surface2: #1a2330;
    --border: #233040; --text: #d1d5db; --text2: #8b949e;
    --heading: #f0f4f8; --accent: #3b82f6; --green: #22c55e; --red: #ef4444;
    --radius: 10px; --shadow: 0 4px 24px rgba(0,0,0,.4);
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Inter', -apple-system, sans-serif;
    background: var(--bg); padding: 24px; color: var(--text); min-height: 100vh;
}
.container { max-width: 1440px; margin: 0 auto; }
.header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 24px; flex-wrap: wrap; gap: 12px;
}
.header h1 { font-size: 22px; font-weight: 700; color: var(--heading); letter-spacing: -0.3px; }
.header h1 span { color: var(--accent); }
.header .badge {
    background: var(--surface2); border: 1px solid var(--border);
    padding: 6px 14px; border-radius: 20px; font-size: 13px; color: var(--text2);
}
.grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 12px; margin-bottom: 24px;
}
.stat-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 20px 16px;
    text-align: center; transition: border-color .2s, transform .2s;
}
.stat-card:hover { border-color: var(--accent); transform: translateY(-2px); }
.stat-card .num { font-size: 30px; font-weight: 700; color: var(--heading); line-height: 1.2; }
.stat-card .num.green { color: var(--green); }
.stat-card .num.red { color: var(--red); }
.stat-card .num.blue { color: var(--accent); }
.stat-card .label {
    font-size: 12px; color: var(--text2); margin-top: 6px;
    text-transform: uppercase; letter-spacing: .5px; font-weight: 500;
}
.toolbar {
    display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; align-items: center;
}
.toolbar input, .toolbar select {
    padding: 9px 14px; border: 1px solid var(--border); border-radius: 8px;
    font-size: 13px; background: var(--surface); color: var(--text); font-family: inherit;
}
.toolbar input::placeholder { color: #4a5568; }
.toolbar input:focus, .toolbar select:focus { outline: none; border-color: var(--accent); }
.toolbar input { flex: 1; min-width: 200px; }
.table-wrap {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); overflow: hidden;
}
.table-scroll { overflow-x: auto; max-height: 620px; overflow-y: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead { position: sticky; top: 0; z-index: 2; }
th {
    background: var(--surface2); padding: 12px 10px; text-align: left; font-weight: 600;
    font-size: 12px; color: var(--text2); text-transform: uppercase; letter-spacing: .4px;
    cursor: pointer; user-select: none; border-bottom: 1px solid var(--border); white-space: nowrap;
}
th:hover { color: var(--text); }
th::after { content: ' \u2195'; opacity: .3; font-size: 10px; }
td { padding: 11px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }
tbody tr { transition: background .15s; }
tbody tr:hover { background: rgba(59,130,246,.06); }
tbody tr:last-child td { border-bottom: none; }
.status-online { color: var(--green); font-weight: 600; }
.status-offline { color: var(--red); font-weight: 600; }
td a { color: var(--accent); text-decoration: none; }
td a:hover { text-decoration: underline; }
td.mono { font-family: 'SF Mono', 'Fira Code', monospace; font-size: 12px; letter-spacing: -.2px; }
.footer { text-align: center; margin-top: 24px; color: var(--text2); font-size: 12px; }
@media (max-width: 768px) {
    body { padding: 12px; }
    .header h1 { font-size: 18px; }
    .grid { grid-template-columns: repeat(3, 1fr); }
    .stat-card { padding: 14px 10px; }
    .stat-card .num { font-size: 22px; }
    table { font-size: 12px; }
    th, td { padding: 8px 6px; }
}
</style>
</head>
<body>
<div class="container">
<div class="header">
    <h1><span>VoIP</span> Inventário</h1>
    <span class="badge">__REGISTROS__ registros</span>
</div>
<div class="grid">
    <div class="stat-card"><div class="num blue">__TOTAL__</div><div class="label">Total</div></div>
    <div class="stat-card"><div class="num green">__ONLINE__</div><div class="label">Online</div></div>
    <div class="stat-card"><div class="num red">__OFFLINE__</div><div class="label">Offline</div></div>
    <div class="stat-card"><div class="num blue">__SETORES__</div><div class="label">Setores</div></div>
    <div class="stat-card"><div class="num blue">__MODELOS__</div><div class="label">Modelos</div></div>
</div>
<div class="table-wrap">
<div class="toolbar" style="padding:12px 16px 0 16px;">
    <input type="text" id="search" placeholder="Buscar IP, setor, telefone, MAC..." oninput="filter()">
    <select id="statusFilter" onchange="filter()">
        <option value="">Status</option>
        <option value="ONLINE">Online</option>
        <option value="OFFLINE">Offline</option>
    </select>
    <select id="modeloFilter" onchange="filter()">
        <option value="">Modelo</option>
        __MODELOS_OPTIONS__
    </select>
</div>
<div class="table-scroll">
<table>
<thead><tr>
    <th onclick="sort(0)">IP</th>
    <th onclick="sort(1)">Setor</th>
    <th onclick="sort(2)">Telefone</th>
    <th onclick="sort(3)">Status</th>
    <th onclick="sort(4)">Modelo</th>
    <th>MAC</th>
    <th>Serial</th>
</tr></thead>
<tbody id="tbody">
__ROWS__
</tbody>
</table>
</div>
</div>
<div class="footer">Atualizado em __DATE__</div>
</div>
<script>
const data = __DATA__;
function filter() {
    const q = document.getElementById('search').value.toLowerCase();
    const st = document.getElementById('statusFilter').value;
    const mo = document.getElementById('modeloFilter').value;
    document.querySelectorAll('#tbody tr').forEach(r => {
        const txt = r.textContent.toLowerCase();
        const status = r.children[3]?.textContent.trim() || '';
        let show = true;
        if (q && !txt.includes(q)) show = false;
        if (st && status !== st) show = false;
        if (mo && !txt.includes(mo.toLowerCase())) show = false;
        r.style.display = show ? '' : 'none';
    });
}
function sort(n) {
    const tbody = document.getElementById('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const dir = tbody.dataset.sort === 'asc' ? 'desc' : 'asc';
    tbody.dataset.sort = dir;
    rows.sort((a, b) => {
        const va = a.children[n].textContent.trim();
        const vb = b.children[n].textContent.trim();
        const na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return dir === 'asc' ? na - nb : nb - na;
        return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    rows.forEach(r => tbody.appendChild(r));
}
</script>
</body>
</html>"""

modelos_options = "\n".join(f'        <option value="{m}">{m}</option>' for m in modelos)

rows = []
for r in data:
    status_class = r.get("Status", "").lower()
    row = f"""    <tr>
        <td><a href="http://{r['IP']}" target="_blank">{r['IP']}</a></td>
        <td>{r.get('Setor', '')}</td>
        <td>{r.get('Telefone', '')}</td>
        <td class="status-{status_class}">{r.get('Status', '')}</td>
        <td>{r.get('Nome do Produto', '')}</td>
        <td class="mono">{r.get('MAC', '')}</td>
        <td class="mono">{r.get('Serial', '')}</td>
    </tr>"""
    rows.append(row)

from datetime import datetime
html = HTML.replace("__REGISTROS__", str(len(data)))
html = html.replace("__TOTAL__", str(s["total"]))
html = html.replace("__ONLINE__", str(s["online"]))
html = html.replace("__OFFLINE__", str(s["offline"]))
html = html.replace("__SETORES__", str(s["setores_unicos"]))
html = html.replace("__MODELOS__", str(len(s["modelos"])))
html = html.replace("__MODELOS_OPTIONS__", modelos_options)
html = html.replace("__ROWS__", "\n".join(rows))
html = html.replace("__DATA__", json.dumps(data))
html = html.replace("__DATE__", datetime.now().strftime("%d/%m/%Y %H:%M"))

OUTPUT.write_text(html, encoding="utf-8")
print(f"index.html gerado com {len(data)} registros em {OUTPUT}")
