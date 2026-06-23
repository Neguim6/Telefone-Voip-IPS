#!/usr/bin/env python3
"""Gera data.json + index.html interativo para GitHub Pages."""
import sys
import json
from pathlib import Path
from datetime import datetime
sys.path.insert(0, str(Path(__file__).parent.parent))
from voip_cli.database import VoipDatabase, COLUMNS

EXCEL_PATH = Path(__file__).parent.parent / "Telefones VOIPS atualizados 15 04 2026.xlsx"
DATA_JSON = Path(__file__).parent.parent / "data.json"
INDEX_HTML = Path(__file__).parent.parent / "index.html"

db = VoipDatabase(str(EXCEL_PATH))
data = db.list_all()
s = db.stats()

DATA_JSON.write_text(json.dumps({"data": data, "stats": s, "gerado": datetime.now().isoformat()}, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"data.json gerado com {len(data)} registros")

modelos = sorted(set(r.get("Nome do Produto", "") for r in data if r.get("Nome do Produto")))

modelos_options = "\n".join(f'        <option value="{m}">{m}</option>' for m in modelos)

rows = []
for r in data:
    sc = r.get("Status", "").lower()
    ip = r['IP']
    setor = r.get('Setor', '') or ''
    tel = r.get('Telefone', '') or ''
    status = r.get('Status', '') or ''
    modelo = r.get('Nome do Produto', '') or ''
    mac = r.get('MAC', '') or ''
    serial = r.get('Serial', '') or ''
    rows.append(f"""    <tr>
        <td><a href="http://{ip}" target="_blank">{ip}</a></td>
        <td>{setor}</td>
        <td>{tel}</td>
        <td class="status-{sc}">{status}</td>
        <td>{modelo}</td>
        <td class="mono">{mac}</td>
        <td class="mono">{serial}</td>
        <td><div class="actions"><button class="edit-btn" onclick="editar('{ip}')">Editar</button><button class="del-btn" onclick="remover('{ip}')">Remover</button></div></td>
    </tr>""")

ROWS_HTML = "\n".join(rows)
MODELOS_OPTIONS = modelos_options
DATA_JSON_STR = json.dumps(data, ensure_ascii=False)
DATE_STR = datetime.now().strftime("%d/%m/%Y %H:%M")

# replace emoji that causes encoding issues on Windows
HTML_CONTENT = f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Inventário VoIP</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: 'Inter', sans-serif; background: #0a0e14; padding: 24px; color: #d1d5db; }}
.container {{ max-width: 1440px; margin: 0 auto; }}
.header {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px; flex-wrap: wrap; gap: 12px; }}
.header h1 {{ font-size: 22px; font-weight: 700; color: #f0f4f8; }}
.header h1 span {{ color: #3b82f6; }}
.header .badge {{ background: #1a2330; border: 1px solid #233040; padding: 6px 14px; border-radius: 20px; font-size: 13px; color: #8b949e; }}
.header-left {{ display: flex; align-items: center; gap: 12px; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }}
.stat-card {{ background: #121820; border: 1px solid #233040; border-radius: 10px; padding: 20px 16px; text-align: center; }}
.stat-card .num {{ font-size: 30px; font-weight: 700; color: #f0f4f8; }}
.stat-card .num.green {{ color: #22c55e; }}
.stat-card .num.red {{ color: #ef4444; }}
.stat-card .num.blue {{ color: #3b82f6; }}
.stat-card .label {{ font-size: 12px; color: #8b949e; margin-top: 6px; text-transform: uppercase; letter-spacing: .5px; font-weight: 500; }}
.toolbar {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; align-items: center; }}
.toolbar input, .toolbar select {{ padding: 9px 14px; border: 1px solid #233040; border-radius: 8px; font-size: 13px; background: #121820; color: #d1d5db; font-family: inherit; }}
.toolbar input::placeholder {{ color: #4a5568; }}
.toolbar input:focus, .toolbar select:focus {{ outline: none; border-color: #3b82f6; }}
.toolbar input[type="text"] {{ flex: 1; min-width: 200px; }}
.btn {{ padding: 9px 18px; border: none; border-radius: 8px; font-family: inherit; font-size: 13px; font-weight: 500; cursor: pointer; transition: all .2s; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }}
.btn-primary {{ background: #3b82f6; color: #fff; }}
.btn-primary:hover {{ background: #60a5fa; }}
.btn-green {{ background: #22c55e; color: #fff; }}
.btn-green:hover {{ background: #16a34a; }}
.btn-outline {{ background: transparent; color: #8b949e; border: 1px solid #233040; }}
.btn-outline:hover {{ background: #1a2330; }}
.btn-red {{ background: #ef4444; color: #fff; }}
.btn-red:hover {{ background: #dc2626; }}
.table-wrap {{ background: #121820; border: 1px solid #233040; border-radius: 10px; overflow: hidden; }}
.table-scroll {{ overflow-x: auto; max-height: 620px; overflow-y: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
thead {{ position: sticky; top: 0; z-index: 2; }}
th {{ background: #1a2330; padding: 12px 10px; text-align: left; font-weight: 600; font-size: 12px; color: #8b949e; text-transform: uppercase; letter-spacing: .4px; cursor: pointer; border-bottom: 1px solid #233040; white-space: nowrap; }}
th:hover {{ color: #d1d5db; }}
td {{ padding: 11px 10px; border-bottom: 1px solid #233040; white-space: nowrap; }}
tbody tr:hover {{ background: rgba(59,130,246,.06); }}
.status-online {{ color: #22c55e; font-weight: 600; }}
.status-offline {{ color: #ef4444; font-weight: 600; }}
td a {{ color: #3b82f6; text-decoration: none; }}
td a:hover {{ text-decoration: underline; }}
td.mono {{ font-family: 'SF Mono', monospace; font-size: 12px; }}
.actions {{ display: flex; gap: 4px; }}
.actions button {{ padding: 4px 10px; border-radius: 6px; border: 1px solid #233040; background: transparent; color: #8b949e; font-size: 12px; cursor: pointer; font-family: inherit; }}
.actions .edit-btn:hover {{ border-color: #3b82f6; color: #3b82f6; }}
.actions .del-btn:hover {{ border-color: #ef4444; color: #ef4444; }}
.modal {{ display: none; position: fixed; z-index: 99; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,.65); backdrop-filter: blur(4px); }}
.modal-content {{ background: #121820; margin: 4% auto; padding: 28px; border-radius: 14px; width: 92%; max-width: 520px; border: 1px solid #233040; animation: slideUp .25s ease-out; }}
@keyframes slideUp {{ from {{ opacity: 0; transform: translateY(20px); }} to {{ opacity: 1; transform: translateY(0); }} }}
.modal h2 {{ font-size: 18px; font-weight: 700; color: #f0f4f8; margin-bottom: 4px; }}
.modal .sub {{ color: #8b949e; font-size: 13px; margin-bottom: 20px; }}
.modal label {{ display: block; margin-top: 14px; font-size: 12px; font-weight: 600; color: #8b949e; text-transform: uppercase; letter-spacing: .3px; }}
.modal input, .modal select {{ width: 100%; padding: 10px 12px; margin-top: 4px; border: 1px solid #233040; border-radius: 8px; background: #0a0e14; color: #d1d5db; font-family: inherit; font-size: 14px; }}
.modal input:focus, .modal select:focus {{ outline: none; border-color: #3b82f6; }}
.modal .row2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; }}
.modal .row3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 10px; }}
.modal-buttons {{ margin-top: 24px; display: flex; gap: 10px; justify-content: end; }}
.modal-buttons button {{ padding: 10px 22px; border-radius: 8px; border: none; font-family: inherit; font-size: 14px; font-weight: 500; cursor: pointer; }}
#toast {{ visibility: hidden; min-width: 260px; background: #1a2330; color: #f0f4f8; text-align: center; border-radius: 10px; padding: 14px 24px; position: fixed; bottom: 30px; left: 50%; transform: translateX(-50%); z-index: 100; border: 1px solid #233040; font-size: 14px; font-weight: 500; }}
#toast.show {{ visibility: visible; animation: toastIn .3s, toastOut .3s 2.7s; }}
@keyframes toastIn {{ from {{ opacity: 0; transform: translateX(-50%) translateY(10px); }} to {{ opacity: 1; transform: translateX(-50%) translateY(0); }} }}
@keyframes toastOut {{ from {{ opacity: 1; }} to {{ opacity: 0; }} }}
</style>
</head>
<body>
<div class="container">

<div class="header">
    <div class="header-left">
        <h1><span>VoIP</span> Inventário</h1>
        <span class="badge" id="totalBadge">0 registros</span>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-outline" onclick="configToken()">&#128273; Token</button>
        <button class="btn btn-primary" onclick="abrirAdicionar()">+ Novo</button>
        <button class="btn btn-green" id="btnSalvar" onclick="salvar()">&#128190; Salvar</button>
    </div>
</div>

<div class="grid" id="statsGrid"></div>

<div class="table-wrap">
<div class="toolbar" style="padding:12px 16px 0 16px;">
    <input type="text" id="search" placeholder="Buscar IP, setor, telefone, MAC..." oninput="filtrar()">
    <select id="statusFilter" onchange="filtrar()">
        <option value="">Status</option>
        <option value="ONLINE">Online</option>
        <option value="OFFLINE">Offline</option>
    </select>
    <select id="modeloFilter" onchange="filtrar()">
        <option value="">Modelo</option>
        {MODELOS_OPTIONS}
    </select>
</div>
<div class="table-scroll">
<table>
<thead><tr>
    <th onclick="ordenar(0)">IP</th>
    <th onclick="ordenar(1)">Setor</th>
    <th onclick="ordenar(2)">Telefone</th>
    <th onclick="ordenar(3)">Status</th>
    <th onclick="ordenar(4)">Modelo</th>
    <th>MAC</th>
    <th>Serial</th>
    <th>Ações</th>
</tr></thead>
<tbody id="tbody">
{ROWS_HTML}
</tbody>
</table>
</div>
</div>
<div class="footer" style="text-align:center;margin-top:20px;color:#8b949e;font-size:12px;">Atualizado em {DATE_STR}</div>
</div>

<div id="modalToken" class="modal">
<div class="modal-content">
    <h2>&#128273; Token do GitHub</h2>
    <p class="sub">Para salvar alterações, crie um token em <strong>github.com/settings/tokens</strong> (escopo: <strong>repo</strong>) e cole abaixo.</p>
    <label>Token</label><input type="password" id="inputToken" placeholder="ghp_...">
    <label>Repositório</label><input id="inputRepo" placeholder="usuario/repositorio" value="Neguim6/Telefone-Voip-IPS">
    <div class="modal-buttons">
        <button class="btn btn-outline" onclick="fecharModal('modalToken')">Cancelar</button>
        <button class="btn btn-primary" onclick="salvarToken()">Salvar</button>
    </div>
</div>
</div>

<div id="modalAdd" class="modal">
<div class="modal-content">
    <h2>Novo Telefone</h2>
    <p class="sub">Preencha os campos para adicionar</p>
    <label>IP *</label><input id="addIp" placeholder="10.75.16.xxx">
    <label>Setor *</label><input id="addSetor" placeholder="Ex: TI">
    <label>Telefone *</label><input id="addTel" placeholder="3027-xxxx">
    <div class="row3">
        <div><label>Modelo</label><input id="addModelo" placeholder="P11G"></div>
        <div><label>MAC</label><input id="addMac" placeholder="00:21:F2:..."></div>
        <div><label>Serial</label><input id="addSerial"></div>
    </div>
    <div class="row2">
        <div><label>Hardware</label><input id="addHardware"></div>
        <div><label>Firmware</label><input id="addFirmware"></div>
    </div>
    <label>Senha *</label><input type="password" id="addSenha" placeholder="Digite a senha">
    <div class="modal-buttons">
        <button class="btn btn-outline" onclick="fecharModal('modalAdd')">Cancelar</button>
        <button class="btn btn-primary" onclick="confirmarAdicao()">Salvar</button>
    </div>
</div>
</div>

<div id="modalEdit" class="modal">
<div class="modal-content">
    <h2>Editar Telefone</h2>
    <p class="sub" id="editIpDisplay"></p>
    <div class="row2">
        <div><label>Setor</label><input id="editSetor"></div>
        <div><label>Telefone</label><input id="editTel"></div>
    </div>
    <div class="row2">
        <div><label>Status</label><select id="editStatus"><option>ONLINE</option><option>OFFLINE</option></select></div>
        <div><label>Modelo</label><input id="editModelo"></div>
    </div>
    <div class="row2">
        <div><label>MAC</label><input id="editMac"></div>
        <div><label>Serial</label><input id="editSerial"></div>
    </div>
    <label>Firmware</label><input id="editFirmware">
    <label>Senha *</label><input type="password" id="editSenha" placeholder="Digite a senha">
    <div class="modal-buttons">
        <button class="btn btn-outline" onclick="fecharModal('modalEdit')">Cancelar</button>
        <button class="btn btn-primary" onclick="confirmarEdicao()">Salvar</button>
    </div>
</div>
</div>

<div id="modalRemove" class="modal">
<div class="modal-content">
    <h2>Remover Telefone</h2>
    <p class="sub">Tem certeza que deseja remover <strong id="removeIpDisplay" style="color:#f0f4f8"></strong>?</p>
    <label>Senha *</label><input type="password" id="removeSenha" placeholder="Digite a senha">
    <div class="modal-buttons">
        <button class="btn btn-outline" onclick="fecharModal('modalRemove')">Cancelar</button>
        <button class="btn btn-red" onclick="confirmarRemocao()">Remover</button>
    </div>
</div>
</div>

<div id="toast"></div>

<script>
const SENHA = "258456";
let dados = {{"data": [], "stats": {{}} }};
let dadosOriginais = [];

function toast(msg, ok) {{
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.style.borderColor = ok ? '#22c55e' : '#ef4444';
    t.className = 'show';
    setTimeout(() => t.className = '', 3000);
}}

function carregar() {{
    fetch('data.json?_=' + Date.now())
        .then(r => r.json())
        .then(j => {{
            dados = j;
            dadosOriginais = JSON.parse(JSON.stringify(j.data));
            renderizar();
        }})
        .catch(() => toast('Erro ao carregar dados', false));
}}

function atualizarStats() {{
    const data = dados.data;
    const total = data.length;
    const online = data.filter(r => r.Status === 'ONLINE').length;
    const offline = data.filter(r => r.Status === 'OFFLINE').length;
    const setores = new Set(data.map(r => r.Setor)).size;
    const modelos = new Set(data.map(r => r['Nome do Produto']).filter(Boolean)).size;
    document.getElementById('totalBadge').textContent = total + ' registros';
    const items = [
        {{num: total, label:'Total', cls:'blue'}},
        {{num: online, label:'Online', cls:'green'}},
        {{num: offline, label:'Offline', cls:'red'}},
        {{num: setores, label:'Setores', cls:'blue'}},
        {{num: modelos, label:'Modelos', cls:'blue'}},
    ];
    document.getElementById('statsGrid').innerHTML = items.map(s =>
        '<div class="stat-card"><div class="num ' + s.cls + '">' + s.num + '</div><div class="label">' + s.label + '</div></div>'
    ).join('');
}}

function renderizar() {{
    atualizarStats();
    const tbody = document.getElementById('tbody');
    tbody.innerHTML = dados.data.map(r => {{
        const sc = (r.Status || '').toLowerCase();
        return '<tr>' +
            '<td><a href=\"http://' + r.IP + '\" target=\"_blank\">' + r.IP + '</a></td>' +
            '<td>' + (r.Setor || '') + '</td>' +
            '<td>' + (r.Telefone || '') + '</td>' +
            '<td class=\"status-' + sc + '\">' + (r.Status || '') + '</td>' +
            '<td>' + (r['Nome do Produto'] || '') + '</td>' +
            '<td class=\"mono\">' + (r.MAC || '') + '</td>' +
            '<td class=\"mono\">' + (r.Serial || '') + '</td>' +
            '<td><div class=\"actions\">' +
            '<button class=\"edit-btn\" onclick=\"editar(\\'' + r.IP + '\\')\">Editar</button>' +
            '<button class=\"del-btn\" onclick=\"remover(\\'' + r.IP + '\\')\">Remover</button>' +
            '</div></td></tr>';
    }}).join('');
    filtrar();
}}

function salvar() {{
    const token = localStorage.getItem('gh_token');
    const repo = localStorage.getItem('gh_repo') || 'Neguim6/Telefone-Voip-IPS';
    if (!token) {{ toast('Configure o token do GitHub primeiro!', false); abrirToken(); return; }}
    if (JSON.stringify(dados.data) === JSON.stringify(dadosOriginais)) {{ toast('Nenhuma alteração para salvar.', true); return; }}

    const btn = document.getElementById('btnSalvar');
    btn.textContent = 'Salvando...';
    btn.disabled = true;

    const payload = {{
        data: dados.data,
        stats: {{
            total: dados.data.length,
            online: dados.data.filter(r => r.Status === 'ONLINE').length,
            offline: dados.data.filter(r => r.Status === 'OFFLINE').length,
            setores_unicos: new Set(dados.data.map(r => r.Setor)).size,
            modelos: {{}}
        }},
        gerado: new Date().toISOString()
    }};

    const url = 'https://api.github.com/repos/' + repo + '/contents/data.json';
    fetch(url + '?ref=main', {{
        headers: {{'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json'}}
    }})
    .then(r => r.json())
    .then(meta => {{
        const sha = meta.sha;
        const content = JSON.stringify(payload);
        return fetch(url, {{
            method: 'PUT',
            headers: {{'Authorization': 'token ' + token, 'Content-Type': 'application/json'}},
            body: JSON.stringify({{
                message: 'Atualiza data.json via site',
                content: btoa(unescape(encodeURIComponent(content))),
                sha: sha,
                branch: 'main'
            }})
        }});
    }})
    .then(r => r.json())
    .then(res => {{
        if (res.content) {{
            toast('Salvo no GitHub! As alterações estarão visíveis em instantes.', true);
            dadosOriginais = JSON.parse(JSON.stringify(dados.data));
        }} else {{
            toast('Erro: ' + (res.message || 'desconhecido'), false);
        }}
    }})
    .catch(err => toast('Erro: ' + err.message, false))
    .finally(() => {{ btn.textContent = '💾 Salvar'; btn.disabled = false; }});
}}

function checkSenha(s) {{ return s === SENHA; }}

function abrirAdicionar() {{
    ['addIp','addSetor','addTel','addModelo','addMac','addSerial','addHardware','addFirmware','addSenha'].forEach(function(id) {{ document.getElementById(id).value = ''; }});
    document.getElementById('modalAdd').style.display = 'block';
}}

function confirmarAdicao() {{
    if (!checkSenha(document.getElementById('addSenha').value)) {{ toast('Senha incorreta!', false); return; }}
    var ip = document.getElementById('addIp').value.trim();
    var setor = document.getElementById('addSetor').value.trim();
    var tel = document.getElementById('addTel').value.trim();
    if (!ip || !setor || !tel) {{ toast('IP, Setor e Telefone são obrigatórios!', false); return; }}
    if (dados.data.some(function(r) {{ return r.IP === ip; }})) {{ toast('IP já existe!', false); return; }}
    dados.data.push({{
        IP: ip, Setor: setor, Telefone: tel, Status: 'ONLINE',
        'Nome do Produto': document.getElementById('addModelo').value.trim(),
        MAC: document.getElementById('addMac').value.trim(),
        Serial: document.getElementById('addSerial').value.trim(),
        Hardware: document.getElementById('addHardware').value.trim(),
        Firmware: document.getElementById('addFirmware').value.trim()
    }});
    fecharModal('modalAdd');
    renderizar();
    toast('Telefone adicionado! Clique em 💾 Salvar para persistir.', true);
}}

function editar(ip) {{
    var r = dados.data.find(function(d) {{ return d.IP === ip; }});
    if (!r) return;
    document.getElementById('editIpDisplay').textContent = 'IP: ' + ip;
    document.getElementById('editSetor').value = r.Setor || '';
    document.getElementById('editTel').value = r.Telefone || '';
    document.getElementById('editStatus').value = r.Status || 'ONLINE';
    document.getElementById('editModelo').value = r['Nome do Produto'] || '';
    document.getElementById('editMac').value = r.MAC || '';
    document.getElementById('editSerial').value = r.Serial || '';
    document.getElementById('editFirmware').value = r.Firmware || '';
    document.getElementById('editSenha').value = '';
    document.getElementById('modalEdit').style.display = 'block';
}}

function confirmarEdicao() {{
    if (!checkSenha(document.getElementById('editSenha').value)) {{ toast('Senha incorreta!', false); return; }}
    var ip = document.getElementById('editIpDisplay').textContent.replace('IP: ','');
    var r = dados.data.find(function(d) {{ return d.IP === ip; }});
    if (!r) return;
    r.Setor = document.getElementById('editSetor').value;
    r.Telefone = document.getElementById('editTel').value;
    r.Status = document.getElementById('editStatus').value;
    r['Nome do Produto'] = document.getElementById('editModelo').value;
    r.MAC = document.getElementById('editMac').value;
    r.Serial = document.getElementById('editSerial').value;
    r.Firmware = document.getElementById('editFirmware').value;
    fecharModal('modalEdit');
    renderizar();
    toast('Telefone atualizado! Clique em 💾 Salvar para persistir.', true);
}}

function remover(ip) {{
    document.getElementById('removeIpDisplay').textContent = ip;
    document.getElementById('removeSenha').value = '';
    document.getElementById('modalRemove').style.display = 'block';
}}

function confirmarRemocao() {{
    if (!checkSenha(document.getElementById('removeSenha').value)) {{ toast('Senha incorreta!', false); return; }}
    var ip = document.getElementById('removeIpDisplay').textContent;
    dados.data = dados.data.filter(function(r) {{ return r.IP !== ip; }});
    fecharModal('modalRemove');
    renderizar();
    toast('Telefone removido! Clique em 💾 Salvar para persistir.', true);
}}

function configToken() {{ abrirToken(); }}
function abrirToken() {{
    document.getElementById('inputToken').value = localStorage.getItem('gh_token') || '';
    document.getElementById('inputRepo').value = localStorage.getItem('gh_repo') || 'Neguim6/Telefone-Voip-IPS';
    document.getElementById('modalToken').style.display = 'block';
}}

function salvarToken() {{
    var t = document.getElementById('inputToken').value.trim();
    var r = document.getElementById('inputRepo').value.trim();
    if (!t || !r) {{ toast('Preencha token e repositório', false); return; }}
    localStorage.setItem('gh_token', t);
    localStorage.setItem('gh_repo', r);
    fecharModal('modalToken');
    toast('Token salvo!', true);
}}

function fecharModal(id) {{ document.getElementById(id).style.display = 'none'; }}
window.onclick = function(e) {{
    ['modalToken','modalAdd','modalEdit','modalRemove'].forEach(function(id) {{ if (e.target == document.getElementById(id)) fecharModal(id); }});
}};

function filtrar() {{
    var q = document.getElementById('search').value.toLowerCase();
    var st = document.getElementById('statusFilter').value;
    var mo = document.getElementById('modeloFilter').value;
    document.querySelectorAll('#tbody tr').forEach(function(r) {{
        var txt = r.textContent.toLowerCase();
        var status = r.children[3]?.textContent.trim() || '';
        var show = true;
        if (q && !txt.includes(q)) show = false;
        if (st && status !== st) show = false;
        if (mo && !txt.includes(mo.toLowerCase())) show = false;
        r.style.display = show ? '' : 'none';
    }});
}}

function ordenar(n) {{
    var tbody = document.getElementById('tbody');
    var rows = Array.from(tbody.querySelectorAll('tr'));
    var dir = tbody.dataset.sort === 'asc' ? 'desc' : 'asc';
    tbody.dataset.sort = dir;
    rows.sort(function(a, b) {{
        var va = a.children[n].textContent.trim();
        var vb = b.children[n].textContent.trim();
        var na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return dir === 'asc' ? na - nb : nb - na;
        return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    }});
    rows.forEach(function(r) {{ tbody.appendChild(r); }});
}}

carregar();
</script>
</body>
</html>"""

INDEX_HTML.write_text(HTML_CONTENT, encoding="utf-8")
print(f"index.html gerado com {len(data)} registros")
