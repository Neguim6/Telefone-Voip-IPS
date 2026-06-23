import sys
import subprocess
import platform
import os
from pathlib import Path
from flask import Flask, render_template_string, request, redirect, url_for, flash, jsonify

sys.path.insert(0, str(Path(__file__).parent.parent))
from voip_cli.database import VoipDatabase, COLUMNS

EXCEL_PATH = Path(__file__).parent.parent / "Telefones VOIPS atualizados 15 04 2026.xlsx"
SENHA = "258456"

app = Flask(__name__)
app.secret_key = "voip-secret-key"


def get_db():
    return VoipDatabase(str(EXCEL_PATH))


HTML = r"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Inventário VoIP</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
    --bg: #0a0e14;
    --surface: #121820;
    --surface2: #1a2330;
    --border: #233040;
    --text: #d1d5db;
    --text2: #8b949e;
    --heading: #f0f4f8;
    --accent: #3b82f6;
    --accent-hover: #60a5fa;
    --green: #22c55e;
    --red: #ef4444;
    --radius: 10px;
    --shadow: 0 4px 24px rgba(0,0,0,.4);
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Inter', -apple-system, sans-serif;
    background: var(--bg); padding: 24px; color: var(--text);
    min-height: 100vh;
}
.container { max-width: 1440px; margin: 0 auto; }

/* header */
.header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 24px; flex-wrap: wrap; gap: 12px;
}
.header h1 {
    font-size: 22px; font-weight: 700; color: var(--heading);
    letter-spacing: -0.3px;
}
.header h1 span { color: var(--accent); }
.header .badge {
    background: var(--surface2); border: 1px solid var(--border);
    padding: 6px 14px; border-radius: 20px; font-size: 13px; color: var(--text2);
}

/* cards */
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
.stat-card .num {
    font-size: 30px; font-weight: 700; color: var(--heading);
    line-height: 1.2;
}
.stat-card .num.green { color: var(--green); }
.stat-card .num.red { color: var(--red); }
.stat-card .num.blue { color: var(--accent); }
.stat-card .label {
    font-size: 12px; color: var(--text2); margin-top: 6px;
    text-transform: uppercase; letter-spacing: .5px; font-weight: 500;
}

/* toolbar */
.toolbar {
    display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px;
    align-items: center;
}
.toolbar input, .toolbar select {
    padding: 9px 14px; border: 1px solid var(--border);
    border-radius: 8px; font-size: 13px; background: var(--surface);
    color: var(--text); font-family: inherit; transition: border-color .2s;
}
.toolbar input::placeholder { color: #4a5568; }
.toolbar input:focus, .toolbar select:focus { outline: none; border-color: var(--accent); }
.toolbar input { flex: 1; min-width: 200px; }
.btn {
    padding: 9px 18px; border: none; border-radius: 8px;
    font-family: inherit; font-size: 13px; font-weight: 500;
    cursor: pointer; transition: all .2s; display: inline-flex;
    align-items: center; gap: 6px; white-space: nowrap;
}
.btn-primary { background: var(--accent); color: #fff; }
.btn-primary:hover { background: var(--accent-hover); transform: translateY(-1px); }
.btn-green { background: var(--green); color: #fff; }
.btn-green:hover { background: #16a34a; transform: translateY(-1px); }
.btn-ghost {
    background: transparent; color: var(--text2);
    border: 1px solid var(--border);
}
.btn-ghost:hover { background: var(--surface2); color: var(--text); }

/* table */
.table-wrap {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); overflow: hidden;
}
.table-scroll { overflow-x: auto; max-height: 620px; overflow-y: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead { position: sticky; top: 0; z-index: 2; }
th {
    background: var(--surface2); padding: 12px 10px;
    text-align: left; font-weight: 600; font-size: 12px;
    color: var(--text2); text-transform: uppercase; letter-spacing: .4px;
    cursor: pointer; user-select: none; border-bottom: 1px solid var(--border);
    white-space: nowrap;
}
th:hover { color: var(--text); }
th::after { content: ' ↕'; opacity: .3; font-size: 10px; }
td {
    padding: 11px 10px; border-bottom: 1px solid var(--border);
    white-space: nowrap;
}
tbody tr { transition: background .15s; }
tbody tr:hover { background: rgba(59,130,246,.06); }
tbody tr:last-child td { border-bottom: none; }
.status-online { color: var(--green); font-weight: 600; }
.status-offline { color: var(--red); font-weight: 600; }
td a { color: var(--accent); text-decoration: none; }
td a:hover { text-decoration: underline; }
td.mono { font-family: 'SF Mono', 'Fira Code', monospace; font-size: 12px; letter-spacing: -.2px; }
.actions { display: flex; gap: 4px; }
.actions button {
    padding: 4px 10px; border-radius: 6px; border: 1px solid var(--border);
    background: transparent; color: var(--text2); font-size: 12px;
    cursor: pointer; transition: all .15s; font-family: inherit;
}
.actions button:hover { background: var(--surface2); }
.actions .edit-btn:hover { border-color: var(--accent); color: var(--accent); }
.actions .del-btn:hover { border-color: var(--red); color: var(--red); }
.actions .ping-btn {
    background: rgba(34,197,94,.12); border-color: transparent;
    color: var(--green); font-size: 11px; padding: 4px 8px;
}
.actions .ping-btn:hover { background: rgba(34,197,94,.2); }

/* modal */
.modal {
    display: none; position: fixed; z-index: 99; left: 0; top: 0;
    width: 100%; height: 100%; background: rgba(0,0,0,.65);
    backdrop-filter: blur(4px); animation: fadeIn .2s;
}
.modal-content {
    background: var(--surface); margin: 4% auto; padding: 28px;
    border-radius: 14px; width: 92%; max-width: 520px;
    border: 1px solid var(--border); box-shadow: var(--shadow);
    animation: slideUp .25s ease-out;
}
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
@keyframes slideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
.modal h2 {
    font-size: 18px; font-weight: 700; color: var(--heading);
    margin-bottom: 6px;
}
.modal .modal-sub { color: var(--text2); font-size: 13px; margin-bottom: 20px; }
.modal label {
    display: block; margin-top: 14px; font-size: 12px; font-weight: 600;
    color: var(--text2); text-transform: uppercase; letter-spacing: .3px;
}
.modal input, .modal select {
    width: 100%; padding: 10px 12px; margin-top: 4px;
    border: 1px solid var(--border); border-radius: 8px;
    background: var(--bg); color: var(--text); font-family: inherit;
    font-size: 14px; transition: border-color .2s;
}
.modal input:focus, .modal select:focus { outline: none; border-color: var(--accent); }
.modal .row2 {
    display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px;
}
.modal .row3 {
    display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 10px;
}
.modal-buttons {
    margin-top: 24px; display: flex; gap: 10px; justify-content: end;
}
.modal-buttons button {
    padding: 10px 22px; border-radius: 8px; border: none;
    font-family: inherit; font-size: 14px; font-weight: 500;
    cursor: pointer; transition: all .2s;
}

/* toast */
#toast {
    visibility: hidden; min-width: 260px; background: var(--surface2);
    color: var(--heading); text-align: center; border-radius: 10px;
    padding: 14px 24px; position: fixed; bottom: 30px; left: 50%;
    transform: translateX(-50%); z-index: 100;
    border: 1px solid var(--border); box-shadow: var(--shadow);
    font-size: 14px; font-weight: 500;
}
#toast.show { visibility: visible; animation: toastIn .3s, toastOut .3s 2.7s; }
@keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(10px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
@keyframes toastOut { from { opacity: 1; } to { opacity: 0; } }

.spinner {
    display: inline-block; width: 14px; height: 14px;
    border: 2px solid rgba(255,255,255,.2); border-top-color: #fff;
    border-radius: 50%; animation: spin .5s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }

@media (max-width: 768px) {
    body { padding: 12px; }
    .header h1 { font-size: 18px; }
    .grid { grid-template-columns: repeat(3, 1fr); }
    .stat-card { padding: 14px 10px; }
    .stat-card .num { font-size: 22px; }
    table { font-size: 12px; }
    th, td { padding: 8px 6px; }
    .modal-content { margin: 10% auto; padding: 20px; }
    .modal .row2, .modal .row3 { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<div class="container">

<div class="header">
    <h1><span>VoIP</span> Inventário</h1>
    <span class="badge">{{ data|length }} registros</span>
</div>

<div class="grid">
    <div class="stat-card"><div class="num blue">{{ stats.total }}</div><div class="label">Total</div></div>
    <div class="stat-card"><div class="num green">{{ stats.online }}</div><div class="label">Online</div></div>
    <div class="stat-card"><div class="num red">{{ stats.offline }}</div><div class="label">Offline</div></div>
    <div class="stat-card"><div class="num blue">{{ stats.setores }}</div><div class="label">Setores</div></div>
    <div class="stat-card"><div class="num blue">{{ stats.modelos }}</div><div class="label">Modelos</div></div>
</div>

<div class="table-wrap">
<div class="toolbar" style="padding:12px 16px 0 16px;">
    <input type="text" id="searchInput" placeholder="Buscar IP, setor, telefone, MAC..." oninput="filterTable()">
    <select id="statusFilter" onchange="filterTable()">
        <option value="">Status</option>
        <option value="ONLINE">Online</option>
        <option value="OFFLINE">Offline</option>
    </select>
    <select id="modeloFilter" onchange="filterTable()">
        <option value="">Modelo</option>
        {% for m in modelos %}
        <option value="{{ m }}">{{ m }}</option>
        {% endfor %}
    </select>
    <button class="btn btn-primary" onclick="openAdd()">+ Novo</button>
    <button class="btn btn-green" onclick="runPing(this)">Ping em massa</button>
    <input type="password" id="pingSenha" placeholder="Senha" style="max-width:120px">
</div>
<div class="table-scroll">
<table>
<thead>
<tr>
    <th onclick="sortTable(0)">IP</th>
    <th onclick="sortTable(1)">Setor</th>
    <th onclick="sortTable(2)">Telefone</th>
    <th onclick="sortTable(3)">Status</th>
    <th onclick="sortTable(4)">Modelo</th>
    <th>MAC</th>
    <th>Serial</th>
    <th>Ações</th>
</tr>
</thead>
<tbody id="tableBody">
{% for r in data %}
<tr>
    <td><a href="http://{{ r.IP }}" target="_blank">{{ r.IP }}</a></td>
    <td>{{ r.Setor }}</td>
    <td>{{ r.Telefone }}</td>
    <td class="status-{{ r.Status|lower }}">{{ r.Status }}</td>
    <td>{{ r["Nome do Produto"] or "" }}</td>
    <td class="mono">{{ r.MAC }}</td>
    <td class="mono">{{ r.Serial }}</td>
    <td>
        <div class="actions">
            <button class="edit-btn" onclick="openEdit('{{ r.IP|e }}')">Editar</button>
            <button class="del-btn" onclick="openRemove('{{ r.IP|e }}')">Remover</button>
            <button class="ping-btn" onclick="pingIp(this, '{{ r.IP }}')">Ping</button>
        </div>
    </td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
</div>
</div>

<div id="addModal" class="modal">
<div class="modal-content">
    <h2>Novo Telefone</h2>
    <p class="modal-sub">Preencha os campos para adicionar um novo dispositivo</p>
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
        <button class="btn-ghost" style="padding:10px 22px;border-radius:8px;border:1px solid var(--border);background:transparent;color:var(--text2);cursor:pointer;font-family:inherit;font-size:14px;" onclick="closeModal('addModal')">Cancelar</button>
        <button class="btn-primary" style="padding:10px 22px;border-radius:8px;border:none;font-family:inherit;font-size:14px;font-weight:500;cursor:pointer;background:var(--accent);color:#fff;" onclick="submitAdd()">Salvar</button>
    </div>
</div>
</div>

<div id="editModal" class="modal">
<div class="modal-content">
    <h2>Editar Telefone</h2>
    <p class="modal-sub" id="editIpDisplay"></p>
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
        <button class="btn-ghost" style="padding:10px 22px;border-radius:8px;border:1px solid var(--border);background:transparent;color:var(--text2);cursor:pointer;font-family:inherit;font-size:14px;" onclick="closeModal('editModal')">Cancelar</button>
        <button class="btn-primary" style="padding:10px 22px;border-radius:8px;border:none;font-family:inherit;font-size:14px;font-weight:500;cursor:pointer;background:var(--accent);color:#fff;" onclick="submitEdit()">Salvar</button>
    </div>
</div>
</div>

<div id="removeModal" class="modal">
<div class="modal-content">
    <h2>Remover Telefone</h2>
    <p class="modal-sub">Tem certeza que deseja remover <strong id="removeIpDisplay" style="color:var(--heading)"></strong>? Esta ação não pode ser desfeita.</p>
    <label>Senha *</label><input type="password" id="removeSenha" placeholder="Digite a senha">
    <div class="modal-buttons">
        <button class="btn-ghost" style="padding:10px 22px;border-radius:8px;border:1px solid var(--border);background:transparent;color:var(--text2);cursor:pointer;font-family:inherit;font-size:14px;" onclick="closeModal('removeModal')">Cancelar</button>
        <button style="padding:10px 22px;border-radius:8px;border:none;font-family:inherit;font-size:14px;font-weight:500;cursor:pointer;background:var(--red);color:#fff;" onclick="submitRemove()">Remover</button>
    </div>
</div>
</div>

<div id="toast"></div>

<script>
const data = {{ data | tojson }};

function showToast(msg, isOk=true) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.style.borderColor = isOk ? '#22c55e' : '#ef4444';
    t.className = 'show';
    setTimeout(() => t.className = '', 3000);
}

function filterTable() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    const st = document.getElementById('statusFilter').value;
    const mo = document.getElementById('modeloFilter').value;
    document.querySelectorAll('#tableBody tr').forEach(r => {
        const text = r.textContent.toLowerCase();
        const status = r.children[3]?.textContent.trim() || '';
        let show = true;
        if (q && !text.includes(q)) show = false;
        if (st && status !== st) show = false;
        if (mo && !text.includes(mo.toLowerCase())) show = false;
        r.style.display = show ? '' : 'none';
    });
}

function sortTable(n) {
    const tbody = document.getElementById('tableBody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const dir = tbody.dataset.sortDir === 'asc' ? 'desc' : 'asc';
    tbody.dataset.sortDir = dir;
    rows.sort((a, b) => {
        const va = a.children[n].textContent.trim();
        const vb = b.children[n].textContent.trim();
        const na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return dir === 'asc' ? na - nb : nb - na;
        return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    rows.forEach(r => tbody.appendChild(r));
}

function openAdd() {
    document.getElementById('addModal').style.display = 'block';
}

function openEdit(ip) {
    document.getElementById('editIpDisplay').textContent = 'IP: ' + ip;
    const r = data.find(d => d.IP === ip);
    if (r) {
        document.getElementById('editSetor').value = r.Setor || '';
        document.getElementById('editTel').value = r.Telefone || '';
        document.getElementById('editStatus').value = r.Status || 'ONLINE';
        document.getElementById('editModelo').value = r['Nome do Produto'] || '';
        document.getElementById('editMac').value = r.MAC || '';
        document.getElementById('editSerial').value = r.Serial || '';
        document.getElementById('editFirmware').value = r.Firmware || '';
    }
    document.getElementById('editSenha').value = '';
    document.getElementById('editModal').style.display = 'block';
}

function openRemove(ip) {
    document.getElementById('removeIpDisplay').textContent = ip;
    document.getElementById('removeSenha').value = '';
    document.getElementById('removeModal').style.display = 'block';
}

function closeModal(id) {
    document.getElementById(id).style.display = 'none';
}

window.onclick = function(e) {
    ['addModal','editModal','removeModal'].forEach(id => {
        if (e.target == document.getElementById(id)) closeModal(id);
    });
};

async function api(url, body) {
    const r = await fetch(url, {
        method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)
    });
    return r.json();
}

async function submitAdd() {
    const res = await api('/api/add', {
        ip: document.getElementById('addIp').value,
        setor: document.getElementById('addSetor').value,
        telefone: document.getElementById('addTel').value,
        modelo: document.getElementById('addModelo').value,
        mac: document.getElementById('addMac').value,
        serial: document.getElementById('addSerial').value,
        hardware: document.getElementById('addHardware').value,
        firmware: document.getElementById('addFirmware').value,
        senha: document.getElementById('addSenha').value,
    });
    showToast(res.message, res.ok);
    if (res.ok) setTimeout(() => location.reload(), 800);
}

async function submitEdit() {
    const ip = document.getElementById('editIpDisplay').textContent.replace('IP: ','');
    const res = await api('/api/edit', {
        ip: ip,
        setor: document.getElementById('editSetor').value,
        telefone: document.getElementById('editTel').value,
        status: document.getElementById('editStatus').value,
        modelo: document.getElementById('editModelo').value,
        mac: document.getElementById('editMac').value,
        serial: document.getElementById('editSerial').value,
        firmware: document.getElementById('editFirmware').value,
        senha: document.getElementById('editSenha').value,
    });
    showToast(res.message, res.ok);
    if (res.ok) setTimeout(() => location.reload(), 800);
}

async function submitRemove() {
    const ip = document.getElementById('removeIpDisplay').textContent;
    const res = await api('/api/remove', {
        ip: ip, senha: document.getElementById('removeSenha').value
    });
    showToast(res.message, res.ok);
    if (res.ok) setTimeout(() => location.reload(), 800);
}

async function pingIp(btn, ip) {
    const senha = document.getElementById('pingSenha').value;
    if (!senha) { showToast('Digite a senha no campo ao lado', false); return; }
    btn.textContent = '...';
    btn.disabled = true;
    const res = await api('/api/ping', {ip: ip, senha: senha});
    btn.textContent = 'Ping';
    btn.disabled = false;
    showToast(res.message, res.ok);
    if (res.novo_status) {
        const row = btn.closest('tr');
        const statusCell = row.querySelector('td:nth-child(4)');
        statusCell.textContent = res.novo_status;
        statusCell.className = 'status-' + res.novo_status.toLowerCase();
    }
}

async function runPing(btn) {
    const senha = document.getElementById('pingSenha').value;
    if (!senha) { showToast('Digite a senha no campo ao lado', false); return; }
    if (!confirm('Executar ping em todos os telefones? Isso pode levar alguns minutos.')) return;
    btn.innerHTML = '<span class="spinner"></span> Pingando...';
    btn.disabled = true;
    const res = await api('/api/ping-all', {senha: senha});
    showToast(res.message, true);
    btn.textContent = '🔄 Ping em massa';
    btn.disabled = false;
    if (res.ok) setTimeout(() => location.reload(), 1000);
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    db = get_db()
    data = db.list_all()
    s = db.stats()
    modelos = sorted(set(r.get("Nome do Produto", "") for r in data if r.get("Nome do Produto")))
    return render_template_string(
        HTML,
        data=data,
        stats={"total": s["total"], "online": s["online"], "offline": s["offline"],
               "setores": s["setores_unicos"], "modelos": len(s["modelos"])},
        modelos=modelos,
        atualizacao=os.path.getmtime(str(EXCEL_PATH)),
    )


@app.route("/api/add", methods=["POST"])
def api_add():
    body = request.get_json()
    if body.get("senha") != SENHA:
        return {"ok": False, "message": "Senha incorreta!"}
    if not body.get("ip") or not body.get("setor") or not body.get("telefone"):
        return {"ok": False, "message": "IP, Setor e Telefone são obrigatórios!"}
    db = get_db()
    db.add({
        "IP": body["ip"], "Setor": body["setor"], "Telefone": body["telefone"],
        "Status": "ONLINE", "Nome do Produto": body.get("modelo", ""),
        "MAC": body.get("mac", ""), "Serial": body.get("serial", ""),
        "Hardware": body.get("hardware", ""), "Firmware": body.get("firmware", ""),
    })
    return {"ok": True, "message": f"Telefone {body['ip']} adicionado!"}


@app.route("/api/edit", methods=["POST"])
def api_edit():
    body = request.get_json()
    if body.get("senha") != SENHA:
        return {"ok": False, "message": "Senha incorreta!"}
    db = get_db()
    updates = {}
    for k, v in {"Setor": "setor", "Telefone": "telefone", "Status": "status",
                  "Nome do Produto": "modelo", "MAC": "mac", "Serial": "serial",
                  "Firmware": "firmware"}.items():
        if body.get(v):
            updates[k] = body[v]
    if db.update(body["ip"], updates):
        return {"ok": True, "message": f"Telefone {body['ip']} atualizado!"}
    return {"ok": False, "message": "Telefone não encontrado!"}


@app.route("/api/remove", methods=["POST"])
def api_remove():
    body = request.get_json()
    if body.get("senha") != SENHA:
        return {"ok": False, "message": "Senha incorreta!"}
    db = get_db()
    if db.delete(body["ip"]):
        return {"ok": True, "message": f"Telefone {body['ip']} removido!"}
    return {"ok": False, "message": "Telefone não encontrado!"}


@app.route("/api/ping", methods=["POST"])
def api_ping():
    body = request.get_json()
    if body.get("senha") != SENHA:
        return {"ok": False, "message": "Senha incorreta!"}
    ip = body.get("ip", "")
    param = "-n" if platform.system().lower() == "windows" else "-c"
    try:
        result = subprocess.run(["ping", param, "1", ip], capture_output=True, text=True, timeout=10)
        novo = "ONLINE" if result.returncode == 0 else "OFFLINE"
        db = get_db()
        db.update(ip, {"Status": novo})
        return {"ok": True, "message": f"{ip} está {novo}", "novo_status": novo}
    except subprocess.TimeoutExpired:
        return {"ok": False, "message": f"{ip} TIMEOUT"}


@app.route("/api/ping-all", methods=["POST"])
def api_ping_all():
    body = request.get_json()
    if not body or body.get("senha") != SENHA:
        return {"ok": False, "message": "Senha incorreta!"}
    db = get_db()
    data = db.list_all()
    param = "-n" if platform.system().lower() == "windows" else "-c"
    online = 0
    offline = 0
    for r in data:
        ip = r.get("IP", "")
        if not ip:
            continue
        try:
            result = subprocess.run(["ping", param, "1", ip], capture_output=True, text=True, timeout=5)
            novo = "ONLINE" if result.returncode == 0 else "OFFLINE"
            db.update(ip, {"Status": novo})
            if result.returncode == 0:
                online += 1
            else:
                offline += 1
        except subprocess.TimeoutExpired:
            offline += 1
    return {"ok": True, "message": f"Ping concluído! Online: {online} | Offline: {offline}"}


if __name__ == "__main__":
    print(f"Iniciando servidor em http://localhost:5000")
    print(f"Planilha: {EXCEL_PATH}")
    app.run(debug=True, host="0.0.0.0", port=5000)
