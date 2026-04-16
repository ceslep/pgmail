"""
Financial Report Generator — Emails from info@pasarelapagosaval.com

Fetches all emails, extracts transaction data, saves JSON,
generates HTML report that reads from JSON.

Usage:
    python generate_report.py
"""

import os
import re
import json
import base64
import html as html_module
import webbrowser
from datetime import datetime

from dateutil import parser as dateutil_parser
from gmail_reader import authenticate
from googleapiclient.discovery import build

SENDER_QUERY = "from:info@pasarelapagosaval.com"
OUTPUT_JSON = "transacciones.json"
OUTPUT_HTML = "informe_financiero.html"


# ---------------------------------------------------------------------------
# Gmail fetching
# ---------------------------------------------------------------------------

def fetch_all_messages(service, query):
    """Fetch all message IDs matching query, handling pagination."""
    messages = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        batch = result.get("messages", [])
        messages.extend(batch)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return messages


def extract_body(payload):
    """Recursively extract plain text body from MIME payload."""
    mime = payload.get("mimeType", "")
    parts = payload.get("parts", [])

    if mime == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

    for part in parts:
        text = extract_body(part)
        if text:
            return text

    if mime == "text/html":
        data = payload.get("body", {}).get("data", "")
        if data:
            raw_html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
            return strip_html(raw_html)

    for part in parts:
        if part.get("mimeType") == "text/html":
            data = part.get("body", {}).get("data", "")
            if data:
                raw_html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
                return strip_html(raw_html)

    return ""


def strip_html(text):
    """Remove HTML tags, decode entities."""
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_module.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def get_full_message(service, msg_id):
    """Fetch message, extract headers + body."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}

    body = extract_body(msg.get("payload", {}))
    if not body:
        body = msg.get("snippet", "")

    return {
        "id": msg_id,
        "subject": headers.get("Subject", "(sin asunto)"),
        "from": headers.get("From", ""),
        "date_str": headers.get("Date", ""),
        "body": body,
        "snippet": msg.get("snippet", ""),
    }


# ---------------------------------------------------------------------------
# Transaction parsing
# ---------------------------------------------------------------------------

RE_AMOUNT = re.compile(
    r"(?:"
    r"\$\s?[\d.,]+|"
    r"COP\s?[\d.,]+|"
    r"[\d.,]+\s?COP|"
    r"USD\s?[\d.,]+|"
    r"[\d.,]+\s?USD"
    r")",
    re.IGNORECASE,
)

RE_REFERENCE = re.compile(
    r"(?:"
    r"(?:referencia|ref|reference|transacci[oó]n|transaction|ticket|recibo|receipt|CUS|id)"
    r"[\s:#.\-]*"
    r"([A-Za-z0-9\-]{4,30})"
    r")",
    re.IGNORECASE,
)

STATUS_MAP = {
    "aprobad": "Aprobada",
    "aprobado": "Aprobada",
    "aprobada": "Aprobada",
    "exitos": "Aprobada",
    "successful": "Aprobada",
    "approved": "Aprobada",
    "rechazad": "Rechazada",
    "rechazado": "Rechazada",
    "rechazada": "Rechazada",
    "declined": "Rechazada",
    "failed": "Rechazada",
    "fallid": "Rechazada",
    "fallido": "Rechazada",
    "fallida": "Rechazada",
    "pendiente": "Pendiente",
    "pending": "Pendiente",
    "en proceso": "Pendiente",
    "processing": "Pendiente",
    "cancelad": "Cancelada",
    "cancelado": "Cancelada",
    "cancelada": "Cancelada",
    "cancelled": "Cancelada",
    "canceled": "Cancelada",
    "reversad": "Reversada",
    "reversado": "Reversada",
    "reversada": "Reversada",
    "reversed": "Reversada",
}


def parse_amount(text):
    match = RE_AMOUNT.search(text)
    if not match:
        return None
    raw = match.group(0)
    cleaned = re.sub(r"[A-Za-z$\s]", "", raw)
    if "," in cleaned and "." in cleaned:
        if cleaned.rindex(",") > cleaned.rindex("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        parts = cleaned.split(",")
        if len(parts[-1]) == 2 and len(parts) == 2:
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_reference(text):
    match = RE_REFERENCE.search(text)
    return match.group(1) if match else ""


def detect_status(text):
    text_lower = text.lower()
    for keyword, status in STATUS_MAP.items():
        if keyword in text_lower:
            return status
    return "Desconocido"


def parse_date(date_str):
    try:
        return dateutil_parser.parse(date_str, fuzzy=True)
    except (ValueError, OverflowError):
        return None


def parse_transaction(msg):
    combined = f"{msg['subject']} {msg['body']}"
    dt = parse_date(msg["date_str"])
    amount = parse_amount(combined)
    reference = parse_reference(combined)
    status = detect_status(combined)

    return {
        "date_iso": dt.strftime("%Y-%m-%dT%H:%M:%S") if dt else None,
        "date_formatted": dt.strftime("%Y-%m-%d %H:%M") if dt else msg["date_str"],
        "subject": msg["subject"],
        "amount": amount,
        "reference": reference or None,
        "status": status,
        "snippet": msg["snippet"][:300],
    }


# ---------------------------------------------------------------------------
# HTML generation (reads from JSON)
# ---------------------------------------------------------------------------

def generate_html(data_json):
    """Generate self-contained HTML with JSON embedded inline."""
    return """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe Financiero — Pasarela Pago Aval</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f7fa;
            color: #2c3e50;
            line-height: 1.6;
        }
        .header {
            background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
            color: white;
            padding: 40px 60px;
        }
        .header h1 { font-size: 28px; font-weight: 300; margin-bottom: 8px; }
        .header h2 { font-size: 16px; font-weight: 400; opacity: 0.85; }
        .header .meta { margin-top: 16px; font-size: 13px; opacity: 0.7; }
        .container { max-width: 1400px; margin: 0 auto; padding: 30px 40px; }

        .highlight-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px; margin-bottom: 30px;
        }
        .highlight {
            background: white; border-radius: 10px; padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .highlight-label {
            font-size: 13px; color: #95a5a6; text-transform: uppercase;
            letter-spacing: 1px; margin-bottom: 8px;
        }
        .highlight-value { font-size: 24px; font-weight: 600; color: #2c3e50; }

        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px; margin-bottom: 30px;
        }
        .card {
            background: white; border-radius: 10px; padding: 24px;
            text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            transition: transform 0.2s;
        }
        .card:hover { transform: translateY(-2px); }
        .card-value { font-size: 36px; font-weight: 700; }
        .card-label { font-size: 14px; color: #7f8c8d; margin-top: 4px; }
        .card-pct { font-size: 12px; color: #bdc3c7; margin-top: 2px; }

        h3 { font-size: 20px; margin-bottom: 16px; color: #2c3e50; }

        .filters {
            background: white; border-radius: 10px; padding: 20px 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 20px;
            display: flex; gap: 16px; flex-wrap: wrap; align-items: center;
        }
        .filters label { font-size: 13px; color: #7f8c8d; font-weight: 500; }
        .filters input, .filters select {
            padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px;
            font-size: 14px; font-family: inherit;
        }
        .filters input:focus, .filters select:focus { outline: none; border-color: #3949ab; }
        .filter-group { display: flex; flex-direction: column; gap: 4px; }
        .filter-count { font-size: 13px; color: #95a5a6; margin-left: auto; }
        .btn {
            padding: 8px 16px; border: 1px solid #ddd; border-radius: 6px;
            cursor: pointer; background: #f8f9fa; font-family: inherit; font-size: 14px;
        }
        .btn:hover { background: #e9ecef; }

        .table-wrapper {
            background: white; border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            overflow-x: auto; margin-bottom: 30px;
        }
        table { width: 100%; border-collapse: collapse; font-size: 14px; }
        th {
            background: #34495e; color: white; padding: 14px 16px;
            text-align: left; font-weight: 500; white-space: nowrap;
            cursor: pointer; user-select: none;
        }
        th:hover { background: #2c3e50; }
        th .sort-arrow { margin-left: 6px; font-size: 10px; }
        td { padding: 12px 16px; border-bottom: 1px solid #ecf0f1; }
        tr:hover td { background: #f8f9fa; }
        .amount { font-family: 'Courier New', monospace; text-align: right; font-weight: 600; }
        .badge {
            display: inline-block; padding: 4px 12px; border-radius: 20px;
            color: white; font-size: 12px; font-weight: 500;
        }
        .snippet {
            max-width: 300px; overflow: hidden; text-overflow: ellipsis;
            white-space: nowrap; font-size: 12px; color: #95a5a6;
        }
        .footer { text-align: center; padding: 30px; color: #bdc3c7; font-size: 12px; }
        .loading { text-align: center; padding: 60px; color: #95a5a6; font-size: 18px; }

        @media print {
            body { background: white; }
            .header { background: #1a237e !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .filters { display: none; }
            .card, .highlight, .table-wrapper { box-shadow: none; border: 1px solid #eee; }
            .badge { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
        @media (max-width: 768px) {
            .header { padding: 24px; }
            .container { padding: 16px; }
            .summary { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>

<div class="header">
    <h1>Informe Financiero</h1>
    <h2>Transacciones Aprobadas — Pasarela Pago Aval</h2>
    <div class="meta" id="headerMeta">Cargando datos...</div>
</div>

<div class="container">
    <div class="highlight-cards" id="highlightCards"></div>

    <h3>Resumen por Estado</h3>
    <div class="summary" id="statusCards"></div>

    <h3>Detalle de Transacciones</h3>
    <div class="filters">
        <div class="filter-group">
            <label>Buscar</label>
            <input type="text" id="filterSearch" placeholder="Asunto, referencia..." oninput="renderTable()">
        </div>
        <div class="filter-group">
            <label>Fecha desde</label>
            <input type="date" id="filterDateFrom" onchange="renderTable()">
        </div>
        <div class="filter-group">
            <label>Fecha hasta</label>
            <input type="date" id="filterDateTo" onchange="renderTable()">
        </div>
        <div class="filter-group">
            <label>Monto min</label>
            <input type="number" id="filterAmountMin" placeholder="0" oninput="renderTable()">
        </div>
        <div class="filter-group">
            <label>Monto max</label>
            <input type="number" id="filterAmountMax" placeholder="999999" oninput="renderTable()">
        </div>
        <div class="filter-group">
            <label>Estado</label>
            <select id="filterStatus" onchange="renderTable()">
                <option value="">Todos</option>
            </select>
        </div>
        <div class="filter-group">
            <label>&nbsp;</label>
            <button class="btn" onclick="clearFilters()">Limpiar</button>
        </div>
        <span class="filter-count" id="filterCount"></span>
    </div>
    <div class="table-wrapper">
        <table>
            <thead>
                <tr>
                    <th onclick="sortBy('index')">#<span class="sort-arrow" id="sort-index"></span></th>
                    <th onclick="sortBy('date')">Fecha<span class="sort-arrow" id="sort-date"></span></th>
                    <th onclick="sortBy('subject')">Asunto<span class="sort-arrow" id="sort-subject"></span></th>
                    <th onclick="sortBy('amount')">Monto<span class="sort-arrow" id="sort-amount"></span></th>
                    <th onclick="sortBy('reference')">Referencia<span class="sort-arrow" id="sort-reference"></span></th>
                    <th onclick="sortBy('status')">Estado<span class="sort-arrow" id="sort-status"></span></th>
                    <th>Detalle</th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <tr><td colspan="7" class="loading">Cargando transacciones...</td></tr>
            </tbody>
        </table>
    </div>
</div>

<div class="footer" id="footer"></div>

<script>
const REPORT_DATA = """ + data_json + """;

const STATUS_COLORS = {
    'Aprobada': '#27ae60',
    'Rechazada': '#e74c3c',
    'Pendiente': '#f39c12',
    'Cancelada': '#95a5a6',
    'Reversada': '#8e44ad',
    'Desconocido': '#7f8c8d'
};

let allData = [];
let currentSort = { key: 'date', asc: false };

async function loadData() {
    try {
        const json = REPORT_DATA;
        allData = json.transactions;

        // Header meta
        document.getElementById('headerMeta').innerHTML =
            `Cuenta: ${esc(json.email_account)} &nbsp;|&nbsp; ` +
            `Remitente: info@pasarelapagosaval.com &nbsp;|&nbsp; ` +
            `Generado: ${json.generated_at}`;

        // Populate status filter
        const statuses = [...new Set(allData.map(t => t.status))].sort();
        const sel = document.getElementById('filterStatus');
        statuses.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s; opt.textContent = s;
            sel.appendChild(opt);
        });

        renderSummary();
        renderTable();
    } catch (e) {
        document.getElementById('tableBody').innerHTML =
            '<tr><td colspan="7" class="loading">Error cargando transacciones.json: ' + e.message + '</td></tr>';
    }
}

function renderSummary() {
    const total = allData.length;
    const amounts = allData.filter(t => t.amount !== null).map(t => t.amount);
    const totalAmount = amounts.reduce((a, b) => a + b, 0);
    const avgAmount = amounts.length ? totalAmount / amounts.length : 0;

    const dates = allData.filter(t => t.date_iso).map(t => t.date_iso).sort();
    const dateStart = dates.length ? dates[0].substring(0, 10) : '—';
    const dateEnd = dates.length ? dates[dates.length - 1].substring(0, 10) : '—';

    document.getElementById('highlightCards').innerHTML = `
        <div class="highlight">
            <div class="highlight-label">Total Transacciones</div>
            <div class="highlight-value">${total}</div>
        </div>
        <div class="highlight">
            <div class="highlight-label">Monto Total</div>
            <div class="highlight-value">$${fmtNum(totalAmount)}</div>
        </div>
        <div class="highlight">
            <div class="highlight-label">Monto Promedio</div>
            <div class="highlight-value">$${fmtNum(avgAmount)}</div>
        </div>
        <div class="highlight">
            <div class="highlight-label">Periodo</div>
            <div class="highlight-value" style="font-size:18px">${dateStart} — ${dateEnd}</div>
        </div>`;

    // Status cards
    const counts = {};
    allData.forEach(t => { counts[t.status] = (counts[t.status] || 0) + 1; });
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

    document.getElementById('statusCards').innerHTML = sorted.map(([status, count]) => {
        const color = STATUS_COLORS[status] || '#7f8c8d';
        const pct = total ? (count / total * 100).toFixed(1) : '0.0';
        return `<div class="card">
            <div class="card-value" style="color:${color}">${count}</div>
            <div class="card-label">${status}</div>
            <div class="card-pct">${pct}%</div>
        </div>`;
    }).join('');
}

function getFiltered() {
    const search = document.getElementById('filterSearch').value.toLowerCase();
    const dateFrom = document.getElementById('filterDateFrom').value;
    const dateTo = document.getElementById('filterDateTo').value;
    const amountMin = parseFloat(document.getElementById('filterAmountMin').value) || 0;
    const amountMax = parseFloat(document.getElementById('filterAmountMax').value) || Infinity;
    const statusFilter = document.getElementById('filterStatus').value;

    return allData.filter(t => {
        const dateStr = (t.date_iso || '').substring(0, 10);
        const amount = t.amount || 0;
        const text = `${t.subject} ${t.reference || ''} ${t.snippet}`.toLowerCase();

        if (search && !text.includes(search)) return false;
        if (dateFrom && dateStr < dateFrom) return false;
        if (dateTo && dateStr > dateTo) return false;
        if (amount < amountMin) return false;
        if (amount > amountMax) return false;
        if (statusFilter && t.status !== statusFilter) return false;
        return true;
    });
}

function renderTable() {
    let filtered = getFiltered();

    // Sort
    const key = currentSort.key;
    filtered.sort((a, b) => {
        let va, vb;
        if (key === 'date') { va = a.date_iso || ''; vb = b.date_iso || ''; }
        else if (key === 'amount') { va = a.amount || 0; vb = b.amount || 0; }
        else if (key === 'subject') { va = a.subject.toLowerCase(); vb = b.subject.toLowerCase(); }
        else if (key === 'reference') { va = (a.reference || '').toLowerCase(); vb = (b.reference || '').toLowerCase(); }
        else if (key === 'status') { va = a.status; vb = b.status; }
        else { va = 0; vb = 0; }

        if (va < vb) return currentSort.asc ? -1 : 1;
        if (va > vb) return currentSort.asc ? 1 : -1;
        return 0;
    });

    // Update sort arrows
    document.querySelectorAll('.sort-arrow').forEach(el => el.textContent = '');
    const arrow = document.getElementById('sort-' + key);
    if (arrow) arrow.textContent = currentSort.asc ? ' ▲' : ' ▼';

    const tbody = document.getElementById('tableBody');
    if (!filtered.length) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:40px;color:#95a5a6;">No se encontraron transacciones</td></tr>';
    } else {
        tbody.innerHTML = filtered.map((t, i) => {
            const color = STATUS_COLORS[t.status] || '#7f8c8d';
            const amountFmt = t.amount !== null ? '$' + fmtNum(t.amount) : '—';
            return `<tr>
                <td>${i + 1}</td>
                <td>${esc(t.date_formatted)}</td>
                <td>${esc(t.subject)}</td>
                <td class="amount">${amountFmt}</td>
                <td>${esc(t.reference || '—')}</td>
                <td><span class="badge" style="background:${color}">${t.status}</span></td>
                <td class="snippet">${esc(t.snippet)}</td>
            </tr>`;
        }).join('');
    }

    document.getElementById('filterCount').textContent =
        filtered.length + ' de ' + allData.length + ' transacciones';

    document.getElementById('footer').innerHTML =
        'Informe generado desde Gmail API &nbsp;|&nbsp; Datos: transacciones.json';
}

function sortBy(key) {
    if (currentSort.key === key) {
        currentSort.asc = !currentSort.asc;
    } else {
        currentSort = { key, asc: true };
    }
    renderTable();
}

function clearFilters() {
    document.getElementById('filterSearch').value = '';
    document.getElementById('filterDateFrom').value = '';
    document.getElementById('filterDateTo').value = '';
    document.getElementById('filterAmountMin').value = '';
    document.getElementById('filterAmountMax').value = '';
    document.getElementById('filterStatus').value = '';
    renderTable();
}

function esc(s) {
    const d = document.createElement('div');
    d.textContent = s || '';
    return d.innerHTML;
}

function fmtNum(n) {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

document.addEventListener('DOMContentLoaded', loadData);
</script>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Autenticando con Gmail...")
    creds = authenticate()
    service = build("gmail", "v1", credentials=creds)

    profile = service.users().getProfile(userId="me").execute()
    email_account = profile["emailAddress"]
    print(f"Conectado: {email_account}")

    print(f"\nBuscando emails de info@pasarelapagosaval.com...")
    msg_refs = fetch_all_messages(service, SENDER_QUERY)
    print(f"Encontrados: {len(msg_refs)} emails")

    if not msg_refs:
        print("No se encontraron emails.")
        msg_refs = []

    # Fetch & parse
    transactions = []
    for i, ref in enumerate(msg_refs, 1):
        if i % 10 == 0 or i == 1:
            print(f"  Procesando {i}/{len(msg_refs)}...")
        msg = get_full_message(service, ref["id"])
        tx = parse_transaction(msg)
        transactions.append(tx)

    print(f"\nTransacciones parseadas: {len(transactions)}")

    # Filter: only approved
    approved = [t for t in transactions if t["status"] == "Aprobada"]
    print(f"Aprobadas: {len(approved)}")

    # Dedup by date+time (keep first occurrence)
    seen_dates = set()
    unique = []
    for t in approved:
        key = t["date_iso"]  # "2024-01-15T10:30:00" — unique per datetime
        if key and key in seen_dates:
            print(f"  Duplicado (misma fecha/hora): {t['date_formatted']} - {t['subject']}")
            continue
        if key:
            seen_dates.add(key)
        unique.append(t)

    print(f"Unicas (sin duplicados): {len(unique)}")

    # Sort by date desc
    unique.sort(key=lambda t: t["date_iso"] or "", reverse=True)

    # Save JSON
    data = {
        "email_account": email_account,
        "sender": "info@pasarelapagosaval.com",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_fetched": len(msg_refs),
        "total_approved": len(approved),
        "total_unique": len(unique),
        "transactions": unique,
    }

    data_json = json.dumps(data, ensure_ascii=False)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nJSON guardado: {os.path.abspath(OUTPUT_JSON)}")

    # Save HTML (JSON embedded inline)
    html_content = generate_html(data_json)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_content)

    abs_path = os.path.abspath(OUTPUT_HTML)
    print(f"HTML guardado: {abs_path}")
    print("Abriendo en navegador...")
    webbrowser.open(f"file:///{abs_path}")


if __name__ == "__main__":
    main()
