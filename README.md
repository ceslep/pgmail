# pgmail — Gmail Financial Report Generator

Reads emails from a specific sender via Gmail API, extracts transaction data, and generates an interactive HTML financial report.

## Features

- **Gmail API integration** — OAuth2 authentication, paginated email fetching
- **Transaction parsing** — Extracts amounts (COP/USD), references, and status from email bodies
- **Interactive HTML report** — Filterable/sortable table, summary cards, responsive design, print-ready
- **Deduplication** — Removes duplicate transactions by date/time

## Setup

### 1. Google Cloud Console

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select existing)
3. Enable **Gmail API**
4. Go to **APIs & Services → Credentials**
5. Create **OAuth 2.0 Client ID** (Application type: Desktop App)
6. Download the JSON file and save it as `credentials.json` in this folder

See `credentials.json.example` for the expected format.

### 2. Configure environment

Copy `.env.example` to `.env` and edit with your values:

```bash
cp .env.example .env
```

```env
GMAIL_ACCOUNT=your-email@gmail.com
SENDER_QUERY=from:sender@example.com
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run

**Read emails (general reader):**
```bash
python gmail_reader.py
```

**Generate financial report:**
```bash
python generate_report.py
```

On first run, a browser window opens for OAuth consent. The token is saved locally as `token.json` for subsequent runs.

## Output

- `transacciones.json` — Parsed transaction data
- `informe_financiero.html` — Self-contained interactive report (opens automatically in browser)

## Project Structure

```
pgmail/
├── gmail_reader.py          # Gmail API auth & basic email reader
├── generate_report.py       # Transaction parser & HTML report generator
├── .env.example             # Environment config template
├── credentials.json.example # Template for OAuth credentials
├── requirements.txt         # Python dependencies
└── .gitignore
```

## Security

`credentials.json`, `token.json`, and `.env` contain secrets and are excluded via `.gitignore`. **Never commit them.**

## License

MIT
