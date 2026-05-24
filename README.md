# Emotional Stability Pipeline

Cloud-Pipeline für die **Emotional Stability Analyse™** von Susanne Wachter / Your Balance.

Dieses Repository enthält **nur die Verarbeitungs-Logik** (Scoring, PDF-Bau, Mail-Versand).
Die eigentlichen Inhalte (Fragenkatalog, Typen-Definitionen, Auswertungs-Prompts) sind
proprietär und liegen in einem privaten Netlify-Blob-Store.

## Architektur

```
Käuferin macht Test auf analyse.your-balance.at
        ↓
Antworten landen im Netlify Blob Store (submissions)
        ↓
GitHub Action wacht auf (alle 15 Min ODER sofort getriggert)
        ↓
Pipeline lädt Konfiguration aus privatem Blob Store (config)
        ↓
Scoring → Claude API → DOCX (YB-Brand) → LibreOffice → PDF → SMTP-Mail
```

## Required Secrets

GitHub Settings → Secrets and variables → Actions:

| Name | Beschreibung |
|---|---|
| `ANTHROPIC_API_KEY` | Anthropic Claude API Key |
| `NETLIFY_SITE_ID` | Netlify Site ID |
| `NETLIFY_AUTH_TOKEN` | Netlify Personal Access Token |
| `SMTP_HOST` | z.B. w018758b.kasserver.com (All-Inkl) |
| `SMTP_PORT` | 465 |
| `SMTP_USE_SSL` | true |
| `SMTP_USER` | susanne@your-balance.at |
| `SMTP_PASSWORD` | Mail-Passwort |
| `REPLY_TO` | susanne@your-balance.at |

---

**Owner:** Susanne Wachter · your-balance.at
