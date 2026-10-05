# Project Plan

**Status**: Integrated
**Created**: 2026-10-03
**Mode**: NEW

---

## 1. Project Overview

**Goal**: Migrate the enterprise Medical Compliance Tracker from Google Apps Script to Azure Static Web Apps with a serverless Node.js Azure Functions API that synchronizes SharePoint and Excel compliance data via Microsoft Graph and exposes a production-quality compliance dashboard. The project is designed so that every module is independently testable.

**App Type**: SPA + API

**API Login**: Yes

**Mode**: NEW

**Deployment Plan**: No deployment plan found

---

## 2. Backend — Azure Functions

| Component | Technology |
|-----------|-----------|
| **Language** | TypeScript |
| **Runtime** | Node |
| **Package Manager** | npm |
| **Test Runner** | vitest |
| **Mocking Library** | vi.mock |
| **Test Command** | npm test |
| **Orchestration** | docker-compose |

> **Language vs Runtime**: `Language` is the source language the user picked in this service's `language` question. `Runtime` is the execution runtime — default `Node` for TypeScript/JavaScript, `CPython` for Python, `.NET` for C#. Only deviate from the default (e.g. `Bun`, `Deno`, `PyPy`) when the user explicitly asks. **Package Manager and Test Runner are language-dependent** — match them to this service's Language (e.g. C# → `dotnet (NuGet)` + `xUnit`/`NUnit`/`MSTest`). The `Orchestration` row is recorded for the scaffold step but hidden in the plan UI — always keep it set to `docker-compose`.

---

## 3. Frontend — Web App

| Component | Technology |
|-----------|-----------|
| **Language** | JavaScript |
| **Framework** | Plain HTML/CSS/JavaScript (no frontend framework) |
| **Package Manager** | npm |
| **Test Runner** | vitest |
| **Mocking Library** | vi.mock |
| **Test Command** | npm test |

---

## 4. Services Required

| Azure Service | Role in App | Environment Variable | Default Value (Local) | Classification |
|---------------|------------|---------------------|----------------------|----------------|
| Azure Static Web Apps | Host the compliance dashboard and serve the static frontend assets | `SWA_API_URL` | `http://localhost:7071` | Essential |
| Azure Blob Storage | Required Azure Functions runtime storage for AzureWebJobsStorage and any uploaded documents or generated exports | `STORAGE_CONNECTION_STRING` | `UseDevelopmentStorage=true` | Essential |

---

## 5. Prerequisites

### Run

| Tool | Service(s) | Installed | Version |
|------|------------|-----------|---------|
| Node.js | * | ✅ | v24.21.0 |
| npm | * | ✅ | 11.19.0 |
| Azure Functions Core Tools | medical-compliance-api | ❓ | not detected |
| Azure Static Web Apps CLI | medical-compliance-frontend | ❓ | not detected |
| Azure CLI | * | ❓ | not detected |

### Debug

| Tool | Service(s) | Installed | Version |
|------|------------|-----------|---------|
| Docker | medical-compliance-api | ✅ | Docker version 29.8.0-1, build 88096ef00576baf72a9cb45caa45c0544c40e0 |
| Docker Compose | medical-compliance-api | ✅ | docker-compose v2.39.4 |
| ms-azuretools.vscode-azurefunctions | medical-compliance-api | ❓ | not detected |
| ms-azuretools.vscode-azurestaticwebapps | medical-compliance-frontend | ❓ | not detected |

---

## 6. Design System & UI

**Component Library**: Pico.css
**Style Direction**: Modern clinical console with calm blue accents, dense operational tables, and clear status emphasis to support quick review of compliance risk and overdue actions.
**Typography**: Inter

### Color Palette

| Token | Hex | Usage |
|-------|-----|-------|
| `primary` | `#0b4f8a` | Primary actions, active navigation, and region health summaries |
| `accent` | `#1f9d8a` | Compliance pass indicators, highlights, and secondary CTAs |
| `surface` | `#f5f8fc` | Main page and panel backgrounds for a calm enterprise workspace |
| `text` | `#111827` | Body copy, titles, and table data |
| `muted` | `#6b7280` | Secondary labels, timestamps, filter captions, and helper text |
| `border` | `#dfe7f1` | Dividers, cards, filter controls, and table borders |

### Pages

| Page | Route | Purpose | Layout |
|------|-------|---------|--------|
| Compliance Overview | `/` | Monitor compliance health across regions and departments | `header + nav + grid + table` |
| Compliance Detail | `/compliance/:id` | Review an employee or site record, remediation status, and history | `two-column(meta+timeline) + action-bar` |
| Review Queue | `/review` | Surface overdue remediations and items requiring action | `header + filters + table` |

### Sample Content

```
Compliance Overview — Region:
| Region | Completion | Open Items | High Risk | Status |
| Northeast | 96% | 14 | 2 | Healthy |
| Southeast | 88% | 29 | 6 | Watch |
| West | 91% | 18 | 3 | Healthy |

Compliance Detail — Record: RN-2147 · Last review: 2026-09-28 · Status: Pending remediation
Review Queue — Priority: High · Department: Radiology · Due: 2026-10-07
```

---

## 7. Project Structure

```
/workspaces/azure_medical
├── .azure/
│   ├── requirements.json
│   ├── project-plan.md
│   └── .preview-temp/
│       ├── manifest.json
│       └── theme.css
├── README.md
├── api/
│   ├── host.json
│   ├── package.json
│   ├── local.settings.json.example
│   └── src/
│       ├── app.ts
│       ├── config.ts
│       ├── services/
│       │   ├── graphClient.ts
│       │   ├── driveSyncService.ts
│       │   └── regionalSyncService.ts
│       └── functions/
│           ├── health.ts
│           ├── compliance.ts
│           ├── sync.ts
│           ├── summary.ts
│           ├── refresh.ts
│           └── admin.ts
├── frontend/
│   ├── index.html
│   ├── staticwebapp.config.json
│   ├── assets/
│   │   ├── app.js
│   │   ├── styles.css
│   │   └── charts.js
│   └── data/
│       └── sample-overview.json
└── package.json
```

---

## 8. Route Definitions

| # | Method | Path | Description | Request Body | Response Body | Status Codes |
|---|--------|------|-------------|-------------|--------------|-------------|
| 1 | GET | `/api/health` | Health check for the API and downstream dependencies | — | `{ status, services }` | 200, 503 |
| 2 | GET | `/api/compliance` | Return a filtered list of compliance records for the dashboard | — | `{ records, total, filters }` | 200 |
| 3 | GET | `/api/compliance/:id` | Fetch one compliance record with review metadata and remediation history | — | `{ record, auditTrail }` | 200, 404 |
| 4 | POST | `/api/sync` | Trigger a Microsoft Graph sync for SharePoint or Excel source files | `{ source, region, since }` | `{ jobId, status, startedAt }` | 202, 400 |
| 5 | POST | `/api/review/complete` | Mark a remediation item as completed and update record state | `{ id, status, reviewer }` | `{ id, status, updatedAt }` | 200, 400 |
| 6 | GET | `/api/summary` | Summarize risk posture by region or department | — | `{ totals, healthyCount, watchCount, overdueCount }` | 200 |

---

## 9. Next Steps

1. Run **azure-project-scaffold** to execute this plan
2. Run **azure-project-integrate** to wire the frontend to live data, smoke-test the backend, and create the migrations
3. Run **azure-debug-plan** → **azure-debug-generate** for Docker emulators and VS Code debugging
4. Run the **azure-deploy** agent when ready; it uses **azure-app-onboard** for architecture, cost estimation, IaC generation, provisioning, and health verification
