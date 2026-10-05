# Integration hand-off

## Backend
- Project folder: `api/`
- Install/build: `npm --prefix api install` then `npm --prefix api run build`
- Local dev: `npm --prefix api run dev`
- Local port: `7071`
- Health endpoint: `GET /api/health`
- Build artifacts: `api/dist/`

## Frontend
- Project folder: `frontend/`
- Install/build: `npm --prefix frontend install` then `npm --prefix frontend run build`
- Local dev: `npm --prefix frontend run dev`
- Local port: `4173`
- API seam to swap: `frontend/assets/app.js` currently loads local sample data from `./data/sample-overview.json`; replace with live fetches to API endpoints from a single wrapper module or fetch helper.
- Mock files to eliminate after live wiring: `frontend/data/sample-overview.json`, `frontend/assets/app.js` local sample rendering, any direct local-only state in page logic.

## API routes
- `GET /api/health`
- `GET /api/compliance`
- `GET /api/compliance/{id}`
- `POST /api/sync`
- `POST /api/review/complete`
- `GET /api/summary`
- `GET /api/ops/admin` (the Functions host reserves `/api/admin`)
- `POST /api/refresh`

## Database / persistence
- Data store type: Azure Blob Storage + Microsoft Graph-driven data sync
- Migration tool: none for this initial scaffold; no seed data and no schema migration required
- Environment variables: `AZURE_FUNCTIONS_ENVIRONMENT`, `AzureWebJobsStorage`, `STORAGE_CONNECTION_STRING`, `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `SHAREPOINT_SITE_ID`

## Shared types
- Shared package: not yet required in this scaffold; keep domain data contracts in API and frontend adapters until live integration is completed.

## Services
- Essential: Azure Static Web Apps, Azure Blob Storage, Azure Functions runtime
- Enhancement: Microsoft Graph sync, SharePoint/Excel ingestion, summary analytics

## Verification checklist
- `npm --prefix api run build` passes
- `npm --prefix frontend run build` passes
- Backend responds on `GET /api/health`
- Frontend loads dashboard shell and regional stats
- No seed data is introduced during integration

## Integration results
- Migrations: not applicable. This project uses Azure Blob Storage and Microsoft Graph, with no SQL/PostgreSQL database.
- Backend: all eight registered routes smoke-tested. `GET /api/health`, list, summary, and ops/admin returned 200; detail returned the expected 404 for an unknown id; sync and refresh returned 202; review completion returned 200.
- Route mapping: administrative status is available at `GET /api/ops/admin`; `/api/admin` conflicts with a reserved Azure Functions route.
- Frontend: removed `frontend/data/sample-overview.json`; the dashboard now fetches `/api/compliance` and `/api/summary`, and refresh calls `/api/refresh`.
- End-to-end: frontend on port 4174 proxied `GET /api/compliance` and `GET /api/summary` to the Functions host; both returned 200.
- Data limitation: the scaffold does not yet retrieve or persist compliance records from SharePoint/Graph. The API currently returns an empty collection, and sync/review handlers are stubs; the dashboard displays that empty live response rather than sample records.
