# B Hub Payroll Stack

Self-hosted n8n (orchestrator) + bhub-engine (Python + LibreOffice payslip engine).

## Deploy
```bash
docker compose up -d --build
```
- n8n UI:  http://<DROPLET_IP>:5678
- Engine:  internal only; n8n calls it at http://bhub-engine:8000

## Engine API
- `GET /health` -> {"status":"ok"}
- `POST /run-payroll` (multipart: salary, master, template, month) -> payslips_generated[], updated_master, flags
