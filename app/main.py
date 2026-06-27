"""
B Hub Payroll Engine — HTTP wrapper around the payroll agent.

n8n downloads the 3 files from Drive, POSTs them here, and gets back:
  - the generated payslip PDFs (base64) with their safe_to_upload + target folder,
  - the updated Employee Master (to write back to Drive),
  - the run flags (ONBOARD / FOLDER / STOP / CHECK).

The engine never touches Drive itself — n8n owns all Drive I/O and the upload safety.
"""
import base64, json, os, shutil, tempfile
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse
import run_payroll_agent as agent

app = FastAPI(title="B Hub Payroll Engine")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/run-payroll")
async def run_payroll(
    salary: UploadFile = File(...),
    master: UploadFile = File(...),
    template: UploadFile = File(...),
    month: str = Form(None),
):
    work = tempfile.mkdtemp()
    out = os.path.join(work, "out")
    os.makedirs(out, exist_ok=True)
    sal = os.path.join(work, salary.filename or "SALARY.xlsx")
    mem = os.path.join(work, "employee_master.json")
    tpl = os.path.join(work, "payslip_template.xlsx")
    for up, path in [(salary, sal), (master, mem), (template, tpl)]:
        with open(path, "wb") as f:
            f.write(await up.read())

    try:
        result = agent.run(drive_dir=None, template=tpl, memory_path=mem,
                           out_dir=out, salary=sal, month=(month or None))
        for g in result.get("payslips_generated", []):
            pdf = os.path.join(out, g["pdf"])
            if os.path.exists(pdf):
                with open(pdf, "rb") as fh:
                    g["pdf_base64"] = base64.b64encode(fh.read()).decode()
        with open(mem) as fh:
            result["updated_master"] = json.load(fh)
        return JSONResponse(result)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
    finally:
        shutil.rmtree(work, ignore_errors=True)
