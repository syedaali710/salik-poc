"""
SALIC AI Insights — chat over the dashboard data, generate charts/tables/
narrative, export a SALIC-branded .pptx report.

Flow: browser chat -> /chat (Claude, grounded in data/dashboard_data.json)
      -> answer + optional chart/table -> "Report" panel -> /export_report
      -> SALIC-template .pptx download.
"""
import os

from fastapi import FastAPI, Body
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

from pptx_builder import build_report_pptx
from chat import answer_question

HERE = os.path.dirname(os.path.abspath(__file__))
app = FastAPI(title="SALIC AI Insights")


@app.get("/", response_class=HTMLResponse)
@app.get("/report", response_class=HTMLResponse)
def report_page():
    with open(os.path.join(HERE, "static", "report.html"), encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers={"Cache-Control": "no-store"})


@app.post("/chat")
async def chat_endpoint(payload: dict = Body(...)):
    """Question (+ history) -> AI narrative answer + optional chart/table, grounded
    in data/dashboard_data.json."""
    result = answer_question(payload.get("question", ""), payload.get("history") or [])
    return JSONResponse(result)


@app.post("/export_report")
async def export_report(payload: dict = Body(...)):
    """Accumulated chat blocks (narrative/chart/table) -> SALIC-template .pptx report."""
    company = payload.get("company") or "SALIC"
    period = payload.get("period") or ""
    pptx_buf = build_report_pptx(payload.get("blocks") or [], company, period)
    fname = f"{company.strip().replace(' ', '_') or 'SALIC'}_AI_Insights_Report.pptx"
    return StreamingResponse(
        pptx_buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
