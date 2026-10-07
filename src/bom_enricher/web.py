"""Local web interface for three-workbook BOM comparison."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path
import tempfile
from threading import Lock
from uuid import uuid4
from zipfile import BadZipFile

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.utils.exceptions import InvalidFileException

from bom_enricher.comparison import compare_workbooks


_MAX_UPLOAD_BYTES = 20 * 1024 * 1024
_RESULT_TTL = timedelta(hours=2)
_STATIC_DIR = Path(__file__).with_name("static")
DEFAULT_RULES_PATH = Path(__file__).resolve().parents[2] / "data" / "rules.xlsx"
_jobs: dict[str, tuple[Path, datetime, tempfile.TemporaryDirectory]] = {}
_jobs_lock = Lock()

app = FastAPI(
    title="DEKA BOM Karşılaştırma",
    description="Müşteri BOM, ERP ve malzeme sözlüğü dosyalarını çevrimdışı karşılaştırma servisi.",
    version="0.2.0",
)

app.mount("/static", StaticFiles(directory=_STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def index() -> HTMLResponse:
    return HTMLResponse((_STATIC_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/api/health", tags=["system"])
def health() -> dict[str, str]:
    return {"status": "ok", "service": "bom-comparison"}


@app.post("/api/compare", tags=["BOM"])
async def compare_uploads(
    customer_file: UploadFile = File(...),
    erp_file: UploadFile = File(...),
    rules_file: UploadFile | None = File(None),
) -> dict[str, object]:
    uploads: dict[str, UploadFile] = {
        "customer": customer_file,
        "erp": erp_file,
    }
    if rules_file is not None and getattr(rules_file, "filename", None):
        uploads["rules"] = rules_file

    job_dir = tempfile.TemporaryDirectory(prefix="bom-comparison-")
    paths: dict[str, Path] = {}
    try:
        for role, upload in uploads.items():
            filename = Path(upload.filename or "").name
            if Path(filename).suffix.casefold() != ".xlsx":
                raise HTTPException(status_code=400, detail=f"{role}: yalnızca .xlsx dosyası yükleyin.")
            content = await upload.read(_MAX_UPLOAD_BYTES + 1)
            if not content:
                raise HTTPException(status_code=400, detail=f"{role}: seçilen dosya boş.")
            if len(content) > _MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail=f"{role}: dosya 20 MB sınırını aşıyor.")
            path = Path(job_dir.name) / f"{role}.xlsx"
            path.write_bytes(content)
            paths[role] = path

        if "rules" not in paths:
            if DEFAULT_RULES_PATH.exists():
                paths["rules"] = DEFAULT_RULES_PATH
            else:
                paths["rules"] = None

        output_path = Path(job_dir.name) / "Hata_ve_Kiyaslama_Raporu.xlsx"
        result = compare_workbooks(paths["customer"], paths["erp"], paths.get("rules"), output_path)
    except HTTPException:
        job_dir.cleanup()
        raise
    except (BadZipFile, InvalidFileException, OSError, ValueError, RuntimeError) as error:
        job_dir.cleanup()
        raise HTTPException(status_code=422, detail=str(error)) from error
    finally:
        for upload in uploads.values():
            await upload.close()

    result_id = uuid4().hex
    _store_result(result_id, output_path, job_dir)
    status_counts: dict[str, int] = {}
    error_count = sum(1 for row in result.rows if row.get("Durum") == "HATA")
    warning_count = sum(1 for row in result.rows if row.get("Durum") == "UYARI")
    clean_count = sum(1 for row in result.rows if row.get("Durum") == "UYGUN")

    for row in result.rows:
        issue_text = row.get("Hata/Uyarı Durumu", "")
        if issue_text:
            for issue in issue_text.split("; "):
                if issue:
                    status_counts[issue] = status_counts.get(issue, 0) + 1
    return {
        "result_id": result_id,
        "row_count": len(result.rows),
        "total_part_count": result.total_part_count,
        "error_count": error_count,
        "warning_count": warning_count,
        "clean_count": clean_count,
        "issue_count": sum(status_counts.values()),
        "status_counts": status_counts,
        "rows": result.rows[:500],
        "preview_limited": len(result.rows) > 500,
        "download_url": f"/api/results/{result_id}/workbook",
    }


@app.get("/api/results/{result_id}/workbook", tags=["BOM"], include_in_schema=False)
def download_result(result_id: str) -> FileResponse:
    with _jobs_lock:
        job = _jobs.get(result_id)
    if job is None or not job[0].is_file():
        raise HTTPException(status_code=404, detail="Rapor bulunamadı veya indirme süresi doldu.")
    return FileResponse(
        job[0],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Hata_ve_Kiyaslama_Raporu.xlsx",
    )


@app.get("/download/master", tags=["BOM"])
def download_master_report() -> FileResponse:
    report_file = Path(__file__).resolve().parents[2] / "reports" / "MASTER_ORTAK_RAPOR.xlsx"
    if not report_file.exists():
        raise HTTPException(status_code=404, detail="Rapor dosyası bulunamadı.")
    return FileResponse(
        report_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="MASTER_ORTAK_RAPOR.xlsx",
    )


def _store_result(
    result_id: str, output_path: Path, directory: tempfile.TemporaryDirectory
) -> None:
    now = datetime.now(timezone.utc)
    with _jobs_lock:
        expired = [job_id for job_id, (_, created, _) in _jobs.items() if now - created > _RESULT_TTL]
        for job_id in expired:
            _, _, owner = _jobs.pop(job_id)
            owner.cleanup()
        _jobs[result_id] = (output_path, now, directory)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
