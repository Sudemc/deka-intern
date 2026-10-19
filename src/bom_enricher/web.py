"""Local web interface for three-workbook BOM comparison."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path
import tempfile
from threading import Lock
from uuid import uuid4
from zipfile import BadZipFile

import json

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.utils.exceptions import InvalidFileException

from bom_enricher.comparison import RuleBook, compare_workbooks, load_rulebook
from bom_enricher.datasheet_analyzer import (
    analyze_single_datasheet,
    generate_datasheet_excel_report,
)


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
    customer_files: list[UploadFile] = File(default=[]),
    customer_file: UploadFile | None = File(None),
    erp_file: UploadFile = File(...),
    rules_file: UploadFile | None = File(None),
) -> dict[str, object]:
    all_customer_uploads: list[UploadFile] = []
    if customer_file is not None and getattr(customer_file, "filename", None):
        all_customer_uploads.append(customer_file)
    for u in customer_files:
        if getattr(u, "filename", None):
            all_customer_uploads.append(u)

    if not all_customer_uploads:
        raise HTTPException(status_code=400, detail="customer: en az bir müşteri BOM dosyası yükleyin.")

    job_dir = tempfile.TemporaryDirectory(prefix="bom-comparison-")
    customer_paths: list[Path] = []
    paths: dict[str, Path | None] = {}
    try:
        for index, upload in enumerate(all_customer_uploads, start=1):
            filename = Path(upload.filename or "").name
            if Path(filename).suffix.casefold() != ".xlsx":
                raise HTTPException(status_code=400, detail="customer: yalnızca .xlsx dosyası yükleyin.")
            content = await upload.read(_MAX_UPLOAD_BYTES + 1)
            if not content:
                raise HTTPException(status_code=400, detail="customer: seçilen dosya boş.")
            if len(content) > _MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="customer: dosya 20 MB sınırını aşıyor.")
            clean_name = Path(filename).name
            c_path = Path(job_dir.name) / clean_name
            if c_path.exists():
                c_path = Path(job_dir.name) / f"{index}_{clean_name}"
            c_path.write_bytes(content)
            customer_paths.append(c_path)

        other_uploads: dict[str, UploadFile] = {"erp": erp_file}
        if rules_file is not None and getattr(rules_file, "filename", None):
            other_uploads["rules"] = rules_file

        for role, upload in other_uploads.items():
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
            paths["rules"] = None

        output_path = Path(job_dir.name) / "Hata_ve_Kiyaslama_Raporu.xlsx"
        result = compare_workbooks(customer_paths, paths["erp"], paths.get("rules"), output_path)
    except HTTPException:
        job_dir.cleanup()
        raise
    except (BadZipFile, InvalidFileException, OSError, ValueError, RuntimeError) as error:
        job_dir.cleanup()
        raise HTTPException(status_code=422, detail=str(error)) from error
    finally:
        for upload in all_customer_uploads:
            await upload.close()
        await erp_file.close()
        if rules_file is not None:
            await rules_file.close()

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


@app.post("/api/datasheet-analyze", tags=["Datasheet"])
async def analyze_datasheets(
    datasheet_files: list[UploadFile] = File(default=[]),
    file: UploadFile | None = File(None),
    datasheet_file: UploadFile | None = File(None),
    rules_file: UploadFile | None = File(None),
    part_numbers_json: str | None = Form(None),
    material_types_json: str | None = Form(None),
    part_number: str | None = Form(None),
    material_type: str | None = Form(None),
) -> dict[str, object]:
    all_datasheet_uploads: list[UploadFile] = []
    if file is not None and getattr(file, "filename", None):
        all_datasheet_uploads.append(file)
    if datasheet_file is not None and getattr(datasheet_file, "filename", None):
        all_datasheet_uploads.append(datasheet_file)
    for u in datasheet_files:
        if getattr(u, "filename", None):
            all_datasheet_uploads.append(u)

    if not all_datasheet_uploads:
        raise HTTPException(status_code=400, detail="En az bir PDF datasheet dosyası yükleyin.")

    job_dir = tempfile.TemporaryDirectory(prefix="datasheet-analysis-")
    try:
        # Load rules
        rules_path: Path
        if rules_file is not None and getattr(rules_file, "filename", None):
            filename = Path(rules_file.filename or "").name
            if Path(filename).suffix.casefold() != ".xlsx":
                raise HTTPException(status_code=400, detail="Kurallar dosyası yalnızca .xlsx olmalıdır.")
            content = await rules_file.read(_MAX_UPLOAD_BYTES + 1)
            if not content:
                raise HTTPException(status_code=400, detail="Kurallar dosyası boş.")
            rules_path = Path(job_dir.name) / "rules.xlsx"
            rules_path.write_bytes(content)
            rulebook = load_rulebook(rules_path)
        elif DEFAULT_RULES_PATH.exists():
            rules_path = DEFAULT_RULES_PATH
            rulebook = load_rulebook(rules_path)
        else:
            rulebook = RuleBook()

        # Parse part number and material type overrides
        pn_overrides: dict[str, str] = {}
        if part_numbers_json:
            try:
                pn_overrides = json.loads(part_numbers_json)
            except Exception:
                pass

        mat_overrides: dict[str, str] = {}
        if material_types_json:
            try:
                mat_overrides = json.loads(material_types_json)
            except Exception:
                pass

        analyses = []
        for upload in all_datasheet_uploads:
            fname = Path(upload.filename or "").name
            if not fname:
                continue
            if Path(fname).suffix.casefold() != ".pdf":
                raise HTTPException(status_code=400, detail=f"'{fname}': Yalnızca PDF dosyası yükleyebilirsiniz.")
            pdf_bytes = await upload.read(_MAX_UPLOAD_BYTES + 1)
            if not pdf_bytes:
                continue
            pn_override = pn_overrides.get(fname) or (part_number.strip() if part_number and part_number.strip() else None)
            mat_override = mat_overrides.get(fname) or (material_type.strip() if material_type and material_type.strip() else None)
            res = analyze_single_datasheet(
                fname,
                pdf_bytes,
                rulebook,
                part_number_override=pn_override,
                material_type_override=mat_override,
            )
            analyses.append(res)

        if not analyses:
            raise HTTPException(status_code=400, detail="Geçerli bir PDF dosyası işlenemedi.")

        # Generate excel report
        report_path = Path(job_dir.name) / "Datasheet_ERP_Esleme_Raporu.xlsx"
        generate_datasheet_excel_report(analyses, report_path)

        result_id = uuid4().hex
        _store_result(result_id, report_path, job_dir)

        primary = analyses[0].to_dict() if analyses else {}
        return {
            **primary,
            "result_id": result_id,
            "count": len(analyses),
            "analyses": [a.to_dict() for a in analyses],
            "download_url": f"/api/results/{result_id}/workbook",
            "analysis_method": primary.get("analysis_method", "deterministic_text"),
            "material_type": primary.get("material_type", ""),
            "part_number": primary.get("part_number", ""),
        }
    except HTTPException:
        job_dir.cleanup()
        raise
    except Exception as exc:
        job_dir.cleanup()
        raise HTTPException(status_code=422, detail=f"Datasheet analizi hatası: {str(exc)}") from exc
    finally:
        for upload in all_datasheet_uploads:
            await upload.close()
        if rules_file is not None:
            await rules_file.close()


@app.get("/api/results/{result_id}/workbook", tags=["BOM"], include_in_schema=False)
def download_result(result_id: str) -> FileResponse:
    with _jobs_lock:
        job = _jobs.get(result_id)
    if job is None or not job[0].is_file():
        raise HTTPException(status_code=404, detail="Rapor bulunamadı veya indirme süresi doldu.")
    return FileResponse(
        job[0],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=job[0].name,
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
