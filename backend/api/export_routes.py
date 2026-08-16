from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
from sqlalchemy.orm import Session

from backend.database.session import get_db
from backend.services.export_service import (
    export_candidates_excel,
    export_dashboard_pdf,
)


router = APIRouter(
    prefix="/export",
    tags=["Exports"]
)


# ============================================================
# EXCEL EXPORT
# ============================================================

@router.get("/trials/{trial_id}/candidates.xlsx")
def export_candidates_excel_route(
    trial_id: str,
    db: Session = Depends(get_db)
):
    try:
        excel_bytes = export_candidates_excel(
            db,
            trial_id
        )

        return Response(
            content=excel_bytes,
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": (
                    f'attachment; '
                    f'filename="candidates_{trial_id}.xlsx"'
                )
            },
        )

    except Exception as e:
        print("EXCEL EXPORT ERROR:", repr(e))
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

# ============================================================
# PDF EXPORT
# ============================================================

@router.get("/trials/{trial_id}/report.pdf")
def export_dashboard_pdf_route(
    trial_id: str,
    db: Session = Depends(get_db)
):
    try:
        pdf_bytes = export_dashboard_pdf(
            db,
            trial_id
        )

        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": (
                    f'attachment; '
                    f'filename="report_{trial_id}.pdf"'
                )
            },
        )

    except Exception:
        raise HTTPException(
            status_code=500,
            detail="Failed to generate PDF export."
        )