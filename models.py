"""Interne states til processen send-breve-via-digitalpost."""

from enum import StrEnum


class ProcessingState(StrEnum):
    READY = "READY"
    DATAFORDELER_LOADED = "DATAFORDELER_LOADED"
    TEMPLATE_LOADED = "TEMPLATE_LOADED"
    DOCX_GENERATED = "DOCX_GENERATED"
    PDF_GENERATED = "PDF_GENERATED"
    DIGITALPOST_QUEUED = "DIGITALPOST_QUEUED"
    NOT_SENT = "NOT_SENT"
    FAILED = "FAILED"
