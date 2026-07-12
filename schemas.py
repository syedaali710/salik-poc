"""Pydantic request/response models for the FastAPI API."""
from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class HistoryTurn(BaseModel):
    role: Literal["user", "assistant"] = "user"
    content: str = ""


class ChatRequest(BaseModel):
    question: str = ""
    history: list[HistoryTurn] = Field(default_factory=list)


class ChartSeries(BaseModel):
    name: str
    values: list[str]


class ChartSpec(BaseModel):
    type: str
    title: str = ""
    insight: str = ""
    unit: str = ""
    categories: list[str] = Field(default_factory=list)
    series: list[ChartSeries] = Field(default_factory=list)


class TableSpec(BaseModel):
    title: str = "Table"
    columns: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class ChatResponse(BaseModel):
    answer: str
    chart: ChartSpec | dict[str, Any] | None = None
    table: TableSpec | dict[str, Any] | None = None
    ok: bool = True


class TranscribeResponse(BaseModel):
    transcript: str


class ReportBlock(BaseModel):
    type: Literal["narrative", "chart", "table"]
    title: str | None = None
    text: str | None = None
    spec: dict[str, Any] | None = None


class ExportReportRequest(BaseModel):
    company: str = "SALIC"
    period: str = ""
    blocks: list[ReportBlock] = Field(default_factory=list)


class HeygenTokenResponse(BaseModel):
    data: Any | None = None
    error: str | None = None
