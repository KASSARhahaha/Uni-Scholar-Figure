"""Feature 8: Research Records — pull Uni-Scholar literature into a PPT table.

Calls GET https://uni-scholar.asia/api/literature/papers with a user-provided
JWT Bearer token and renders the returned papers as a table on a new slide.
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

API_BASE = "https://uni-scholar.asia/api/literature/papers"
TOKEN_CACHE_PATH = Path(
    os.environ.get(
        "UNISFIG_TOKEN_FILE",
        str(Path.home() / ".config" / "unisfigure" / "token.txt"),
    )
)


def read_cached_token() -> str | None:
    """Return cached token, or None if missing/expired-looking."""
    if not TOKEN_CACHE_PATH.exists():
        return None
    tok = TOKEN_CACHE_PATH.read_text(encoding="utf-8").strip()
    return tok or None


def write_cached_token(token: str) -> None:
    TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_PATH.write_text(token, encoding="utf-8")
    try:
        os.chmod(TOKEN_CACHE_PATH, 0o600)
    except OSError:
        pass


def clear_cached_token() -> None:
    if TOKEN_CACHE_PATH.exists():
        TOKEN_CACHE_PATH.unlink()


def fetch_papers(
    token: str,
    *,
    limit: int = 20,
    year_from: int | None = None,
    year_to: int | None = None,
    search: str | None = None,
    sort_by: str = "createdAt",
    sort_order: str = "desc",
    base_url: str = API_BASE,
    timeout: float = 15.0,
) -> list[dict]:
    """Call Uni-Scholar /api/literature/papers, return list of paper dicts."""
    params = [f"limit={limit}", f"sortBy={sort_by}", f"sortOrder={sort_order}"]
    if year_from:
        params.append(f"yearFrom={year_from}")
    if year_to:
        params.append(f"yearTo={year_to}")
    if search:
        params.append(f"search={urllib.parse.quote(search)}")
    url = f"{base_url}?{'&'.join(params)}"

    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "User-Agent": "unisfigure/1.3.0",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise PermissionError(
                "Token rejected (HTTP 401). Re-copy from "
                "https://uni-scholar.asia/app/settings"
            ) from e
        raise RuntimeError(f"HTTP {e.code}: {e.reason}") from e

    payload = json.loads(body)
    if not payload.get("success"):
        raise RuntimeError(f"API returned success=false: {payload}")
    data = payload.get("data") or {}
    return data.get("items") or []


def render_records_table(
    deck_path: Path,
    papers: list[dict],
    *,
    title: str | None = None,
) -> int:
    """Append a slide with a table of papers. Returns rows written."""
    prs = Presentation(str(deck_path)) if deck_path.exists() else Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    if title:
        slide.shapes.title.text = title

    headers = ["Title", "Authors", "Journal", "Year", "DOI", "Catalyst"]
    n_rows = len(papers) + 1
    n_cols = len(headers)
    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols, Inches(0.3), Inches(1.2), Inches(9.4), Inches(0.35 * n_rows)
    )
    tbl = tbl_shape.table

    for c, h in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string("1F4E79")
        for para in cell.text_frame.paragraphs:
            para.alignment = 1  # center
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor.from_string("FFFFFF")

    def _short(s: str | None, n: int) -> str:
        if not s:
            return ""
        s = str(s)
        return s if len(s) <= n else s[: n - 3] + "..."

    for i, p in enumerate(papers):
        authors_raw = p.get("authors") or ""
        if isinstance(authors_raw, str):
            try:
                authors = ", ".join(json.loads(authors_raw))
            except (json.JSONDecodeError, TypeError):
                authors = authors_raw
        else:
            authors = ", ".join(authors_raw)
        row = [
            _short(p.get("title"), 60),
            _short(authors, 30),
            _short(p.get("journal"), 25),
            str(p.get("year") or ""),
            p.get("doi") or "",
            _short(p.get("catalystName"), 25),
        ]
        for c, val in enumerate(row):
            cell = tbl.cell(i + 1, c)
            cell.text = val
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10)

    prs.save(str(deck_path))
    return len(papers)
