"""
DOCX report evaluator for benchmarking the CMD report generation pipeline.

Compares a pipeline-generated DOCX against a ground truth DOCX on two axes:
  1. Tables   — cell-by-cell comparison with exact / normalised / mismatch buckets
  2. Charts   — OOXML embedded series data compared point-by-point with tolerance

Usage (direct import):
    from generation.evaluator import evaluate
    report = evaluate(generated_path, ground_truth_path, label="sept_2025")
    report.print_report()
    report.to_dict()        # save as JSON
"""

from __future__ import annotations

import json
import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants — friendly names matching service.py TABLE_INDICES / CHART_INDICES
# ---------------------------------------------------------------------------

TABLE_NAMES: Dict[str, Dict[int, str]] = {
    "12": {0: "gdp", 1: "gdp_forecast", 2: "satat", 3: "cgd_targets", 4: "cgd_cumulative"},
    "6":  {0: "gdp", 1: "gdp_forecast", 2: "satat", 3: "retail_lng",  4: "cgd_targets", 5: "cgd_cumulative"},
}

CHART_NAMES: Dict[str, Dict[int, str]] = {
    "12": {0: "inflation", 1: "interest_rates", 2: "gas_yearly", 3: "gas_monthly", 4: "share_price"},
    "6":  {0: "inflation", 1: "interest_rates", 2: "share_price"},
}

# Map generated chart index → ground-truth chart index.
#
# Ground truth CMD Dashboards consistently have exactly 3 charts:
#   gt[0] = Inflation (CPI, WPI, IIP)
#   gt[1] = Interest Rates (Repo, MCLR-1Y)
#   gt[2] = Share Price (Market Cap, GAIL share)
#
# The template1.docx used for generation has 5 charts:
#   gen[0] = Inflation
#   gen[1] = Interest Rates
#   gen[2] = Gas Yearly   ← no GT counterpart
#   gen[3] = Gas Monthly  ← no GT counterpart
#   gen[4] = Share Price
#
# template2 has 6 charts, with Share Price at index 5.
_GEN_TO_GT_CHART_MAP: Dict[str, Dict[int, int]] = {
    "12": {0: 0, 1: 1, 4: 2},   # share_price: gen[4] → gt[2]
    "6":  {0: 0, 1: 1, 5: 2},   # share_price: gen[5] → gt[2]
}

# Map generated table index → ground-truth table index.
#
# Ground truth CMD Dashboards have a Retail LNG table at index 3 that is
# absent from template1.docx, shifting the CGD tables by one position:
#
#   gen[0] = GDP           → gt[0]
#   gen[1] = GDP Forecast  → gt[1]
#   gen[2] = SATAT         → gt[2]
#   (no gen counterpart)   ← gt[3] = Retail LNG stations
#   gen[3] = CGD Targets   → gt[4]
#   gen[4] = CGD Cumulative→ gt[5]
#
# template2.docx already includes a Retail LNG table at index 3, so its
# generated indices align naturally with the GT (no cross-mapping needed).
_GEN_TO_GT_TABLE_MAP: Dict[str, Dict[int, int]] = {
    "12": {0: 0, 1: 1, 2: 2, 3: 4, 4: 5},  # cgd_targets: gen[3]→gt[4], cgd_cum: gen[4]→gt[5]
    "6":  {},  # template2 aligns with GT directly
}

_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_C = f"{{{_CHART_NS}}}"


# ---------------------------------------------------------------------------
# Value normalisation
# ---------------------------------------------------------------------------

def _try_numeric(val: str) -> Optional[float]:
    """
    Attempt to extract a plain numeric value from a cell string.

    Handles:  "4,218,520"  "6.7%"  "₹ 84.50"  " 488.1 "
    Returns a rounded float, or None if non-numeric.
    """
    cleaned = re.sub(r"[,₹$\s]", "", val.strip())
    cleaned = cleaned.rstrip("%")       # strip trailing % (keep sign for negatives)
    try:
        return round(float(cleaned), 4)
    except ValueError:
        return None


def _normalise_text(val: str) -> str:
    """
    Lowercase, treat dashes/underscores/spaces as equivalent separators,
    and strip non-alphanumeric punctuation.

    Handles:  "Q1-FY25" ≈ "Q1 FY25"   "World_Bank" ≈ "World Bank"
    """
    s = val.lower().strip()
    s = re.sub(r"[-_]", " ", s)        # dash/underscore → space
    s = re.sub(r"\s+", " ", s)         # collapse whitespace
    s = re.sub(r"[^\w\s/%.()]", "", s) # remove remaining punctuation
    return s.strip()


def compare_cell(generated: str, ground_truth: str) -> str:
    """
    Compare two cell values.

    Returns one of:
      "both_empty"  — both blank after strip
      "exact"       — identical strings after strip
      "normalised"  — differ in formatting but represent the same value
      "mismatch"    — genuinely different
    """
    g  = generated.strip()
    gt = ground_truth.strip()

    if not g and not gt:
        return "both_empty"
    if g == gt:
        return "exact"

    # Numeric comparison (handles comma-formatting, trailing %)
    gn  = _try_numeric(g)
    gtn = _try_numeric(gt)
    if gn is not None and gtn is not None:
        # Absolute tolerance only: 0.011 handles floating-point rounding at 2 dp
        # (e.g. 6.7% vs 6.70%, or 4,218,520 vs 4218520 → 0 difference).
        # No relative tolerance — 488.1 vs 488.2 is a genuine mismatch.
        if abs(gn - gtn) < 0.011:
            return "normalised"

    # Text normalisation (handles capitalisation, punctuation)
    if _normalise_text(g) == _normalise_text(gt):
        return "normalised"

    return "mismatch"


# ---------------------------------------------------------------------------
# Table extraction  (merged-cell aware)
# ---------------------------------------------------------------------------

def _unique_cells(row) -> List:
    """Yield unique cell objects from a row, skipping merged duplicates."""
    seen: set = set()
    result = []
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen:
            seen.add(tc_id)
            result.append(cell)
    return result


def extract_tables(docx_path: Path) -> List[List[List[str]]]:
    """
    Extract all table data from a DOCX file.

    Returns:
        list[tables]  →  each table is  list[rows]  →  each row is  list[str]
    Merged cells are deduplicated (each cell counted once).
    """
    doc = Document(str(docx_path))
    result = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            cells = [c.text.strip() for c in _unique_cells(row)]
            rows.append(cells)
        result.append(rows)
    return result


# ---------------------------------------------------------------------------
# Chart series extraction  (OOXML / numCache)
# ---------------------------------------------------------------------------

def _get_series_name(ser_el) -> str:
    """Extract the series display name from a <c:ser> element."""
    tx = ser_el.find(f".//{_C}tx")
    if tx is not None:
        v = tx.find(f".//{_C}v")
        if v is not None and v.text:
            return v.text.strip()
    return ""


def _get_series_values(ser_el) -> List[Optional[float]]:
    """
    Extract numeric point values from a <c:ser> element's numCache.

    Missing indices in the ptCount range are returned as None.
    Returns empty list if no numCache is found (e.g. external data link only).
    """
    val_el = ser_el.find(f".//{_C}val")
    if val_el is None:
        return []
    cache = val_el.find(f".//{_C}numCache")
    if cache is None:
        return []

    pt_count_el = cache.find(f"{_C}ptCount")
    n = int(pt_count_el.get("val", "0")) if pt_count_el is not None else 0

    values: List[Optional[float]] = [None] * n
    for pt in cache.findall(f"{_C}pt"):
        idx_attr = pt.get("idx")
        v_el = pt.find(f"{_C}v")
        if idx_attr is not None and v_el is not None and v_el.text:
            idx = int(idx_attr)
            if idx < n:
                try:
                    values[idx] = float(v_el.text)
                except ValueError:
                    pass
    return values


def extract_chart_series(docx_path: Path) -> Dict[int, Dict[str, List[Optional[float]]]]:
    """
    Extract numeric series data from all OOXML charts inside a DOCX.

    Chart index is 0-based, ordered by filename (chart1.xml → 0, chart2.xml → 1, …).

    Returns:
        {chart_idx: {series_name: [value_or_None, …]}}

    A chart is omitted if it has no embedded numCache (external data link only).
    """
    result: Dict[int, Dict[str, List[Optional[float]]]] = {}

    with zipfile.ZipFile(str(docx_path)) as zf:
        chart_files = sorted(
            (n for n in zf.namelist() if re.match(r"word/charts/chart\d+\.xml$", n)),
            key=lambda n: int(re.search(r"chart(\d+)", n).group(1)),
        )

        for chart_idx, chart_file in enumerate(chart_files):
            with zf.open(chart_file) as fh:
                root = ET.parse(fh).getroot()

            series_data: Dict[str, List[Optional[float]]] = {}
            for i, ser_el in enumerate(root.iter(f"{_C}ser")):
                values = _get_series_values(ser_el)
                if not values:
                    continue           # skip series with no embedded data
                name = _get_series_name(ser_el) or f"series_{i}"
                # Handle duplicate series names (append index suffix)
                if name in series_data:
                    name = f"{name}_{i}"
                series_data[name] = values

            if series_data:
                result[chart_idx] = series_data

    return result


# ---------------------------------------------------------------------------
# Result dataclasses
# ---------------------------------------------------------------------------

@dataclass
class CellDiff:
    row:          int
    col:          int
    generated:    str
    ground_truth: str
    match:        str   # "mismatch" (always, since matched cells are not stored here)


@dataclass
class TableEvaluation:
    index:       int
    name:        str
    rows:        int
    cols:        int
    total_cells: int
    exact:       int
    normalised:  int
    mismatch:    int
    both_empty:  int
    diffs:       List[CellDiff] = field(default_factory=list)

    @property
    def data_cells(self) -> int:
        """Non-empty cells in ground truth (the denominator for accuracy)."""
        return self.total_cells - self.both_empty

    @property
    def exact_accuracy(self) -> float:
        return self.exact / self.data_cells if self.data_cells else 1.0

    @property
    def fuzzy_accuracy(self) -> float:
        return (self.exact + self.normalised) / self.data_cells if self.data_cells else 1.0


@dataclass
class SeriesDiff:
    series_name:  str
    index:        int
    generated:    Optional[float]
    ground_truth: Optional[float]


@dataclass
class ChartEvaluation:
    index:        int
    name:         str
    total_points: int
    matches:      int
    mismatches:   int
    tolerance:    float
    diffs:        List[SeriesDiff] = field(default_factory=list)

    @property
    def accuracy(self) -> float:
        return self.matches / self.total_points if self.total_points else 1.0


@dataclass
class EvaluationReport:
    label:              str
    generated_path:     str
    ground_truth_path:  str
    template_type:      str
    table_evaluations:  List[TableEvaluation]
    chart_evaluations:  List[ChartEvaluation]
    evaluated_at:       str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    # ------------------------------------------------------------------
    # Aggregate scores
    # ------------------------------------------------------------------

    @property
    def overall_table_exact(self) -> float:
        total = sum(t.data_cells for t in self.table_evaluations)
        if not total:
            return 1.0
        return sum(t.exact for t in self.table_evaluations) / total

    @property
    def overall_table_fuzzy(self) -> float:
        total = sum(t.data_cells for t in self.table_evaluations)
        if not total:
            return 1.0
        return sum(t.exact + t.normalised for t in self.table_evaluations) / total

    @property
    def overall_chart_accuracy(self) -> float:
        total = sum(c.total_points for c in self.chart_evaluations)
        if not total:
            return 1.0
        return sum(c.matches for c in self.chart_evaluations) / total

    # ------------------------------------------------------------------
    # Serialisation
    # ------------------------------------------------------------------

    def to_dict(self) -> dict:
        return {
            "label":              self.label,
            "generated_path":     self.generated_path,
            "ground_truth_path":  self.ground_truth_path,
            "template_type":      self.template_type,
            "evaluated_at":       self.evaluated_at,
            "summary": {
                "table_exact_accuracy": round(self.overall_table_exact, 4),
                "table_fuzzy_accuracy": round(self.overall_table_fuzzy, 4),
                "chart_accuracy":       round(self.overall_chart_accuracy, 4),
            },
            "tables": [asdict(t) for t in self.table_evaluations],
            "charts": [asdict(c) for c in self.chart_evaluations],
        }

    def save_json(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w") as fh:
            json.dump(self.to_dict(), fh, indent=2, default=str)
        logger.info(f"Evaluation JSON saved: {path}")

    # ------------------------------------------------------------------
    # Human-readable output
    # ------------------------------------------------------------------

    def print_report(self, show_diffs: bool = True, max_diffs: int = 10) -> None:
        _print_report(self, show_diffs=show_diffs, max_diffs=max_diffs)


# ---------------------------------------------------------------------------
# Internal: table comparison
# ---------------------------------------------------------------------------

def _evaluate_tables(
    gen_tables: List[List[List[str]]],
    gt_tables:  List[List[List[str]]],
    template_type:   str  = "12",
    skip_header_row: bool = True,
    last_col_only:   bool = False,
) -> List[TableEvaluation]:
    """
    Compare tables cell-by-cell.

    When ``last_col_only=True``, table index 0 (GDP) is compared using only
    the rightmost column of each document — the most recently extracted quarter.
    All other tables are always compared in full.
    """
    names      = TABLE_NAMES.get(template_type, TABLE_NAMES["12"])
    gen_to_gt  = _GEN_TO_GT_TABLE_MAP.get(template_type, {})

    # Build (gen_idx, gt_idx) pairs to compare.
    # When a cross-index map is available, use it so that e.g. the CGD tables
    # (gen[3], gen[4]) are compared against the correct GT positions (gt[4], gt[5])
    # rather than gt[3] (LNG stations) and gt[4] (wrong CGD table).
    if gen_to_gt:
        pairs = [
            (gen_idx, gt_idx)
            for gen_idx, gt_idx in sorted(gen_to_gt.items())
            if gen_idx < len(gen_tables) and gt_idx < len(gt_tables)
        ]
        logger.info(
            f"Tables: using cross-index map {gen_to_gt}. Comparing {len(pairs)} pairs."
        )
    else:
        n_tables = min(len(gen_tables), len(gt_tables))
        pairs = [(i, i) for i in range(n_tables)]
        if len(gen_tables) != len(gt_tables):
            logger.warning(
                f"Table count mismatch: generated={len(gen_tables)}, gt={len(gt_tables)}. "
                f"Comparing first {n_tables}."
            )

    results: List[TableEvaluation] = []
    for tbl_idx, (gen_idx, gt_idx) in enumerate(pairs):
        gen_rows = gen_tables[gen_idx]
        gt_rows  = gt_tables[gt_idx]
        name     = names.get(gen_idx, f"table_{gen_idx}")

        start_row = 1 if skip_header_row else 0
        n_rows    = min(len(gen_rows), len(gt_rows))

        if last_col_only and tbl_idx == 0:
            # GDP table: compare only the rightmost column of each document
            # (the most recently extracted quarter value).
            gen_last_col = max((len(r) - 1 for r in gen_rows if r), default=0)
            gt_last_col  = max((len(r) - 1 for r in gt_rows  if r), default=0)
            col_pairs       = [(gen_last_col, gt_last_col)]
            reported_n_cols = 1
        else:
            # Full comparison: all columns up to the widest row in either doc.
            n_cols = max(
                (len(r) for r in gen_rows[start_row:n_rows]),
                default=0,
            )
            n_cols = max(n_cols, max(
                (len(r) for r in gt_rows[start_row:n_rows]),
                default=0,
            ))
            col_pairs       = [(c, c) for c in range(n_cols)]
            reported_n_cols = n_cols

        exact = normalised = mismatch = both_empty = 0
        diffs: List[CellDiff] = []

        for row_idx in range(start_row, n_rows):
            gen_row = gen_rows[row_idx] if row_idx < len(gen_rows) else []
            gt_row  = gt_rows[row_idx]  if row_idx < len(gt_rows)  else []

            for (gen_col, gt_col) in col_pairs:
                g  = gen_row[gen_col] if gen_col < len(gen_row) else ""
                gt = gt_row[gt_col]   if gt_col  < len(gt_row)  else ""

                verdict = compare_cell(g, gt)
                if   verdict == "exact":       exact      += 1
                elif verdict == "normalised":  normalised += 1
                elif verdict == "mismatch":
                    mismatch += 1
                    diffs.append(CellDiff(row_idx, gen_col, g, gt, verdict))
                else:                          both_empty += 1

        total = exact + normalised + mismatch + both_empty
        results.append(TableEvaluation(
            index=gen_idx, name=name,
            rows=n_rows - start_row, cols=reported_n_cols,
            total_cells=total, exact=exact, normalised=normalised,
            mismatch=mismatch, both_empty=both_empty,
            diffs=diffs,
        ))

    return results


# ---------------------------------------------------------------------------
# Internal: chart comparison
# ---------------------------------------------------------------------------

def _evaluate_charts(
    gen_charts: Dict[int, Dict[str, List[Optional[float]]]],
    gt_charts:  Dict[int, Dict[str, List[Optional[float]]]],
    template_type:  str   = "12",
    tolerance:      float = 0.01,
    last_n_points:  int   = 0,
) -> List[ChartEvaluation]:
    """
    Compare chart series point-by-point.

    When ``last_n_points > 0``, only the last N values of each series are
    compared (e.g. ``last_n_points=1`` compares only the most recently
    extracted data point per series).  The ``index`` field in ``SeriesDiff``
    reflects the original position in the generated series for readability.
    """
    names       = CHART_NAMES.get(template_type, CHART_NAMES["12"])
    gen_to_gt   = _GEN_TO_GT_CHART_MAP.get(template_type, {})
    results: List[ChartEvaluation] = []

    # Build (gen_idx, gt_idx) pairs.
    # When a cross-index map is available, use it so that e.g. the share_price
    # chart (gen[4]) is compared to the correct GT chart (gt[2]) rather than
    # the GT chart at position 4 (which may not exist or may be a different chart).
    if gen_to_gt:
        pairs = [
            (gen_idx, gt_idx)
            for gen_idx, gt_idx in gen_to_gt.items()
            if gen_idx in gen_charts and gt_idx in gt_charts
        ]
        logger.info(
            f"Charts: using cross-index map {gen_to_gt}. "
            f"Comparing {len(pairs)} pairs."
        )
    else:
        common_indices = sorted(set(gen_charts) & set(gt_charts))
        pairs = [(idx, idx) for idx in common_indices]
        if len(gen_charts) != len(gt_charts):
            logger.warning(
                f"Chart count mismatch: generated={len(gen_charts)}, gt={len(gt_charts)}. "
                f"Comparing {len(pairs)} same-index pairs."
            )

    for gen_idx, gt_idx in pairs:
        name       = names.get(gen_idx, f"chart_{gen_idx}")
        gen_series = gen_charts[gen_idx]
        gt_series  = gt_charts[gt_idx]

        # Match series by name; fall back to positional matching
        common_series = sorted(set(gen_series) & set(gt_series))
        if not common_series:
            gen_vals_list = list(gen_series.values())
            gt_vals_list  = list(gt_series.values())
            common_series = [f"s{i}" for i in range(min(len(gen_vals_list), len(gt_vals_list)))]
            gen_series = {f"s{i}": v for i, v in enumerate(gen_vals_list)}
            gt_series  = {f"s{i}": v for i, v in enumerate(gt_vals_list)}

        matches = mismatches = 0
        diffs: List[SeriesDiff] = []

        for sname in common_series:
            g_vals  = gen_series[sname]
            gt_vals = gt_series[sname]

            if last_n_points > 0:
                # Slice to last N points; remember original offsets for display.
                g_offset  = max(0, len(g_vals)  - last_n_points)
                gt_offset = max(0, len(gt_vals) - last_n_points)
                g_vals    = g_vals[-last_n_points:]
                gt_vals   = gt_vals[-last_n_points:]
            else:
                # Tail-align: when series have different lengths, compare the last
                # min(len_gen, len_gt) points from each (most recent data aligns at end).
                # This handles the common case where generated charts have shorter history
                # than ground-truth charts but carry the same recent data.
                n_common = min(len(g_vals), len(gt_vals))
                g_offset  = max(0, len(g_vals)  - n_common)
                gt_offset = max(0, len(gt_vals) - n_common)
                g_vals    = g_vals[-n_common:]  if n_common > 0 else g_vals
                gt_vals   = gt_vals[-n_common:] if n_common > 0 else gt_vals

            n = max(len(g_vals), len(gt_vals))

            for pt_idx in range(n):
                gv  = g_vals[pt_idx]  if pt_idx < len(g_vals)  else None
                gtv = gt_vals[pt_idx] if pt_idx < len(gt_vals) else None

                # Both None → skip (no data at this position in either)
                if gv is None and gtv is None:
                    continue

                if (gv is not None and gtv is not None
                        and (abs(gv - gtv) <= tolerance
                             or abs(gv - gtv) / (abs(gtv) + 1e-9) <= 0.001)):
                    matches += 1
                else:
                    mismatches += 1
                    diffs.append(SeriesDiff(sname, g_offset + pt_idx, gv, gtv))

        results.append(ChartEvaluation(
            index=gen_idx, name=name,
            total_points=matches + mismatches,
            matches=matches, mismatches=mismatches,
            tolerance=tolerance, diffs=diffs,
        ))

    return results


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def evaluate(
    generated_path:    Path,
    ground_truth_path: Path,
    label:             str   = "",
    template_type:     str   = "12",
    skip_header_row:   bool  = True,
    chart_tolerance:   float = 0.01,
) -> EvaluationReport:
    """
    Compare a pipeline-generated DOCX against a known-good ground truth DOCX.

    Args:
        generated_path:    Path to the generated report DOCX.
        ground_truth_path: Path to the ground truth / reference DOCX.
        label:             Human-readable label for the benchmark (e.g. "sept_2025").
        template_type:     "12" or "6" — controls table/chart name mapping.
        skip_header_row:   Ignore row 0 of each table (column headers).
        chart_tolerance:   Absolute tolerance for numeric chart value matching.

    Returns:
        EvaluationReport with per-table and per-chart breakdowns.
    """
    logger.info(
        f"Evaluating '{label}': {generated_path.name} vs {ground_truth_path.name}"
    )

    gen_tables = extract_tables(generated_path)
    gt_tables  = extract_tables(ground_truth_path)
    logger.info(f"Tables: generated={len(gen_tables)}, gt={len(gt_tables)}")

    gen_charts = extract_chart_series(generated_path)
    gt_charts  = extract_chart_series(ground_truth_path)
    logger.info(
        f"Charts with embedded data: generated={len(gen_charts)}, gt={len(gt_charts)}"
    )

    table_evals = _evaluate_tables(gen_tables, gt_tables, template_type, skip_header_row)
    chart_evals = _evaluate_charts(gen_charts, gt_charts, template_type, chart_tolerance)

    return EvaluationReport(
        label=label,
        generated_path=str(generated_path),
        ground_truth_path=str(ground_truth_path),
        template_type=template_type,
        table_evaluations=table_evals,
        chart_evaluations=chart_evals,
    )


def evaluate_extracted_only(
    generated_path:    Path,
    ground_truth_path: Path,
    label:             str   = "",
    template_type:     str   = "12",
    skip_header_row:   bool  = True,
    chart_tolerance:   float = 0.01,
) -> EvaluationReport:
    """
    Compare only the freshly extracted values for the current data month.

    - Charts: last data point per series only (the value extracted this run).
    - GDP table (index 0): rightmost column only (the most recently extracted quarter).
    - All other tables (GDP Forecast, SATAT, CGD): full cell-by-cell comparison,
      since every cell in those tables comes from fresh extraction.

    Use this instead of ``evaluate()`` when you want a signal on extraction
    quality rather than historical reproduction accuracy.
    """
    logger.info(
        f"Evaluating extracted-only '{label}': "
        f"{generated_path.name} vs {ground_truth_path.name}"
    )

    gen_tables = extract_tables(generated_path)
    gt_tables  = extract_tables(ground_truth_path)
    gen_charts = extract_chart_series(generated_path)
    gt_charts  = extract_chart_series(ground_truth_path)

    table_evals = _evaluate_tables(
        gen_tables, gt_tables, template_type, skip_header_row,
        last_col_only=True,
    )
    chart_evals = _evaluate_charts(
        gen_charts, gt_charts, template_type, chart_tolerance,
        last_n_points=1,
    )

    return EvaluationReport(
        label=f"{label} [extracted-only]",
        generated_path=str(generated_path),
        ground_truth_path=str(ground_truth_path),
        template_type=template_type,
        table_evaluations=table_evals,
        chart_evaluations=chart_evals,
    )


# ---------------------------------------------------------------------------
# Human-readable report printer
# ---------------------------------------------------------------------------

_W = 72


def _print_report(report: EvaluationReport, show_diffs: bool = True, max_diffs: int = 10) -> None:
    print("=" * _W)
    print(f"  BENCHMARK: {report.label.upper()}")
    print(f"  Generated    : {Path(report.generated_path).name}")
    print(f"  Ground truth : {Path(report.ground_truth_path).name}")
    print(f"  Evaluated    : {report.evaluated_at}")
    print("=" * _W)

    # ── Tables ──────────────────────────────────────────────────────────
    print("\nTABLE COMPARISON")
    print("-" * _W)
    for t in report.table_evaluations:
        e_pct = f"{t.exact_accuracy * 100:5.1f}%"
        f_pct = f"{t.fuzzy_accuracy * 100:5.1f}%"
        print(
            f"  [{t.index}] {t.name:<18}  "
            f"exact={e_pct}  fuzzy={f_pct}  "
            f"| {t.exact}✓ {t.normalised}~ {t.mismatch}✗ {t.both_empty}∅"
        )
        if show_diffs and t.diffs:
            for d in t.diffs[:max_diffs]:
                gen_repr = repr(d.generated)[:35]
                gt_repr  = repr(d.ground_truth)[:35]
                print(f"      r{d.row:02d}c{d.col:02d}  GEN {gen_repr:<37} GT {gt_repr}")
            if len(t.diffs) > max_diffs:
                print(f"      … and {len(t.diffs) - max_diffs} more mismatches (see JSON)")

    # ── Charts ───────────────────────────────────────────────────────────
    if report.chart_evaluations:
        print("\nCHART COMPARISON  (OOXML embedded series data)")
        print("-" * _W)
        for c in report.chart_evaluations:
            acc = f"{c.accuracy * 100:5.1f}%"
            print(
                f"  [{c.index}] {c.name:<18}  "
                f"accuracy={acc}  "
                f"| {c.matches}/{c.total_points} pts  tol=±{c.tolerance}"
            )
            if show_diffs and c.diffs:
                for d in c.diffs[:5]:
                    print(
                        f"      series={d.series_name!r:20s}  idx={d.index:2d}  "
                        f"GEN={d.generated}  GT={d.ground_truth}"
                    )
                if len(c.diffs) > 5:
                    print(f"      … and {len(c.diffs) - 5} more")
    else:
        print("\nCHART COMPARISON  — no embedded series data found in one or both files")

    # ── Summary ──────────────────────────────────────────────────────────
    print("\nSUMMARY")
    print("-" * _W)
    print(f"  Table accuracy (exact) : {report.overall_table_exact  * 100:.1f}%")
    print(f"  Table accuracy (fuzzy) : {report.overall_table_fuzzy  * 100:.1f}%")
    if report.chart_evaluations:
        print(f"  Chart accuracy         : {report.overall_chart_accuracy * 100:.1f}%")
    print("=" * _W)
