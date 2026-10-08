"""
Native OOXML chart data updater for DOCX files.

Operates at the ZIP byte level — reads the template DOCX, updates the
chart XML with literal chart data for each chart, and returns modified bytes.

No matplotlib, xlwings, or Excel installation required.
Works on Linux and any platform with Python + lxml.

Template chart mapping (template1.docx):
  chart1.xml → Inflation (WPI, CPI, IIP) — lineChart, catAx
  chart2.xml → Interest Rates (Repo, MCLR-1Y) — lineChart, dateAx, values as decimals
  chart3.xml → Gas Yearly — barChart, catAx (string labels)
  chart4.xml → Gas Monthly — barChart, dateAx
  chart5.xml → Share Price (Market Cap + GAIL Share) — combo bar+line, catAx (string labels)
"""

import io
import logging
import zipfile
from datetime import date, datetime
from typing import Optional

from lxml import etree

logger = logging.getLogger(__name__)

# DrawingML chart namespace
_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_C = f"{{{_C_NS}}}"
_NS = {
    "c": _C_NS,
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


# ─── Low-level XML helpers ─────────────────────────────────────────────────


def _excel_date(year_month: str) -> int:
    """Convert 'YYYY-MM' to Excel date serial number (days since 1899-12-30)."""
    y, m = map(int, year_month.split("-"))
    return (date(y, m, 1) - date(1899, 12, 30)).days


def _month_label(year_month: str) -> str:
    """Convert 'YYYY-MM' to display label like "Jan'24"."""
    dt = datetime.strptime(year_month, "%Y-%m")
    return dt.strftime("%b-%y")


def _safe_float(val, scale: float = 1.0) -> Optional[float]:
    """Safely convert to float with optional scale factor. Returns None on failure."""
    if val is None:
        return None

    try:
        return float(val) * scale
    except (TypeError, ValueError):
        return None


def _rebuild_num_cache(
    cache_el,
    values: list,
    format_code: str = "General",
) -> None:
    """
    Replace all children of a <c:numCache> element with new values.

    This helper is also used before converting the parent c:numRef
    into a c:numLit element.
    """

    for child in list(cache_el):
        cache_el.remove(child)

    fc = etree.SubElement(
        cache_el,
        f"{_C}formatCode",
    )
    fc.text = format_code

    pt_count = etree.SubElement(
        cache_el,
        f"{_C}ptCount",
    )
    pt_count.set(
        "val",
        str(len(values)),
    )

    for idx, v in enumerate(values):

        pt = etree.SubElement(
            cache_el,
            f"{_C}pt",
        )

        pt.set(
            "idx",
            str(idx),
        )

        v_el = etree.SubElement(
            pt,
            f"{_C}v",
        )

        if v is None:
            v_el.text = ""
        else:
            v_el.text = str(
                round(float(v), 6)
            )


def _rebuild_str_cache(
    cache_el,
    values: list,
) -> None:
    """
    Replace all children of a <c:strCache> element with new string values.
    """

    for child in list(cache_el):
        cache_el.remove(child)

    pt_count = etree.SubElement(
        cache_el,
        f"{_C}ptCount",
    )

    pt_count.set(
        "val",
        str(len(values)),
    )

    for idx, v in enumerate(values):

        pt = etree.SubElement(
            cache_el,
            f"{_C}pt",
        )

        pt.set(
            "idx",
            str(idx),
        )

        v_el = etree.SubElement(
            pt,
            f"{_C}v",
        )

        v_el.text = (
            str(v)
            if v is not None
            else ""
        )


def _convert_num_ref_to_literal(
    num_ref_el,
) -> None:
    """
    Convert:

        <c:numRef>
            <c:f>Dashbard!$C$8:$AC$8</c:f>
            <c:numCache>...</c:numCache>
        </c:numRef>

    into:

        <c:numLit>
            <c:formatCode>...</c:formatCode>
            <c:ptCount .../>
            <c:pt .../>
        </c:numLit>

    The worksheet formula is completely removed.

    This is important because the generated DOCX does not contain
    the workbook referenced by Dashbard!$C$... .
    """
    # print converter

    cache = num_ref_el.find(
        "c:numCache",
        _NS,
    )

    if cache is None:
        return

    # Remove worksheet formula if present.
    formula = num_ref_el.find(
        "c:f",
        _NS,
    )

    if formula is not None:
        num_ref_el.remove(formula)

    # Remove cache from the reference element.
    num_ref_el.remove(cache)

    # Change:
    # c:numRef -> c:numLit
    num_ref_el.tag = (
        f"{_C}numLit"
    )

    # Move cache children directly into numLit.
    for child in list(cache):
        num_ref_el.append(child)


def _convert_str_ref_to_literal(
    str_ref_el,
) -> None:
    """
    Convert:

        <c:strRef>
            <c:f>...</c:f>
            <c:strCache>...</c:strCache>
        </c:strRef>

    into:

        <c:strLit>
            <c:ptCount .../>
            <c:pt .../>
        </c:strLit>

    Used for string category axes.
    """
    # print converter

    cache = str_ref_el.find(
        "c:strCache",
        _NS,
    )

    if cache is None:
        return

    # Remove worksheet formula if present.
    formula = str_ref_el.find(
        "c:f",
        _NS,
    )

    if formula is not None:
        str_ref_el.remove(formula)

    # Remove cache from reference.
    str_ref_el.remove(cache)

    # Change:
    # c:strRef -> c:strLit
    str_ref_el.tag = (
        f"{_C}strLit"
    )

    # Move cache children directly into strLit.
    for child in list(cache):
        str_ref_el.append(child)


def _replace_series_name_with_literal(
    tx_el,
    name: str,
) -> None:
    """
    Replace a series title's worksheet reference with a literal c:v.

    Correct OOXML structure:

        <c:tx>
            <c:v>CPI (%)</c:v>
        </c:tx>

    instead of:

        <c:tx>
            <c:strRef>
                <c:f>Dashbard!$B$8</c:f>
                <c:strCache>...</c:strCache>
            </c:strRef>
        </c:tx>
    """

    # Remove existing c:strRef if present.
    str_ref = tx_el.find(
        "c:strRef",
        _NS,
    )

    if str_ref is not None:
        tx_el.remove(str_ref)

    # Remove any existing literal c:v.
    old_v = tx_el.find(
        "c:v",
        _NS,
    )

    if old_v is not None:
        tx_el.remove(old_v)

    # Add literal series name.
    value_el = etree.Element(
        f"{_C}v"
    )

    value_el.text = str(name)

    tx_el.append(value_el)


# ─── Single-chart updater ───────────────────────────────────────────────────


def _update_chart_xml(
    chart_bytes: bytes,
    series_configs: list,
    axis_min: Optional[float] = None,
    axis_max: Optional[float] = None,
    major_unit: Optional[float] = None,
    plot_area_h: Optional[float] = None,
    category_format_code: Optional[str] = None,
) -> bytes:
    """
    Update a single chart XML with new series data.

    Each entry in series_configs corresponds to one <c:ser> in document order:

        {
            "categories": list,
            "cat_type": "num" | "str",
            "values": list[float],
            "format_code": str,
            "name": str,
        }

    Chart-level options:
        category_format_code: Excel format for date axes, e.g. "mmm-yy".

    Numeric categories are converted from c:numRef to c:numLit.

    String categories are converted from c:strRef to c:strLit.

    Numeric values are converted from c:numRef to c:numLit.

    Series names are converted from c:strRef to c:v.

    This intentionally removes worksheet formulas such as:

        Dashbard!$C$7:$AC$7

    because the DOCX does not contain the external workbook.
    """

    root = etree.fromstring(
        chart_bytes
    )

    # Find every series in document order.
    series_list = root.findall(
        ".//c:ser",
        _NS,
    )

    logger.debug(
        "Chart contains %d series; received %d configs",
        len(series_list),
        len(series_configs),
    )

    for ser_index, (ser_el, cfg) in enumerate(
        zip(
            series_list,
            series_configs,
        )
    ):

        # Rotate GAIL share-price data labels vertically.
        if cfg.get("vertical_data_labels"):
            dLbls = ser_el.find(
                "c:dLbls",
                _NS,
            )

            if dLbls is not None:
                txPr = dLbls.find(
                    "c:txPr",
                    _NS,
                )

                if txPr is not None:
                    bodyPr = txPr.find(
                        "a:bodyPr",
                        _NS,
                    )

                    if bodyPr is not None:
                        bodyPr.set(
                            "rot",
                            "5400000",
                        )
                        bodyPr.set(
                            "vert",
                            "horz",
                        )

        categories = cfg.get(
            "categories",
            [],
        )

        values = cfg.get(
            "values",
            [],
        )

        cat_type = cfg.get(
            "cat_type",
            "num",
        )

        fmt = cfg.get(
            "format_code",
            "General",
        )

        # ----------------------------------------------------
        # Update series name
        # ----------------------------------------------------

        name = cfg.get(
            "name"
        )

        if name is not None:

            tx_el = ser_el.find(
                "c:tx",
                _NS,
            )

            if tx_el is not None:

                _replace_series_name_with_literal(
                    tx_el,
                    name,
                )

        # ----------------------------------------------------
        # Update category data
        # ----------------------------------------------------

        cat_el = ser_el.find(
            "c:cat",
            _NS,
        )

        if cat_el is not None:

            # ====================================================
            # STRING CATEGORIES
            # ====================================================
            #
            # Inflation chart should use:
            #
            #   "Apr-24"
            #   "May-24"
            #   "Jun-24"
            #
            # instead of Excel serials:
            #
            #   45415
            #   45445
            #   45476
            #
            # ====================================================

            if cat_type == "str":

                # First try the normal string reference.
                str_ref = cat_el.find(
                    "c:strRef",
                    _NS,
                )

                if str_ref is not None:

                    cache = str_ref.find(
                        "c:strCache",
                        _NS,
                    )

                    if cache is not None:

                        _rebuild_str_cache(
                            cache,
                            categories,
                        )

                        _convert_str_ref_to_literal(
                            str_ref
                        )

                else:

                    # ------------------------------------------------
                    # IMPORTANT:
                    #
                    # The inflation template may contain numRef/date
                    # categories instead of strRef.
                    #
                    # Convert that existing numeric category reference
                    # into a literal string category.
                    # ------------------------------------------------

                    num_ref = cat_el.find(
                        "c:numRef",
                        _NS,
                    )

                    if num_ref is not None:

                        # Remove the existing numeric reference.
                        cat_el.remove(num_ref)

                    # Remove any existing string reference too.
                    old_str_ref = cat_el.find(
                        "c:strRef",
                        _NS,
                    )

                    if old_str_ref is not None:
                        cat_el.remove(old_str_ref)

                    # Create literal string categories.
                    str_lit = etree.SubElement(
                        cat_el,
                        f"{_C}strLit",
                    )

                    _rebuild_str_cache(
                        str_lit,
                        categories,
                    )

            # ====================================================
            # NUMERIC CATEGORIES
            # ====================================================

            else:

                num_ref = cat_el.find(
                    "c:numRef",
                    _NS,
                )

                if num_ref is not None:

                    cache = num_ref.find(
                        "c:numCache",
                        _NS,
                    )

                    if cache is not None:

                        _rebuild_num_cache(
                            cache,
                            categories,
                            format_code="mmm-yy",
                        )

                        _convert_num_ref_to_literal(
                            num_ref
                        )

                else:

                    # Chart may already contain literal numeric
                    # date categories from a previous generation.
                    num_lit = cat_el.find(
                        "c:numLit",
                        _NS,
                    )

                    if num_lit is not None:

                        _rebuild_num_cache(
                            num_lit,
                            categories,
                            format_code="mmm-yy",
                        )

        # ----------------------------------------------------
        # Update value data
        # ----------------------------------------------------

        val_el = ser_el.find(
            "c:val",
            _NS,
        )

        if val_el is not None:

            num_ref = val_el.find(
                "c:numRef",
                _NS,
            )

            if num_ref is not None:

                cache = num_ref.find(
                    "c:numCache",
                    _NS,
                )

                if cache is not None:

                    _rebuild_num_cache(
                        cache,
                        values,
                        format_code=fmt,
                    )

                    # Convert numRef -> numLit
                    _convert_num_ref_to_literal(
                        num_ref
                    )

            else:

                # Chart may already contain literal numeric values.
                num_lit = val_el.find(
                    "c:numLit",
                    _NS,
                )

                if num_lit is not None:

                    _rebuild_num_cache(
                        num_lit,
                        values,
                        format_code=fmt,
                    )


        logger.debug(
            "Updated chart series %d: categories=%d values=%d",
            ser_index,
            len(categories),
            len(values),
        )

    # --------------------------------------------------------
    # Constrain plot area height if requested
    # --------------------------------------------------------

    if plot_area_h is not None:

        plot_area = root.find(
            ".//c:plotArea",
            _NS,
        )

        if plot_area is not None:

            layout_el = plot_area.find(
                "c:layout",
                _NS,
            )

            if layout_el is None:

                layout_el = etree.SubElement(
                    plot_area,
                    f"{_C}layout",
                )

            for child in list(layout_el):
                layout_el.remove(child)

            man = etree.SubElement(
                layout_el,
                f"{_C}manualLayout",
            )

            etree.SubElement(
                man,
                f"{_C}xMode",
            ).set(
                "val",
                "factor",
            )

            etree.SubElement(
                man,
                f"{_C}yMode",
            ).set(
                "val",
                "factor",
            )

            etree.SubElement(
                man,
                f"{_C}x",
            ).set(
                "val",
                "0",
            )

            etree.SubElement(
                man,
                f"{_C}y",
            ).set(
                "val",
                "0",
            )

            etree.SubElement(
                man,
                f"{_C}w",
            ).set(
                "val",
                "1",
            )

            etree.SubElement(
                man,
                f"{_C}h",
            ).set(
                "val",
                str(
                    round(
                        float(plot_area_h),
                        4,
                    )
                ),
            )

    # --------------------------------------------------------
    # Convert string-category charts from date axis to
    # category axis.
    #
    # The IIP & Inflation chart now uses labels such as:
    #   Apr-24
    #   May-24
    #   Jun-24
    #
    # Therefore it must use a category axis instead of a
    # numeric/date axis.
    # --------------------------------------------------------

    if any(
        cfg.get("cat_type") == "str"
        for cfg in series_configs
    ):

        for axis in root.findall(
            ".//c:dateAx",
            _NS,
        ):

            # Change the XML element from dateAx to catAx.
            axis.tag = f"{_C}catAx"

            # Remove date-axis-specific elements.
            for child_name in (
                "baseTimeUnit",
                "majorUnit",
                "majorTimeUnit",
                "minorUnit",
                "minorTimeUnit",
            ):

                child = axis.find(
                    f"c:{child_name}",
                    _NS,
                )

                if child is not None:
                    axis.remove(child)    

    # --------------------------------------------------------
    # Explicit date-axis formatting
    # --------------------------------------------------------
    #
    # Numeric categories are Excel date serials. The chart may
    # otherwise display them as raw numbers such as 45717.
    # Explicitly format date axes using the requested Excel
    # date format, e.g. "mmm-yy" -> Apr-24, May-24, etc.
    #
    if category_format_code:
        # Word/Office can render the category axis as either a date
        # axis or a category axis depending on the original template.
        # Format both possibilities and explicitly keep their labels
        # visible.
        for axis_tag in ("dateAx", "catAx"):
            for axis in root.findall(
                f".//c:{axis_tag}",
                _NS,
            ):
                num_fmt = axis.find(
                    "c:numFmt",
                    _NS,
                )

                if num_fmt is None:
                    num_fmt = etree.SubElement(
                        axis,
                        f"{_C}numFmt",
                    )

                num_fmt.set(
                    "formatCode",
                    category_format_code,
                )
                num_fmt.set(
                    "sourceLinked",
                    "0",
                )

                delete_el = axis.find(
                    "c:delete",
                    _NS,
                )
                if delete_el is None:
                    delete_el = etree.SubElement(
                        axis,
                        f"{_C}delete",
                    )
                delete_el.set("val", "0")

                tick_pos = axis.find(
                    "c:tickLblPos",
                    _NS,
                )
                if tick_pos is None:
                    tick_pos = etree.SubElement(
                        axis,
                        f"{_C}tickLblPos",
                    )
                tick_pos.set("val", "low")

    # --------------------------------------------------------
    # Force horizontal/category axis to cross at the minimum
    # --------------------------------------------------------
    # The inflation chart contains negative WPI/IIP values.  The
    # template can otherwise place the horizontal axis at zero,
    # which makes the negative portion of the series extend into
    # the chart data table.  Crossing at "min" puts the horizontal
    # axis on the bottom boundary (-2 for the inflation chart).
    #
    # Apply this only to category/date axes so other chart behavior
    # is not changed.
    # --------------------------------------------------------
    for axis_tag in ("dateAx", "catAx"):
        for axis in root.findall(
            f".//c:{axis_tag}",
            _NS,
        ):
            crosses = axis.find(
                "c:crosses",
                _NS,
            )

            if crosses is None:
                crosses = etree.SubElement(
                    axis,
                    f"{_C}crosses",
                )

            crosses.set("val", "min")

    # --------------------------------------------------------
    # Apply explicit axis bounds
    # --------------------------------------------------------

    if (
        axis_min is not None
        or axis_max is not None
    ):
        val_axes = root.findall(
            ".//c:valAx",
            _NS,
        )

        if val_axes:
            # Primary/left Y-axis only.
            val_ax = val_axes[0]

            scaling = val_ax.find(
                "c:scaling",
                _NS,
            )

            if scaling is not None:

                if axis_min is not None:
                    min_el = scaling.find(
                        "c:min",
                        _NS,
                    )

                    if min_el is None:
                        min_el = etree.SubElement(
                            scaling,
                            f"{_C}min",
                        )

                    min_el.set(
                        "val",
                        str(float(axis_min)),
                    )

                if axis_max is not None:
                    max_el = scaling.find(
                        "c:max",
                        _NS,
                    )

                    if max_el is None:
                        max_el = etree.SubElement(
                            scaling,
                            f"{_C}max",
                        )

                    max_el.set(
                        "val",
                        str(float(axis_max)),
                    )

            # Apply major tick interval to the primary/left Y-axis.
            if major_unit is not None:
                major_el = val_ax.find(
                    "c:majorUnit",
                    _NS,
                )

                if major_el is None:
                    major_el = etree.SubElement(
                        val_ax,
                        f"{_C}majorUnit",
                    )

                major_el.set(
                    "val",
                    str(float(major_unit)),
                )
    # --------------------------------------------------------
    # DEBUG: Verify chart references before serialization
    # --------------------------------------------------------

    logger.info(
        "CHART XML BEFORE SERIALIZE: numRef=%d numLit=%d strRef=%d strLit=%d",
        len(root.findall(".//c:numRef", _NS)),
        len(root.findall(".//c:numLit", _NS)),
        len(root.findall(".//c:strRef", _NS)),
        len(root.findall(".//c:strLit", _NS)),
    )

    logger.info(
        "CHART XML FORMULAS BEFORE SERIALIZE: %s",
        [
            el.text
            for el in root.findall(".//c:f", _NS)
            if el.text
        ],
    )

    # --------------------------------------------------------
    # Remove externalData
    # --------------------------------------------------------

    ext_data = root.find(
        "c:externalData",
        _NS,
    )

    if ext_data is not None:

        root.remove(
            ext_data
        )

        logger.debug(
            "Removed externalData link from chart XML"
        )

    # --------------------------------------------------------
    # Serialize XML
    # --------------------------------------------------------

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


# ─── DOCX ZIP-level updater ────────────────────────────────────────────────


def update_docx_charts(
    docx_bytes: bytes,
    chart_updates: list,
) -> bytes:
    """
    Apply chart data updates to a DOCX file.

    chart_updates is a list of dicts:

        {
            "chart_file":
                "word/charts/chart1.xml",

            "series": [
                {
                    "name": "CPI (%)",
                    "categories": [...],
                    "cat_type": "num",
                    "values": [...],
                    "format_code": "General",
                }
            ],
            "category_format_code": "mmm-yy",
            "axis_min": -2,
            "axis_max": 10,
        }

    Returns modified DOCX bytes.
    """

    chart_map = {
        u["chart_file"]: u
        for u in chart_updates
    }

    # Namespaces / types for stripping external data links.
    _OLE_TYPE = (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships/"
        "oleObject"
    )

    _RELS_NS = (
        "http://schemas.openxmlformats.org/"
        "package/2006/relationships"
    )

    _C_NS_LOCAL = (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/chart"
    )

    in_buf = io.BytesIO(
        docx_bytes
    )

    out_buf = io.BytesIO()

    with (
        zipfile.ZipFile(
            in_buf,
            "r",
        ) as zin,
        zipfile.ZipFile(
            out_buf,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as zout
    ):

        for item in zin.infolist():

            data = zin.read(
                item.filename
            )

            # ------------------------------------------------
            # Update requested chart
            # ------------------------------------------------

            if item.filename in chart_map:

                try:

                    cfg = chart_map[
                        item.filename
                    ]

                    data = _update_chart_xml(
                        data,
                        cfg["series"],
                        axis_min=cfg.get(
                            "axis_min"
                        ),
                        axis_max=cfg.get(
                            "axis_max"
                        ),
                        major_unit=cfg.get(
                            "major_unit"
                        ),
                        plot_area_h=cfg.get(
                            "plot_area_h"
                        ),
                        category_format_code=cfg.get(
                            "category_format_code"
                        ),
                    )

                    logger.info(
                        "Updated native chart: %s",
                        item.filename,
                    )

                except Exception as exc:

                    logger.error(
                        "Failed to update %s: %s",
                        item.filename,
                        exc,
                        exc_info=True,
                    )

            # ------------------------------------------------
            # Strip externalData from other chart XMLs
            # ------------------------------------------------

            elif (
                item.filename.startswith(
                    "word/charts/chart"
                )
                and item.filename.endswith(
                    ".xml"
                )
            ):

                try:

                    root = etree.fromstring(
                        data
                    )

                    ext_data = root.find(
                        f"{{{_C_NS_LOCAL}}}externalData"
                    )

                    if ext_data is not None:

                        root.remove(
                            ext_data
                        )

                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )

                        logger.debug(
                            "Stripped externalData from %s",
                            item.filename,
                        )

                except Exception as exc:

                    logger.warning(
                        "Could not strip externalData "
                        "from %s: %s",
                        item.filename,
                        exc,
                    )

            # ------------------------------------------------
            # Strip OLE relationships
            # ------------------------------------------------

            elif (
                item.filename.startswith(
                    "word/charts/_rels/"
                )
                and item.filename.endswith(
                    ".rels"
                )
            ):

                try:

                    root = etree.fromstring(
                        data
                    )

                    removed = 0

                    for rel in root.findall(
                        f"{{{_RELS_NS}}}Relationship"
                    ):

                        if rel.get(
                            "Type"
                        ) == _OLE_TYPE:

                            root.remove(
                                rel
                            )

                            removed += 1

                    if removed:

                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )

                        logger.debug(
                            "Stripped %d OLE relationship(s) "
                            "from %s",
                            removed,
                            item.filename,
                        )

                except Exception as exc:

                    logger.warning(
                        "Could not strip OLE rels "
                        "from %s: %s",
                        item.filename,
                        exc,
                    )

            # ------------------------------------------------
            # Write ZIP entry
            # ------------------------------------------------

            zout.writestr(
                item,
                data,
            )

    return out_buf.getvalue()


# ─── Chart-specific update builders ────────────────────────────────────────


def make_inflation_update(
    data: list,
    last_n: int = 18,
) -> dict:
    """
    Build chart update config for chart1.xml.

    Series order:

        ser[0] = CPI (%)
        ser[1] = WPI (%)
        ser[2] = IIP growth (%)

    Data format:

        [
            {
                "month": "YYYY-MM",
                "WPI": float,
                "CPI": float,
                "IIP": float
            },
            ...
        ]
    """

    if not data:
        raise ValueError(
            "Inflation chart received no data"
        )

    rows = sorted(
        data,
        key=lambda d: d["month"],
    )[-last_n:]

    if not rows:
        raise ValueError(
            "Inflation chart received no rows"
        )

    categories = [
    _month_label(
        r["month"]
    )
    for r in rows
    ]

    logger.info(
        "========== INFLATION CHART DATA =========="
    )

    logger.info(
        "Inflation chart rows: %d",
        len(rows),
    )

    for row in rows:

        logger.info(
            "Inflation chart row: "
            "month=%s CPI=%s WPI=%s IIP=%s",
            row.get("month"),
            row.get("CPI"),
            row.get("WPI"),
            row.get("IIP"),
        )

    return {
        "chart_file":
            "word/charts/chart1.xml",

        "axis_min": -2,
        "axis_max": 10,
        # Do not use manual plot-area height here. Word's native
        # data-table layout is template-controlled; the axis crossing
        # above is the reliable separation mechanism.
        "plot_area_h": None,
        "category_format_code": None,

        "series": [

            {
                "name":
                    "CPI (%)",

                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("CPI")
                    )
                    for r in rows
                ],
            },

            {
                "name":
                    "WPI (%)",

                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("WPI")
                    )
                    for r in rows
                ],
            },

            {
                "name":
                    "IIP Growth (%)",

                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("IIP")
                    )
                    for r in rows
                ],
            },
        ],
    }


def make_interest_rates_update(
    data: list,
    last_n: int = 22,
) -> dict:
    """
    Build chart update config for chart2.xml.

    Interest rates:
        Repo Rate
        MCLR 1Y

    Template stores values as decimals:
        6.5% -> 0.065

    Therefore values are scaled by 0.01.
    """

    rows = sorted(
        data,
        key=lambda d: d["month"],
    )[-last_n:]

    logger.info(
    "INTEREST CHART: received=%d rows, last_n=%d, selected_months=%s",
    len(data),
    last_n,
    [r.get("month") for r in rows],
)

    categories = [
        _excel_date(
            r["month"]
        )
        for r in rows
    ]

    return {
        "chart_file":
            "word/charts/chart2.xml",

        "series": [

            {
                "name":
                    "Repo Rate",

                "categories":
                    categories,

                "cat_type":
                    "num",

                "values": [
                    _safe_float(
                        r.get("Repo"),
                        scale=0.01,
                    )
                    for r in rows
                ],

                "format_code":
                    "0.00%",
            },

            {
                "name":
                    "MCLR 1Y",

                "categories":
                    categories,

                "cat_type":
                    "num",

                "values": [
                    _safe_float(
                        r.get("MCLR1Y"),
                        scale=0.01,
                    )
                    for r in rows
                ],

                "format_code":
                    "0.00%",
            },
        ],
        "category_format_code": "mmm-yy",
    }


def make_gas_yearly_update(
    data: list,
) -> dict:
    """
    Build chart update config for chart3.xml.

    Gas yearly:
        net
        LNG
        total

    Categories are fiscal-year strings.
    """

    rows = sorted(
        data,
        key=lambda d: str(
            d.get(
                "year",
                "",
            )
        ),
    )

    categories = [
        str(
            r.get(
                "year",
                "",
            )
        )
        for r in rows
    ]

    return {
        "chart_file":
            "word/charts/chart3.xml",

        "series": [

            {
                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("net")
                    )
                    for r in rows
                ],
            },

            {
                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("lng")
                    )
                    for r in rows
                ],
            },

            {
                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get("total")
                    )
                    for r in rows
                ],
            },
        ],
    }


def make_gas_monthly_update(
    data: list,
) -> dict:
    """
    Build chart update config for chart4.xml.

    Gas monthly:
        net
        LNG
        total
    """

    rows = sorted(
        data,
        key=lambda d: d["month"],
    )

    categories = [
        _excel_date(
            r["month"]
        )
        for r in rows
    ]

    return {
        "chart_file":
            "word/charts/chart4.xml",

        "series": [

            {
                "categories":
                    categories,

                "cat_type":
                    "num",

                "values": [
                    _safe_float(
                        r.get("net")
                    )
                    for r in rows
                ],
            },

            {
                "categories":
                    categories,

                "cat_type":
                    "num",

                "values": [
                    _safe_float(
                        r.get("lng")
                    )
                    for r in rows
                ],
            },

            {
                "categories":
                    categories,

                "cat_type":
                    "num",

                "values": [
                    _safe_float(
                        r.get("total")
                    )
                    for r in rows
                ],
            },
        ],
    }


def make_share_price_update(
    data: list,
    last_n: int = 12,
    chart_file: str = "word/charts/chart5.xml",
) -> dict:
    """
    Build chart update config for the share price chart.

    Data format:

        [
            {
                "month": "YYYY-MM",
                "Market Cap (Rs Cr.)": float,
                "GAIL Share": float
            },
            ...
        ]

    Uses catAx / string categories rather than dateAx.

    chart_file:
        chart5.xml for template1
        chart3.xml for template2
    """

    rows = sorted(
        data,
        key=lambda d: d["month"],
    )[-last_n:]

    categories = [
        _month_label(
            r["month"]
        )
        for r in rows
    ]

    return {
        "chart_file":
            chart_file,

        "series": [

            {
                "name":
                    "Market Cap (Rs Cr.)",

                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get(
                            "Market Cap (Rs Cr.)"
                        )
                    )
                    for r in rows
                ],
            },

            {
                "name":
                    "GAIL Share Price (Rs)",

                "categories":
                    categories,

                "cat_type":
                    "str",

                "values": [
                    _safe_float(
                        r.get(
                            "GAIL Share"
                        )
                    )
                    for r in rows
                ],
                "vertical_data_labels":
                    True,
            },
        ],
        "axis_min": 0,
        "axis_max": 200000,
        "major_unit": 50000,
    }