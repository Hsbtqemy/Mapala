"""Template Builder: build outputs from a template sheet and mappings."""

from __future__ import annotations

from dataclasses import dataclass, field
import numbers
from typing import Any

import pandas as pd

from mapala.config import ConcatSource, LaConcordeError
from mapala.io_excel import load_sheet, load_sheet_raw


class TemplateBuilderError(LaConcordeError):
    """Erreur Template Builder."""


@dataclass
class ZoneHeaderSpec:
    label_rows: list[int] = field(default_factory=list)
    tech_row: int = 1

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> ZoneHeaderSpec:
        return cls(
            label_rows=[int(x) for x in d.get("label_rows", [])],
            tech_row=int(d.get("tech_row", 1)),
        )


@dataclass
class ConcatFieldSpec:
    sources: list[ConcatSource] = field(default_factory=list)
    separator: str = "; "
    deduplicate: bool = False
    skip_empty: bool = True

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> ConcatFieldSpec:
        sources = [ConcatSource.from_dict(s) for s in d.get("sources", [])]
        return cls(
            sources=sources,
            separator=str(d.get("separator", "; ")),
            deduplicate=bool(d.get("deduplicate", False)),
            skip_empty=bool(d.get("skip_empty", True)),
        )


@dataclass
class FieldMappingSpec:
    col_index: int
    target: str | None = None
    mode: str = "simple"  # simple | concat
    source_col: str | None = None
    concat: ConcatFieldSpec | None = None

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> FieldMappingSpec:
        mode = str(d.get("mode", "simple"))
        concat = None
        if mode == "concat":
            concat = ConcatFieldSpec.from_dict(d.get("concat", {}))
        return cls(
            col_index=int(d.get("col_index", 0)),
            target=d.get("target"),
            mode=mode,
            source_col=d.get("source_col"),
            concat=concat,
        )


@dataclass
class ZoneSpec:
    name: str
    row_start: int
    row_end: int | None
    col_start: int
    col_end: int | None
    header: ZoneHeaderSpec = field(default_factory=ZoneHeaderSpec)
    field_mappings: list[FieldMappingSpec] = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> ZoneSpec:
        header = ZoneHeaderSpec.from_dict(d.get("header", {}))
        field_mappings = [FieldMappingSpec.from_dict(m) for m in d.get("field_mappings", [])]
        return cls(
            name=str(d.get("name", "Zone")),
            row_start=int(d.get("row_start", 1)),
            row_end=None if d.get("row_end") in (None, "") else int(d.get("row_end")),
            col_start=int(d.get("col_start", 1)),
            col_end=None if d.get("col_end") in (None, "") else int(d.get("col_end")),
            header=header,
            field_mappings=field_mappings,
        )


@dataclass
class TemplateBuilderConfig:
    template_file: str
    template_sheet: str | None
    source_file: str
    source_sheet: str | None
    source_header_row: int = 1
    zones: list[ZoneSpec] = field(default_factory=list)
    output_sheet_name: str = "Output"

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> TemplateBuilderConfig:
        zones = [ZoneSpec.from_dict(z) for z in d.get("zones", [])]
        return cls(
            template_file=str(d.get("template_file", "")),
            template_sheet=d.get("template_sheet"),
            source_file=str(d.get("source_file", "")),
            source_sheet=d.get("source_sheet"),
            source_header_row=int(d.get("source_header_row", 1)),
            zones=zones,
            output_sheet_name=str(d.get("output_sheet_name", "Output")),
        )


def build_output(config: TemplateBuilderConfig) -> dict[str, pd.DataFrame]:
    """Construit les DataFrames de sortie pour le Template Builder."""
    if not config.template_file or not config.source_file:
        raise TemplateBuilderError("template_file et source_file requis")
    if not config.zones:
        raise TemplateBuilderError("Aucune zone définie")

    df_template = load_sheet_raw(config.template_file, config.template_sheet)
    df_source = load_sheet(
        config.source_file,
        config.source_sheet,
        header_row=config.source_header_row,
    )
    outputs: list[pd.DataFrame] = []
    for zone in config.zones:
        outputs.append(_build_zone_output(zone, df_template, df_source))

    # single sheet: stack zones
    stacked = pd.concat(outputs, ignore_index=True)
    return {config.output_sheet_name: stacked}


def _build_zone_output(zone: ZoneSpec, df_template: pd.DataFrame, df_source: pd.DataFrame) -> pd.DataFrame:
    _validate_zone(zone, df_template)
    row_start = zone.row_start - 1
    row_end = zone.row_end - 1 if zone.row_end is not None else len(df_template) - 1
    col_start = zone.col_start - 1
    col_end = zone.col_end - 1 if zone.col_end is not None else df_template.shape[1] - 1

    zone_raw = df_template.iloc[row_start : row_end + 1, col_start : col_end + 1].copy()
    zone_width = zone_raw.shape[1]

    header_end_row = _infer_header_end_row(zone)
    header_rows_count = max(0, header_end_row - zone.row_start + 1)

    header_block = zone_raw.iloc[:header_rows_count, :]

    tech_row_idx = zone.header.tech_row - zone.row_start
    labels: list[str] = []
    if 0 <= tech_row_idx < len(zone_raw):
        labels = [
            "" if pd.isna(v) else str(v)
            for v in zone_raw.iloc[tech_row_idx, :].tolist()
        ]
    mappings: dict[int, FieldMappingSpec] = {}
    for m in zone.field_mappings:
        col_idx = _resolve_mapping_col_index(m, labels, zone_width)
        if col_idx is None:
            continue
        mappings[col_idx] = m

    data_rows: list[list[object]] = []
    for _, row in df_source.iterrows():
        data_rows.append(_build_data_row_row(row, mappings, zone_width))

    out_rows: list[list[object]] = []
    for _, row in header_block.iterrows():
        out_rows.append(_normalize_row(list(row), zone_width))
    out_rows.extend(data_rows)
    return pd.DataFrame(out_rows)


def _validate_zone(zone: ZoneSpec, df_template: pd.DataFrame) -> None:
    if zone.row_start < 1 or zone.col_start < 1:
        raise TemplateBuilderError("row_start et col_start doivent être >= 1")
    if zone.row_end is not None and zone.row_end < zone.row_start:
        raise TemplateBuilderError("row_end doit être >= row_start")
    if zone.col_end is not None and zone.col_end < zone.col_start:
        raise TemplateBuilderError("col_end doit être >= col_start")
    if zone.row_start > len(df_template):
        raise TemplateBuilderError("row_start hors limites du template")
    if zone.col_start > df_template.shape[1]:
        raise TemplateBuilderError("col_start hors limites du template")


def _infer_header_end_row(zone: ZoneSpec) -> int:
    rows = list(zone.header.label_rows)
    rows.append(zone.header.tech_row)
    if not rows:
        return zone.row_start
    return max(rows)


def _build_data_row_row(
    row: pd.Series,
    mappings: dict[int, FieldMappingSpec],
    width: int,
) -> list[object]:
    row_out = ["" for _ in range(width)]
    for col_idx, mapping in mappings.items():
        if col_idx < 0 or col_idx >= width:
            continue
        if mapping.mode == "concat" and mapping.concat:
            text = _concat_from_row(row, mapping.concat)
        else:
            text = _safe_value(row, mapping.source_col)
        row_out[col_idx] = text
    return row_out


def _concat_from_row(row: pd.Series, spec: ConcatFieldSpec) -> str:
    parts = _concat_parts_from_row(row, spec)
    sep = _normalize_separator(spec.separator)
    return sep.join(parts) if parts else ""


def _normalize_separator(value: str | None) -> str:
    if value is None:
        return "; "
    sep = str(value)
    sep = sep.replace("\\r\\n", "\r\n")
    sep = sep.replace("\\n", "\n")
    sep = sep.replace("\\t", "\t")
    return sep


def _concat_parts_from_row(row: pd.Series, spec: ConcatFieldSpec) -> list[str]:
    parts: list[str] = []
    for src in spec.sources:
        text = _safe_value(row, src.col)
        if spec.skip_empty and text.strip() == "":
            continue
        parts.append(f"{src.prefix}{text}" if src.prefix else text)
    return parts


def _safe_value(row: pd.Series, col: str | None) -> str:
    if not col or col not in row.index:
        return ""
    val = row[col]
    return _format_value(val)


def _format_value(val: object) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, numbers.Integral):
        return str(int(val))
    if isinstance(val, numbers.Real):
        try:
            if float(val).is_integer():
                return str(int(val))
        except (TypeError, ValueError, OverflowError):
            pass
        return str(val)
    return str(val)


def _normalize_row(row: list[object], width: int) -> list[object]:
    if len(row) < width:
        row = row + [""] * (width - len(row))
    elif len(row) > width:
        row = row[:width]
    return ["" if pd.isna(v) else v for v in row]


def _resolve_mapping_col_index(
    mapping: FieldMappingSpec,
    labels: list[str],
    width: int,
) -> int | None:
    if mapping.target:
        for idx, label in enumerate(labels):
            if label == str(mapping.target):
                return idx
    if 0 <= mapping.col_index < width:
        return mapping.col_index
    return None
