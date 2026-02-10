"""I/O tableurs : chargement et sauvegarde (Excel, ODS, CSV)."""

from __future__ import annotations

from pathlib import Path
import csv

import pandas as pd

from mapala.config import LaConcordeError


# Formats supportés
SUPPORTED_INPUT_EXTENSIONS = (".xlsx", ".xls", ".ods", ".csv")
SUPPORTED_INPUT_FILTER = "Tableurs (*.xlsx *.xls *.ods *.csv);;Excel (*.xlsx *.xls);;ODS (*.ods);;CSV (*.csv);;Tous (*.*)"
SUPPORTED_OUTPUT_FILTER = "Excel (*.xlsx);;ODS (*.ods);;CSV (*.csv)"


class ExcelFileError(LaConcordeError):
    """Erreur de chargement d'un fichier (fichier absent, feuille inexistante)."""


def _get_engine(path: Path) -> str | None:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    if suffix in (".ods", ".odt"):
        return "odf"
    if suffix == ".xlsb":
        return "pyxlsb"
    return None


def _is_csv(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def _detect_csv_delimiter(path: Path, encoding: str, *, skip_rows: int = 0) -> str | None:
    try:
        with path.open("r", encoding=encoding, errors="replace") as f:
            for _ in range(skip_rows):
                if f.readline() == "":
                    return None
            sample_lines: list[str] = []
            for line in f:
                if line.strip() == "":
                    continue
                sample_lines.append(line)
                if len(sample_lines) >= 5:
                    break
        if not sample_lines:
            return None
        sample = "".join(sample_lines)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
            return dialect.delimiter
        except Exception:
            first = sample_lines[0]
            counts = {d: first.count(d) for d in [",", ";", "\t", "|"]}
            best = max(counts, key=counts.get)
            return best if counts[best] > 0 else None
    except Exception:
        return None


def list_sheets(filepath: str | Path) -> list[str]:
    path = Path(filepath)
    if not path.exists():
        raise ExcelFileError(f"Fichier introuvable: {path}")
    if _is_csv(path):
        return ["(données)"]
    try:
        engine = _get_engine(path)
        xl = pd.ExcelFile(path, engine=engine) if engine else pd.ExcelFile(path)
        return list(xl.sheet_names)  # type: ignore[return-value]
    except ImportError as e:
        ext = path.suffix.lower()
        if ext == ".xls":
            raise ExcelFileError(f"Format .xls requis: pip install xlrd. Détail: {e}") from e
        if ext in (".ods", ".odt"):
            raise ExcelFileError(f"Format ODS requis: pip install odfpy. Détail: {e}") from e
        raise ExcelFileError(f"Impossible de lire {path}: {e}") from e
    except Exception as e:
        raise ExcelFileError(f"Impossible de lire le fichier {path}: {e}") from e


def load_sheet(
    filepath: str | Path,
    sheet_name: str | None = None,
    *,
    dtype: type | dict[str, type] | None = None,
    header_row: int = 1,
) -> pd.DataFrame:
    path = Path(filepath)
    if not path.exists():
        raise ExcelFileError(f"Fichier introuvable: {path}")
    header = max(0, int(header_row) - 1)
    try:
        if _is_csv(path):
            encoding = "utf-8"
            try:
                path.open("r", encoding=encoding).close()
            except UnicodeDecodeError:
                encoding = "latin-1"
            delimiter = _detect_csv_delimiter(path, encoding, skip_rows=header)
            kwargs = {
                "dtype": dtype,
                "header": header,
                "encoding": encoding,
            }
            if delimiter:
                kwargs["sep"] = delimiter
            try:
                df = pd.read_csv(path, **kwargs)
            except TypeError:
                df = pd.read_csv(path, **kwargs)
            return df
        engine = _get_engine(path)
        if engine:
            return pd.read_excel(path, sheet_name=sheet_name, engine=engine, dtype=dtype, header=header)
        return pd.read_excel(path, sheet_name=sheet_name, dtype=dtype, header=header)
    except ImportError as e:
        ext = path.suffix.lower()
        if ext == ".xls":
            raise ExcelFileError("Format .xls requis: pip install xlrd") from e
        if ext in (".ods", ".odt"):
            raise ExcelFileError("Format ODS requis: pip install odfpy") from e
        raise ExcelFileError(f"Impossible de lire {path}: {e}") from e
    except Exception as e:
        raise ExcelFileError(f"Impossible de lire le fichier {path}: {e}") from e


def load_sheet_raw(
    filepath: str | Path,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    path = Path(filepath)
    if not path.exists():
        raise ExcelFileError(f"Fichier introuvable: {path}")
    try:
        if _is_csv(path):
            encoding = "utf-8"
            try:
                path.open("r", encoding=encoding).close()
            except UnicodeDecodeError:
                encoding = "latin-1"
            delimiter = _detect_csv_delimiter(path, encoding)
            kwargs = {"dtype": str, "header": None, "encoding": encoding}
            if delimiter:
                kwargs["sep"] = delimiter
            df = pd.read_csv(path, **kwargs)
            return df
        engine = _get_engine(path)
        if engine:
            return pd.read_excel(path, sheet_name=sheet_name, engine=engine, dtype=str, header=None)
        return pd.read_excel(path, sheet_name=sheet_name, dtype=str, header=None)
    except ImportError as e:
        ext = path.suffix.lower()
        if ext == ".xls":
            raise ExcelFileError("Format .xls requis: pip install xlrd") from e
        if ext in (".ods", ".odt"):
            raise ExcelFileError("Format ODS requis: pip install odfpy") from e
        raise ExcelFileError(f"Impossible de lire {path}: {e}") from e
    except Exception as e:
        raise ExcelFileError(f"Impossible de lire le fichier {path}: {e}") from e


def save_spreadsheet(
    filepath: str | Path,
    dataframes: dict[str, pd.DataFrame],
    *,
    header: bool = True,
    index: bool = False,
) -> None:
    path = Path(filepath)
    suffix = path.suffix.lower()
    if suffix == ".ods":
        engine = "odf"
    elif suffix == ".xlsx":
        engine = "openpyxl"
    else:
        raise ExcelFileError(f"Format de sortie non supporté: {suffix}")

    with pd.ExcelWriter(path, engine=engine) as writer:
        for sheet_name, df in dataframes.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=index, header=header)


def save_output(
    filepath: str | Path,
    dataframes: dict[str, pd.DataFrame],
    *,
    header: bool = True,
    index: bool = False,
) -> None:
    path = Path(filepath)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        if not dataframes:
            raise ExcelFileError("Aucune donnée à exporter.")
        df = next(iter(dataframes.values()))
        df.to_csv(path, index=index, header=header)
        return
    save_spreadsheet(path, dataframes, header=header, index=index)
