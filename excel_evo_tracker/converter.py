"""
XLSB → XLSX conversion using Excel COM automation.

Caches conversions by MD5 of the source XLSB so repeated runs skip work.
Windows + Excel required. The rest of the pipeline has no Windows
dependency — this is the only module that does.

Typical usage:

    from excel_evo_tracker.converter import convert_batch
    xlsx_paths = convert_batch(Path("xlsb_input"))
"""

from __future__ import annotations

import hashlib
import json
import logging
import shutil
import tempfile
from pathlib import Path

from . import config

logger = logging.getLogger(__name__)

# Excel SaveAs format code for .xlsx
_XL_OPEN_XML_WORKBOOK = 51


# ── Hashing & manifest ────────────────────────────────────────────────


def compute_file_hash(path: Path, chunk_size: int = 1 << 20) -> str:
    """Return the MD5 hex digest of a file (streaming, memory-safe)."""
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_manifest() -> dict:
    if config.CONVERSION_MANIFEST.exists():
        try:
            return json.loads(config.CONVERSION_MANIFEST.read_text())
        except (json.JSONDecodeError, OSError) as e:
            logger.warning("Could not read conversion manifest (%s). Starting fresh.", e)
    return {}


def _save_manifest(manifest: dict) -> None:
    config.CONVERSION_MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    config.CONVERSION_MANIFEST.write_text(json.dumps(manifest, indent=2))


# ── Core conversion ───────────────────────────────────────────────────


def convert_single(
    xlsb_path: Path,
    output_dir: Path | None = None,
    force: bool = False,
) -> Path:
    """
    Convert a single XLSB to XLSX using Excel COM.

    Returns the path to the resulting XLSX. If the source file's hash is
    already in the manifest and the cached XLSX exists, returns the cached
    path without reopening Excel.

    Args:
        xlsb_path: Path to the source XLSB file.
        output_dir: Directory for the XLSX. Defaults to config.XLSX_CACHE_DIR.
        force: If True, re-convert even if a cached copy exists.

    Raises:
        FileNotFoundError: If xlsb_path doesn't exist.
        RuntimeError: If Excel COM fails or win32com is unavailable.
    """
    xlsb_path = Path(xlsb_path).resolve()
    if not xlsb_path.exists():
        raise FileNotFoundError(f"XLSB not found: {xlsb_path}")

    output_dir = Path(output_dir) if output_dir else config.XLSX_CACHE_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    # Check cache
    file_hash = compute_file_hash(xlsb_path)
    manifest = _load_manifest()

    if not force and file_hash in manifest:
        cached_name = manifest[file_hash]
        cached_path = output_dir / cached_name
        if cached_path.exists():
            logger.info("Cache hit for %s → %s", xlsb_path.name, cached_name)
            return cached_path
        else:
            logger.warning("Manifest references missing file %s; re-converting.", cached_name)

    # Lazy import so non-Windows environments can still import the module
    try:
        import win32com.client as win32  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "pywin32 is required for XLSB conversion. "
            "Install with: pip install pywin32"
        ) from e

    # Build output filename — use original stem with .xlsx extension
    output_path = output_dir / f"{xlsb_path.stem}.xlsx"

    # Copy source to a temp location to avoid file-lock issues
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_xlsb = Path(tmpdir) / xlsb_path.name
        shutil.copy2(xlsb_path, tmp_xlsb)

        excel = None
        wb = None
        try:
            logger.info("Converting %s → %s", xlsb_path.name, output_path.name)
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            # Disable macros so VBA compile errors (e.g. missing PtrSafe
            # in 64-bit Office) don't pop blocking dialogs during batch runs.
            # msoAutomationSecurityForceDisable = 3
            excel.AutomationSecurity = 3

            wb = excel.Workbooks.Open(
                str(tmp_xlsb),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            # Remove any existing output before writing
            if output_path.exists():
                output_path.unlink()

            wb.SaveAs(str(output_path), FileFormat=_XL_OPEN_XML_WORKBOOK)
        except Exception as e:
            raise RuntimeError(f"Excel conversion failed for {xlsb_path.name}: {e}") from e
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass

    # Update manifest
    manifest[file_hash] = output_path.name
    _save_manifest(manifest)

    return output_path


def convert_batch(
    xlsb_dir: Path,
    output_dir: Path | None = None,
    force: bool = False,
    pattern: str = "*.xlsb",
) -> list[Path]:
    """
    Convert every XLSB in a directory, returning their XLSX paths.

    Files are processed sequentially (Excel COM does not like concurrent
    instances). Failures are logged and skipped — the function returns
    the successful conversions.
    """
    xlsb_dir = Path(xlsb_dir)
    if not xlsb_dir.is_dir():
        raise NotADirectoryError(f"Not a directory: {xlsb_dir}")

    xlsb_files = sorted(xlsb_dir.glob(pattern))
    if not xlsb_files:
        logger.warning("No files matching %r in %s", pattern, xlsb_dir)
        return []

    results: list[Path] = []
    for xlsb in xlsb_files:
        try:
            results.append(convert_single(xlsb, output_dir=output_dir, force=force))
        except Exception as e:
            logger.error("Failed to convert %s: %s", xlsb.name, e)

    logger.info("Batch conversion complete: %d/%d succeeded", len(results), len(xlsb_files))
    return results
