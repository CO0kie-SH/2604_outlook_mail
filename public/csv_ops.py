import csv
import logging
from pathlib import Path
from typing import Any

from .file_ops import FileOps


class CsvOps:
    logger = logging.getLogger(__name__)

    @staticmethod
    def read_rows(path: Path, encoding: str = "utf-8-sig") -> list[dict[str, str]]:
        target = path.resolve()
        if not target.exists():
            CsvOps.logger.warning("csv not found, return empty rows: path=%s", target)
            return []
        with target.open("r", encoding=encoding, newline="") as f:
            rows = list(csv.DictReader(f))
        CsvOps.logger.info("read csv rows: path=%s rows=%s encoding=%s", target, len(rows), encoding)
        return rows

    @staticmethod
    def write_rows(
        path: Path,
        fieldnames: list[str],
        rows: list[dict[str, Any]],
        encoding: str = "utf-8-sig",
        quoting: int = csv.QUOTE_MINIMAL,
    ) -> Path:
        target = path.resolve()
        FileOps.ensure_parent(target)
        with target.open("w", encoding=encoding, newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=quoting)
            writer.writeheader()
            for row in rows:
                writer.writerow({key: row.get(key, "") for key in fieldnames})
        CsvOps.logger.info(
            "write csv rows: path=%s rows=%s fields=%s encoding=%s quoting=%s",
            target,
            len(rows),
            len(fieldnames),
            encoding,
            quoting,
        )
        return target
