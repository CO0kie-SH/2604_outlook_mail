import logging
from pathlib import Path


class FileOps:
    logger = logging.getLogger(__name__)

    @staticmethod
    def ensure_parent(path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        FileOps.logger.debug("ensure parent dir: %s", path.parent)

    @staticmethod
    def write_text(path: Path, content: str, encoding: str = "utf-8") -> Path:
        target = path.resolve()
        FileOps.ensure_parent(target)
        target.write_text(content, encoding=encoding)
        FileOps.logger.info("write text file: path=%s chars=%s encoding=%s", target, len(content), encoding)
        return target

    @staticmethod
    def read_text(path: Path, encoding: str = "utf-8") -> str:
        target = path.resolve()
        text = target.read_text(encoding=encoding)
        FileOps.logger.info("read text file: path=%s chars=%s encoding=%s", target, len(text), encoding)
        return text

    @staticmethod
    def write_bytes(path: Path, content: bytes) -> Path:
        target = path.resolve()
        FileOps.ensure_parent(target)
        target.write_bytes(content)
        FileOps.logger.info("write bytes file: path=%s bytes=%s", target, len(content))
        return target

    @staticmethod
    def read_bytes(path: Path) -> bytes:
        target = path.resolve()
        data = target.read_bytes()
        FileOps.logger.info("read bytes file: path=%s bytes=%s", target, len(data))
        return data
