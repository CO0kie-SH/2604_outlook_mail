import os
from dataclasses import dataclass
from pathlib import Path


def _to_int(value: str, default: int) -> int:
    try:
        return int((value or "").strip())
    except (TypeError, ValueError):
        return default


@dataclass(frozen=True)
class AppConstants:
    ws_server_host: str
    ws_server_port: int
    ws_client_account: str
    ws_client_password: str
    outlook_log_level: str
    outlook_log_retention_days: int
    outlook_profile: str
    outlook_config_path: str
    idle_check_interval_seconds: int
    idle_zero_limit: int

    @classmethod
    def from_env(cls) -> "AppConstants":
        return cls(
            ws_server_host=os.getenv("WS_SERVER_HOST", "127.0.0.1"),
            ws_server_port=_to_int(os.getenv("WS_SERVER_PORT", "8765"), 8765),
            ws_client_account=os.getenv("WS_CLIENT_ACCOUNT", "").strip(),
            ws_client_password=os.getenv("WS_CLIENT_PASSWORD", "******"),
            outlook_log_level=os.getenv("OUTLOOK_LOG_LEVEL", "INFO"),
            outlook_log_retention_days=_to_int(os.getenv("OUTLOOK_LOG_RETENTION_DAYS", "30"), 30),
            outlook_profile=os.getenv("OUTLOOK_PROFILE", "outlook").strip() or "outlook",
            outlook_config_path=os.getenv("OUTLOOK_CONFIG_PATH", "config/OutLook.csv").strip() or "config/OutLook.csv",
            idle_check_interval_seconds=_to_int(os.getenv("WS_IDLE_CHECK_INTERVAL_SECONDS", "30"), 30),
            idle_zero_limit=_to_int(os.getenv("WS_IDLE_ZERO_LIMIT", "2"), 2),
        )

    def resolved_config_path(self, project_dir: Path) -> Path:
        path = Path(self.outlook_config_path)
        if path.is_absolute():
            return path
        return (project_dir / path).resolve()

