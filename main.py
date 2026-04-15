import argparse
import os
import sys
import time
from pathlib import Path

import imap_outlook_oauth2
from client.internal_ws_client import InternalWSClient
from server.websocket_server import InternalWSServer


def resolve_default_client_account() -> str:
    account_from_env = os.getenv("WS_CLIENT_ACCOUNT", "").strip()
    if account_from_env:
        return account_from_env

    try:
        profile = os.getenv("OUTLOOK_PROFILE", "outlook")
        config_path = Path(imap_outlook_oauth2.resolve_default_config_path())
        row = imap_outlook_oauth2.OutlookMailService.load_outlook_config(config_path, profile)
        email_addr = (row.get("user", "") or "").strip()
        if email_addr:
            return email_addr
    except Exception:
        pass
    return "mail"


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="WebSocket server/client threaded runtime")
    parser.add_argument("--server-host", default=os.getenv("WS_SERVER_HOST", "127.0.0.1"))
    parser.add_argument("--server-port", type=int, default=int(os.getenv("WS_SERVER_PORT", "8765")))
    parser.add_argument("--client-account", default=resolve_default_client_account())
    parser.add_argument("--client-password", default=os.getenv("WS_CLIENT_PASSWORD", "******"))
    parser.add_argument("--log-level", default=os.getenv("OUTLOOK_LOG_LEVEL", "INFO"), help="log level")
    parser.add_argument("--log-file", default=imap_outlook_oauth2.resolve_default_log_file(), help="log path")
    parser.add_argument(
        "--log-retention-days",
        type=int,
        default=int(os.getenv("OUTLOOK_LOG_RETENTION_DAYS", "30")),
        help="log retention days",
    )
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()

    logger = imap_outlook_oauth2.setup_logger(args.log_level, args.log_file)
    imap_outlook_oauth2.cleanup_old_logs(args.log_file, args.log_retention_days, logger)

    runtime_dir = Path.cwd().resolve()
    project_dir = Path(__file__).resolve().parent
    log_dir = Path(args.log_file).resolve().parent
    python_path = Path(sys.executable).resolve()
    logger.info("runtime_dir=%s", runtime_dir)
    logger.info("project_dir=%s", project_dir)
    logger.info("log_dir=%s", log_dir)
    logger.info("python_path=%s", python_path)

    server = InternalWSServer(host=args.server_host, port=args.server_port, logger=logger)
    client = InternalWSClient(
        server_host=args.server_host,
        server_port=args.server_port,
        account=args.client_account,
        password=args.client_password,
        logger=logger,
    )

    server.start()
    ready = server.ready_event.wait(timeout=10)
    if not ready:
        logger.error("server failed to get ready in 10 seconds")
        server.stop()
        return 1

    client.start()
    logger.info("server/client threads started, press Ctrl+C to stop")

    try:
        while True:
            if server.shutdown_requested_event.is_set():
                logger.warning("shutdown requested by server idle guard, stopping process")
                break
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("keyboard interrupt received, shutting down")
    finally:
        client.stop()
        server.stop()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
