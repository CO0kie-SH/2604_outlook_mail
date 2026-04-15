import argparse
import sys
import time
from pathlib import Path

import imap_outlook_oauth2
from client.internal_ws_client import InternalWSClient
from public.app_constants import AppConstants
from server.websocket_server import InternalWSServer


def resolve_default_client_account(constants: AppConstants) -> str:
    account_from_env = constants.ws_client_account
    if account_from_env:
        return account_from_env

    try:
        profile = constants.outlook_profile
        config_path = constants.resolved_config_path(Path(__file__).resolve().parent)
        row = imap_outlook_oauth2.OutlookMailService.load_outlook_config(config_path, profile)
        email_addr = (row.get("user", "") or "").strip()
        if email_addr:
            return email_addr
    except Exception:
        pass
    return "mail"


def build_arg_parser() -> argparse.ArgumentParser:
    constants = AppConstants.from_env()
    parser = argparse.ArgumentParser(description="WebSocket server/client threaded runtime")
    parser.add_argument("--server-host", default=constants.ws_server_host)
    parser.add_argument("--server-port", type=int, default=constants.ws_server_port)
    parser.add_argument("--client-account", default=resolve_default_client_account(constants))
    parser.add_argument("--client-password", default=constants.ws_client_password)
    parser.add_argument("--log-level", default=constants.outlook_log_level, help="log level")
    parser.add_argument("--log-file", default=imap_outlook_oauth2.resolve_default_log_file(), help="log path")
    parser.add_argument(
        "--log-retention-days",
        type=int,
        default=constants.outlook_log_retention_days,
        help="log retention days",
    )
    return parser


def main() -> int:
    constants = AppConstants.from_env()
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
    logger.info(
        "env constants: profile=%s config_path=%s idle_check_interval_seconds=%s idle_zero_limit=%s",
        constants.outlook_profile,
        constants.resolved_config_path(project_dir),
        constants.idle_check_interval_seconds,
        constants.idle_zero_limit,
    )

    server = InternalWSServer(
        host=args.server_host,
        port=args.server_port,
        logger=logger,
        idle_check_interval_seconds=constants.idle_check_interval_seconds,
        idle_zero_limit=constants.idle_zero_limit,
    )
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
