import sys
from pathlib import Path

import imap_outlook_oauth2


if __name__ == "__main__":
    parser = imap_outlook_oauth2.build_arg_parser()
    args = parser.parse_args()

    logger = imap_outlook_oauth2.setup_logger(args.log_level, args.log_file)
    imap_outlook_oauth2.cleanup_old_logs(args.log_file, args.log_retention_days, logger)

    runtime_dir = Path.cwd().resolve()
    project_dir = Path(__file__).resolve().parent
    log_dir = Path(args.log_file).resolve().parent
    python_path = Path(sys.executable).resolve()

    print(f"运行目录: {runtime_dir}")
    print(f"项目目录: {project_dir}")
    print(f"日志目录: {log_dir}")
    print(f"解释器路径: {python_path}")

    raise SystemExit(imap_outlook_oauth2.main(parsed_args=args, logger=logger))
