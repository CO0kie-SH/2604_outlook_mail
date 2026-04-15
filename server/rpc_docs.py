from typing import Any

from aiohttp import web


class InternalWSDocPages:
    async def handle_doc_mail(self, request: web.Request) -> web.Response:
        login_request = """{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "auth.login",
  "params": {
    "account": "MarkGordon7281@hotmail.com",
    "password": "***"
  },
  "unixtime_ms": 1776000000000
}"""
        login_response = """{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "success": true,
    "cookie": "***",
    "login_at": "2026-04-12T13:54:31.351000+00:00"
  },
  "request_unixtime_ms": 1776000000000,
  "response_unixtime_ms": 1776000000123
}"""
        title_request = """{
  "jsonrpc": "2.0",
  "id": 1001,
  "method": "title",
  "params": {
    "cookie": "***",
    "folder_name": "Inbox",
    "known_max_uid": 1024,
    "incremental_count": 10
  },
  "unixtime_ms": 1776000000000
}"""
        title_response = """{
  "jsonrpc": "2.0",
  "id": 1001,
  "result": {
    "success": true,
    "folder_name": "Inbox",
    "titles": [
      {
        "mail_id": "1",
        "uid": "2",
        "message_id": "<...>",
        "title": "New app(s) connected...",
        "sender": "Microsoft account team <...>",
        "received_at": "2025-12-08T22:15:45+00:00",
        "received_unixtime_ms": 1765232145000
      }
    ],
    "update_unixtime_ms": 1776000000999
  }
}"""
        html = self._build_doc_page(
            title="Mail RPC Docs",
            subtitle="WebSocket endpoint: /ws/mail",
            sections=[
                {
                    "title": "Transport",
                    "body": "Protocol is JSON-RPC 2.0 over WebSocket. Open /ws/mail, then call auth.login -> auth.confirm -> outlook.token.acquire before querying folder count/title.",
                    "table": {
                        "headers": ["Item", "Value"],
                        "rows": [
                            ["WebSocket URL", "/ws/mail"],
                            ["Encoding", "JSON text frames"],
                            ["Server push", "mail.capabilities (notification)"],
                            ["Timing fields", "request_unixtime_ms / response_unixtime_ms"],
                        ],
                    },
                },
                {
                    "title": "Execution Order",
                    "table": {
                        "headers": ["Step", "Method", "Notes"],
                        "rows": [
                            ["1", "auth.login", "create session cookie"],
                            ["2", "auth.confirm", "confirm cookie and get enabled methods"],
                            ["3", "outlook.token.acquire", "get/reuse token and folders"],
                            ["4", "mail.folder.count", "query folder mail count"],
                            ["5", "title / title.base64a", "title = header only; base64a = include raw body Base64A"],
                            ["6", "feishu.notify", "optional summary notify"],
                            ["7", "auth.logout", "close session and cleanup"],
                        ],
                    },
                },
                {
                    "title": "auth.login",
                    "body": "Create a temporary login session and receive a cookie.",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["account", "string", "yes", "Mailbox account"],
                            ["password", "string", "yes", "Local login password"],
                        ],
                    },
                },
                {
                    "title": "auth.login response",
                    "table": {
                        "headers": ["Field", "Type", "Description"],
                        "rows": [
                            ["success", "bool", "Whether login request succeeded"],
                            ["cookie", "string", "Session token for subsequent calls"],
                            ["login_at", "string(ISO8601)", "UTC login time"],
                        ],
                    },
                    "code": login_request,
                },
                {
                    "title": "auth.login sample response",
                    "code": login_response,
                },
                {
                    "title": "auth.confirm / outlook.token.acquire",
                    "table": {
                        "headers": ["Method", "Required params", "Main result fields"],
                        "rows": [
                            ["auth.confirm", "cookie", "success, enabled_methods"],
                            ["outlook.token.acquire", "cookie", "success, token_cached, token_preview, folders[]"],
                        ],
                    },
                },
                {
                    "title": "title.base64a",
                    "body": "Fetch folder mail summary plus raw mail body encoded as URL-safe Base64 without '=' padding (field: Base64A). Supports incremental fetch via known_max_uid + incremental_count.",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["cookie", "string", "yes", "Session cookie"],
                            ["folder_name", "string", "yes", "Folder to query"],
                            ["known_max_uid", "int", "no", "Client local max UID; server searches UID>(known_max_uid)"],
                            ["incremental_count", "int", "no", "Limit returned incremental rows; usually online_count-local_count"],
                        ],
                    },
                },
                {
                    "title": "mail.folder.count",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["cookie", "string", "yes", "Session cookie"],
                            ["folder_name", "string", "yes", "Mailbox folder display name"],
                            ["current_count", "int", "no", "Client local count (for logging/sync)"],
                        ],
                    },
                },
                {
                    "title": "title",
                    "body": "Fetch header-level mail summary from one folder. Supports incremental fetch via known_max_uid + incremental_count.",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["cookie", "string", "yes", "Session cookie"],
                            ["folder_name", "string", "yes", "Folder to query"],
                            ["known_max_uid", "int", "no", "Client local max UID; server searches UID>(known_max_uid)"],
                            ["incremental_count", "int", "no", "Limit returned incremental rows; usually online_count-local_count"],
                        ],
                    },
                    "code": title_request,
                },
                {
                    "title": "title response fields",
                    "table": {
                        "headers": ["Field", "Type", "Description"],
                        "rows": [
                            ["mail_id", "string", "IMAP sequence number in current selected mailbox"],
                            ["uid", "string", "IMAP UID in this mailbox; more stable than mail_id"],
                            ["message_id", "string", "Message-ID from email header"],
                            ["title", "string", "Decoded subject"],
                            ["sender", "string", "Decoded From"],
                            ["received_at", "string(ISO8601)", "UTC timestamp"],
                            ["received_unixtime_ms", "int", "UTC epoch milliseconds"],
                        ],
                    },
                    "code": title_response,
                },
                {
                    "title": "title.base64a extra field",
                    "table": {
                        "headers": ["Field", "Type", "Description"],
                        "rows": [
                            ["Base64A", "string", "URL-safe Base64 encoded raw message (UTF-8 string, '=' removed)"],
                        ],
                    },
                },
                {
                    "title": "Server -> Client RPC",
                    "body": "Server may send JSON-RPC requests to online client via same WebSocket connection.",
                    "table": {
                        "headers": ["Method", "Direction", "Description"],
                        "rows": [
                            ["mail.folders.local.list", "server -> client", "client reads local *_folders.csv and returns rows"],
                            ["mail.client.force.logout", "server -> client", "client stops pull loop and calls auth.logout"],
                        ],
                    },
                },
                {
                    "title": "View Pages",
                    "table": {
                        "headers": ["Path", "Purpose", "Params"],
                        "rows": [
                            ["/view/mail/clients", "list online clients", "none"],
                            ["/view/mail/folders", "view one client's local folder list", "cookie"],
                            ["/view/mail/titles", "view title list for one folder (server-side title query)", "cookie, folder_name"],
                            ["/view/mail/logout", "trigger client force logout", "cookie"],
                        ],
                    },
                },
                {
                    "title": "Performance Logs",
                    "body": "Incremental title sync performance can be observed in log/*.log.",
                    "table": {
                        "headers": ["Keyword", "Meaning"],
                        "rows": [
                            ["perf title incremental", "client-side incremental rpc/write timing"],
                            ["perf query title / perf query title.base64a", "server RPC elapsed time"],
                            ["perf fetch title / perf fetch title.base64a", "server IMAP search/fetch timing details"],
                        ],
                    },
                },
                {
                    "title": "auth.logout",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["cookie", "string", "yes", "Session cookie to close"],
                        ],
                    },
                },
                {
                    "title": "Common error codes",
                    "table": {
                        "headers": ["Code", "Meaning", "Typical cause"],
                        "rows": [
                            ["-32700", "Parse error", "Malformed JSON"],
                            ["-32600", "Invalid Request", "jsonrpc != 2.0 or request shape invalid"],
                            ["-32601", "Method not found", "Unknown method name"],
                            ["-32602", "Invalid params", "Missing required params"],
                            ["-32003/-32004/-32005/-32006/-32009", "Business errors", "cookie invalid/token acquire/folder query/base64a failure"],
                        ],
                    },
                },
            ],
        )
        return web.Response(text=html, content_type="text/html", charset="utf-8")

    async def handle_doc_feishu(self, request: web.Request) -> web.Response:
        sample = """{
  "jsonrpc": "2.0",
  "id": 2001,
  "method": "feishu.notify",
  "params": {
    "cookie": "***",
    "title": "Outlook title sync",
    "body": "Summary lines...",
    "tag": "optional_tag"
  },
  "unixtime_ms": 1776000000000
}"""
        response_sample = """{
  "jsonrpc": "2.0",
  "id": 2001,
  "result": {
    "success": true,
    "results": {
      "T2025": true,
      "T2026": false
    },
    "success_count": 1
  }
}"""
        html = self._build_doc_page(
            title="Feishu RPC Docs",
            subtitle="Method: feishu.notify (executed on server side)",
            sections=[
                {
                    "title": "Overview",
                    "body": "Client sends one JSON-RPC method and server calls send_feishu_message. This keeps webhook execution and retries on server side.",
                },
                {
                    "title": "Input params",
                    "table": {
                        "headers": ["Param", "Type", "Required", "Description"],
                        "rows": [
                            ["cookie", "string", "yes", "Valid login session cookie"],
                            ["body", "string", "yes", "Message body to send"],
                            ["title", "string", "no", "Post title"],
                            ["tag", "string", "no", "Send only to one Feishu config tag"],
                        ],
                    },
                },
                {
                    "title": "Sample request",
                    "code": sample,
                },
                {
                    "title": "Sample response",
                    "code": response_sample,
                },
                {
                    "title": "Response fields",
                    "table": {
                        "headers": ["Field", "Type", "Description"],
                        "rows": [
                            ["success", "bool", "Method executed on server"],
                            ["results", "object", "Per-tag send result, e.g. {\"T2025\": true}"],
                            ["success_count", "int", "Number of tags with success=true"],
                        ],
                    },
                },
                {
                    "title": "Error codes",
                    "table": {
                        "headers": ["Code", "Meaning", "Typical cause"],
                        "rows": [
                            ["-32602", "Invalid params", "cookie/body missing"],
                            ["-32007", "invalid cookie", "session expired or not found"],
                            ["-32008", "feishu notify failed", "send_feishu_message exception"],
                        ],
                    },
                },
            ],
        )
        return web.Response(text=html, content_type="text/html", charset="utf-8")

    @staticmethod
    def _build_doc_page(title: str, subtitle: str, sections: list[dict[str, Any]]) -> str:
        cards: list[str] = []
        for item in sections:
            section_title = InternalWSDocPages._html_escape(str(item.get("title", "")).strip())
            body = InternalWSDocPages._html_escape(str(item.get("body", "")).strip())
            code = InternalWSDocPages._html_escape(str(item.get("code", "")).strip())
            lines = item.get("list", [])
            lines = lines if isinstance(lines, list) else []
            table = item.get("table", {})
            table = table if isinstance(table, dict) else {}
            headers = table.get("headers", [])
            rows = table.get("rows", [])
            headers = headers if isinstance(headers, list) else []
            rows = rows if isinstance(rows, list) else []

            content_parts: list[str] = []
            if body:
                content_parts.append(f"<p>{body}</p>")
            if lines:
                list_rows = "".join(f"<li>{InternalWSDocPages._html_escape(str(x))}</li>" for x in lines)
                content_parts.append(f"<ul>{list_rows}</ul>")
            if headers and rows:
                head_html = "".join(f"<th>{InternalWSDocPages._html_escape(str(x))}</th>" for x in headers)
                body_html_parts: list[str] = []
                for row in rows:
                    row = row if isinstance(row, list) else [row]
                    tds = "".join(f"<td>{InternalWSDocPages._html_escape(str(col))}</td>" for col in row)
                    body_html_parts.append(f"<tr>{tds}</tr>")
                body_html = "".join(body_html_parts)
                content_parts.append(
                    f"<div class=\"table-wrap\"><table><thead><tr>{head_html}</tr></thead><tbody>{body_html}</tbody></table></div>"
                )
            if code:
                content_parts.append(f"<pre><code>{code}</code></pre>")
            content_html = "".join(content_parts)
            cards.append(f"<section><h2>{section_title}</h2>{content_html}</section>")

        cards_html = "".join(cards)
        safe_title = InternalWSDocPages._html_escape(title)
        safe_subtitle = InternalWSDocPages._html_escape(subtitle)
        return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{safe_title}</title>
  <style>
    :root {{
      --bg: #f1f5f9;
      --card: #ffffff;
      --text: #0f172a;
      --muted: #475569;
      --line: #dbe2ea;
      --accent: #0284c7;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Trebuchet MS", "Segoe UI", Tahoma, sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at 5% -10%, #bae6fd 0%, rgba(186, 230, 253, 0) 35%),
        radial-gradient(circle at 100% 0%, #bfdbfe 0%, rgba(191, 219, 254, 0) 28%),
        var(--bg);
    }}
    .wrap {{
      max-width: 1120px;
      margin: 0 auto;
      padding: 28px 16px 40px;
    }}
    header {{
      margin-bottom: 14px;
      padding: 16px;
      border: 1px solid #cbd5e1;
      border-radius: 14px;
      background: linear-gradient(120deg, #ffffff, #eff6ff);
    }}
    h1 {{
      margin: 0;
      font-size: 32px;
      line-height: 1.2;
    }}
    .sub {{
      margin-top: 8px;
      color: var(--muted);
      font-size: 15px;
    }}
    .nav {{
      margin-top: 12px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }}
    .nav a {{
      text-decoration: none;
      color: #075985;
      background: #e0f2fe;
      border: 1px solid #bae6fd;
      border-radius: 999px;
      padding: 6px 12px;
      font-size: 13px;
    }}
    section {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px;
      margin-top: 12px;
      box-shadow: 0 4px 16px rgba(15, 23, 42, 0.04);
    }}
    h2 {{
      margin: 0 0 8px;
      font-size: 17px;
    }}
    p, li {{
      color: var(--muted);
      font-size: 14px;
      line-height: 1.55;
    }}
    ul {{
      margin: 0;
      padding-left: 18px;
    }}
    pre {{
      margin: 8px 0 0;
      background: #0b1020;
      color: #dbeafe;
      border-radius: 10px;
      padding: 12px;
      overflow: auto;
      border: 1px solid #1e293b;
      font-size: 12px;
      line-height: 1.45;
    }}
    code {{
      font-family: "Cascadia Code", Consolas, monospace;
    }}
    .table-wrap {{
      margin-top: 8px;
      overflow-x: auto;
      border: 1px solid #e2e8f0;
      border-radius: 10px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      background: #ffffff;
      min-width: 680px;
    }}
    thead th {{
      text-align: left;
      font-size: 12px;
      color: #334155;
      background: #f8fafc;
      border-bottom: 1px solid #e2e8f0;
      padding: 9px 10px;
    }}
    tbody td {{
      font-size: 13px;
      color: #1e293b;
      border-top: 1px solid #f1f5f9;
      padding: 9px 10px;
      vertical-align: top;
    }}
    .foot {{
      color: #64748b;
      font-size: 12px;
      margin-top: 14px;
    }}
  </style>
</head>
<body>
  <main class="wrap">
    <header>
      <h1>{safe_title}</h1>
      <div class="sub">{safe_subtitle}</div>
      <div class="nav">
        <a href="/doc/mail">/doc/mail</a>
        <a href="/doc/feishu">/doc/feishu</a>
        <a href="/ws/mail">/ws/mail</a>
      </div>
    </header>
    {cards_html}
    <div class="foot">This is a preview documentation page for concept validation.</div>
  </main>
</body>
</html>"""

    @staticmethod
    def _html_escape(text: str) -> str:
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
        )
