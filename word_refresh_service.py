from __future__ import annotations

import base64
import json
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from word_refresh_backend import refresh_docx_fields_with_local_word


SERVICE_HOST = os.environ.get("WORD_REFRESH_SERVICE_HOST", "127.0.0.1")
SERVICE_PORT = int(os.environ.get("WORD_REFRESH_SERVICE_PORT", "8765"))
SERVICE_TOKEN = os.environ.get("WORD_REFRESH_SERVICE_TOKEN", "").strip()


class WordRefreshHandler(BaseHTTPRequestHandler):
    server_version = "WordRefreshService/1.0"

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _is_authorized(self) -> bool:
        if not SERVICE_TOKEN:
            return True
        return self.headers.get("Authorization", "") == f"Bearer {SERVICE_TOKEN}"

    def do_GET(self) -> None:
        if self.path.rstrip("/") == "/health":
            self._send_json(HTTPStatus.OK, {"ok": True})
            return
        self._send_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})

    def do_POST(self) -> None:
        if self.path.rstrip("/") != "/refresh":
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})
            return
        if not self._is_authorized():
            self._send_json(HTTPStatus.UNAUTHORIZED, {"error": "unauthorized"})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0
        if content_length <= 0:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": "empty_body"})
            return

        try:
            request_payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            docx_bytes = base64.b64decode(request_payload["docx_base64"])
            appendix_number = request_payload.get("appendix_number")
            bookmarks_payload = request_payload.get("toc_page_range_bookmarks") or []
            toc_page_range_bookmarks = [
                (str(start), str(end))
                for start, end in bookmarks_payload
            ]
        except Exception:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": "invalid_payload"})
            return

        output_bytes = refresh_docx_fields_with_local_word(
            docx_bytes,
            appendix_number,
            toc_page_range_bookmarks,
        )
        self._send_json(
            HTTPStatus.OK,
            {"docx_base64": base64.b64encode(output_bytes).decode("ascii")},
        )

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    server = ThreadingHTTPServer((SERVICE_HOST, SERVICE_PORT), WordRefreshHandler)
    print(f"Word refresh service listening on http://{SERVICE_HOST}:{SERVICE_PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
