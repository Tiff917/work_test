from __future__ import annotations

import json
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from paper_generator import ParseError, TEMPLATE_OPTIONS, generate_professional_docx


BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "web"
OUTPUT_DIR = BASE_DIR / "output"
HOST = "127.0.0.1"
PORT = 8765


class FormatterHandler(BaseHTTPRequestHandler):
    server_version = "ProfessionalFormatter/1.0"

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._serve_file(WEB_DIR / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/static/"):
            rel_path = parsed.path.removeprefix("/static/")
            target = (WEB_DIR / rel_path).resolve()
            if WEB_DIR not in target.parents and target != WEB_DIR:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if target.suffix == ".css":
                self._serve_file(target, "text/css; charset=utf-8")
                return
            if target.suffix == ".js":
                self._serve_file(target, "application/javascript; charset=utf-8")
                return
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if parsed.path == "/api/templates":
            self._send_json({"templates": TEMPLATE_OPTIONS})
            return
        if parsed.path == "/download":
            params = parse_qs(parsed.query)
            name = params.get("name", [""])[0]
            safe_name = Path(unquote(name)).name
            if not safe_name:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            target = OUTPUT_DIR / safe_name
            if not target.exists():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self._serve_file(
                target,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                download_name=safe_name,
            )
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/generate":
            length = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(length)
            try:
                payload = json.loads(raw_body.decode("utf-8"))
            except json.JSONDecodeError:
                self._send_json({"ok": False, "error": "無法讀取送出的資料。"}, status=HTTPStatus.BAD_REQUEST)
                return

            raw_text = (payload.get("content") or "").strip()
            template_key = (payload.get("template") or "thesis").strip()
            file_name = (payload.get("fileName") or "專業文件初稿").strip()

            if not raw_text:
                self._send_json({"ok": False, "error": "請先貼上內容。"}, status=HTTPStatus.BAD_REQUEST)
                return

            safe_name = _slugify_file_name(file_name)
            output_path = OUTPUT_DIR / f"{safe_name}.docx"
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

            try:
                generated = generate_professional_docx(raw_text, output_path, template_key=template_key)
            except ParseError as exc:
                self._send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
                return
            except Exception as exc:  # noqa: BLE001
                self._send_json({"ok": False, "error": f"生成失敗：{exc}"}, status=HTTPStatus.INTERNAL_SERVER_ERROR)
                return

            self._send_json(
                {
                    "ok": True,
                    "fileName": generated.name,
                    "outputPath": str(generated),
                    "downloadUrl": f"/download?name={generated.name}",
                }
            )
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        return

    def _serve_file(self, path: Path, content_type: str, download_name: str | None = None) -> None:
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        if download_name:
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        self.wfile.write(content)

    def _send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def _slugify_file_name(file_name: str) -> str:
    cleaned = "".join(ch for ch in file_name if ch not in '<>:"/\\|?*').strip().rstrip(".")
    return cleaned or "專業文件初稿"


def run_server() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((HOST, PORT), FormatterHandler)
    url = f"http://{HOST}:{PORT}"
    print(f"Professional formatter running at {url}")

    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    run_server()
