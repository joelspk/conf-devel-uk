from __future__ import annotations

import os
import io
import json
import mimetypes
import zipfile
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from mvp_app.reconciliation import InputFile, ReconciliationError, process_files

BASE_DIR = Path(__file__).resolve().parent
INDEX_FILE = BASE_DIR / "templates" / "index.html"
STATIC_DIR = BASE_DIR / "static"
DOWNLOADS: dict[str, str] = {}
STATE_FILE = Path(os.environ.get("CONFERENCIA_STATE_FILE", "/tmp/conferencia_bancaria_state.json"))

def _load_state() -> dict:
    try:
        if STATE_FILE.exists():
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _save_state(payload: dict) -> None:
    try:
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        current = _load_state()
        current.update(payload)
        STATE_FILE.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        # Persistência não deve impedir o processamento principal.
        pass

class MVPRequestHandler(BaseHTTPRequestHandler):
    def _send(self, body: bytes, status: int = 200, content_type: str = "text/html; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict, status: int = 200) -> None:
        self._send(json.dumps(payload, ensure_ascii=False).encode("utf-8"), status=status, content_type="application/json; charset=utf-8")

    def _parse_multipart_files(self) -> list[InputFile]:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length) if content_length > 0 else b""
        message = BytesParser(policy=default).parsebytes((f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n").encode("utf-8") + body)
        items: list[InputFile] = []
        if not message.is_multipart():
            return items
        for part in message.iter_parts():
            filename = part.get_filename()
            if not filename:
                continue
            items.append(InputFile(filename=filename, content=part.get_payload(decode=True) or b""))
        return items

    def _finish_result(self, result: dict) -> None:
        path = result.pop("excel_path", None)
        if path:
            token = Path(path).stem
            DOWNLOADS[token] = path
            result["excel_url"] = f"/api/download/{token}"
        last_date = result.get("summary", {}).get("last_conference_date")
        if last_date:
            _save_state({"last_conference_date": last_date})
        self._send_json(result)

    def do_GET(self) -> None:  # noqa: N802
        if self.path == "/":
            self._send(INDEX_FILE.read_bytes())
            return
        if self.path == "/api/state":
            self._send_json(_load_state())
            return
        if self.path.startswith("/static/"):
            file_path = STATIC_DIR / self.path.removeprefix("/static/")
            if not file_path.exists() or not file_path.is_file():
                self._send(b"Not found", status=HTTPStatus.NOT_FOUND)
                return
            mime_type, _ = mimetypes.guess_type(str(file_path))
            self._send(file_path.read_bytes(), content_type=mime_type or "application/octet-stream")
            return
        if self.path.startswith("/api/download/"):
            token = self.path.rsplit("/", 1)[-1]
            path = DOWNLOADS.get(token)
            if not path or not Path(path).exists():
                self._send(b"Arquivo nao encontrado", status=HTTPStatus.NOT_FOUND)
                return
            data = Path(path).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="conferencia_bancaria_resultado.xlsx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self._send(b"Not found", status=HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        try:
            if self.path in {"/api/process-movimento", "/api/process-folder"}:
                result = process_files(self._parse_multipart_files())
                if self.path == "/api/process-folder":
                    result["folder_upload"] = True
                self._finish_result(result)
                return
            if self.path == "/api/process-zip":
                uploaded = self._parse_multipart_files()
                if not uploaded:
                    self._send_json({"detail": "Envie um arquivo ZIP."}, status=400)
                    return
                files: list[InputFile] = []
                try:
                    with zipfile.ZipFile(io.BytesIO(uploaded[0].content)) as zf:
                        for info in zf.infolist():
                            if info.is_dir():
                                continue
                            files.append(InputFile(filename=info.filename, content=zf.read(info)))
                except zipfile.BadZipFile:
                    self._send_json({"detail": "O arquivo enviado não é um ZIP válido."}, status=400)
                    return
                result = process_files(files)
                result["zip_file"] = uploaded[0].filename
                self._finish_result(result)
                return
            self._send(b"Not found", status=HTTPStatus.NOT_FOUND)
        except ReconciliationError as exc:
            self._send_json({"detail": str(exc)}, status=400)
        except Exception as exc:
            self._send_json({"detail": f"Falha ao processar arquivos: {exc}"}, status=500)

def create_server(host: str = "0.0.0.0", port: int = 8000) -> ThreadingHTTPServer:
    return ThreadingHTTPServer((host, port), MVPRequestHandler)
