"""Local HTTP bridge for the Data Audit Agent Python verification engine.

Run:
    python server.py --port 8000

Endpoints:
    GET  /health
    POST /api/verify
    GET  /download/<filename>
"""

from __future__ import annotations

import argparse
import cgi
import importlib.util
import json
import mimetypes
import shutil
import tempfile
import sys
import re
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from email.header import decode_header
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent
VERIFY_PATH = ROOT / "verify.py"
OUTPUT_DIR = ROOT.parent / "outputs"


def load_verify_module():
    spec = importlib.util.spec_from_file_location("verify_module", VERIFY_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load verification engine: {VERIFY_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


VERIFY = load_verify_module()


def normalize_upload_filename(filename: str, fallback: str) -> str:
    raw_name = filename or fallback
    try:
        decoded_parts = []
        for value, encoding in decode_header(raw_name):
            if isinstance(value, bytes):
                decoded_parts.append(value.decode(encoding or "utf-8", errors="ignore"))
            else:
                decoded_parts.append(value)
        raw_name = "".join(decoded_parts) or fallback
    except Exception:
        raw_name = raw_name or fallback

    name = Path(raw_name).name
    name = re.sub(r'[<>:"/\\\\|?*]', "_", name).strip().strip(".")
    return name or fallback


def json_response(handler: BaseHTTPRequestHandler, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type")
    handler.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
    handler.end_headers()
    handler.wfile.write(body)


class VerificationHandler(BaseHTTPRequestHandler):
    server_version = "DataAuditAgentHTTP/1.0"

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        return

    def do_OPTIONS(self):  # noqa: N802
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.end_headers()

    def do_GET(self):  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            return json_response(self, {"ok": True})

        if parsed.path.startswith("/download/"):
            filename = Path(parsed.path.replace("/download/", "", 1)).name
            file_path = OUTPUT_DIR / filename
            if not file_path.exists():
                return json_response(self, {"ok": False, "error": "file not found"}, HTTPStatus.NOT_FOUND)

            data = file_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", mimetypes.guess_type(str(file_path))[0] or "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(data)
            return

        return json_response(self, {"ok": False, "error": "not found"}, HTTPStatus.NOT_FOUND)

    def do_POST(self):  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path != "/api/verify":
            return json_response(self, {"ok": False, "error": "not found"}, HTTPStatus.NOT_FOUND)

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            return json_response(self, {"ok": False, "error": "multipart/form-data expected"}, HTTPStatus.BAD_REQUEST)

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
            },
            keep_blank_values=True,
        )

        report_items = form["report"] if "report" in form else []
        source_items = form["source"] if "source" in form else []
        if not isinstance(report_items, list):
            report_items = [report_items]
        if not isinstance(source_items, list):
            source_items = [source_items]

        if not report_items or not source_items:
            return json_response(
                self,
                {"ok": False, "error": "보고서와 source file을 모두 업로드해 주세요."},
                HTTPStatus.BAD_REQUEST,
            )

        temp_dir = Path(tempfile.mkdtemp(prefix="data-audit-"))
        try:
            report_paths = []
            source_paths = []

            for item in report_items:
                filename = normalize_upload_filename(item.filename or "report", "report")
                target = temp_dir / f"report_{len(report_paths) + 1}_{filename}"
                with target.open("wb") as f:
                    shutil.copyfileobj(item.file, f)
                report_paths.append(target)

            for item in source_items:
                filename = normalize_upload_filename(item.filename or "source", "source")
                target = temp_dir / f"source_{len(source_paths) + 1}_{filename}"
                with target.open("wb") as f:
                    shutil.copyfileobj(item.file, f)
                source_paths.append(target)

            result = VERIFY.run_verification(report_paths, source_paths)
            output_path = VERIFY.ensure_output_path(OUTPUT_DIR)
            VERIFY.write_excel(result, output_path)

            payload = {
                "ok": True,
                "summary": [
                    result["summary"]["total"],
                    result["summary"]["match"],
                    result["summary"]["mismatch"],
                    result["summary"]["missing"],
                    result["summary"]["unavailable"],
                ],
                "summaryObject": result["summary"],
                "rows": result["rows"],
                "notice": result["notice"],
                "reportCount": result["report_count"],
                "sourceCount": result["source_count"],
                "mismatchRows": [row for row in result["rows"] if row["result"] == "불일치"],
                "downloadUrl": f"http://127.0.0.1:{self.server.server_port}/download/{output_path.name}",
                "outputFileName": output_path.name,
            }
            return json_response(self, payload)
        except Exception as exc:  # noqa: BLE001
            return json_response(self, {"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


def parse_args():
    parser = argparse.ArgumentParser(description="Data Audit Agent verification HTTP bridge")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), VerificationHandler)
    print(f"Data Audit Agent server running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
