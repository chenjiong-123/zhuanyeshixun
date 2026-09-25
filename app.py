"""本地网页应用入口：python app.py"""

from __future__ import annotations

import argparse
import json
import mimetypes
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from grader import grade_essay
from repository import PROJECT_ROOT, load_samples, public_samples


STATIC_DIR = PROJECT_ROOT / "static"
MAX_BODY_SIZE = 1024 * 1024


class EssayGraderHandler(BaseHTTPRequestHandler):
    server_version = "EssayGrader/1.0"

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"[访问] {self.address_string()} - {fmt % args}")

    def _json(self, payload: object, status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict[str, object]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ValueError("请求长度无效") from exc
        if not 0 < length <= MAX_BODY_SIZE:
            raise ValueError("请求为空或超过 1 MB")
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("请求不是有效 JSON") from exc
        if not isinstance(payload, dict):
            raise ValueError("请求格式错误")
        return payload

    def _static(self, name: str) -> None:
        path = (STATIC_DIR / name).resolve()
        if path != STATIC_DIR.resolve() and STATIC_DIR.resolve() not in path.parents:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = path.read_bytes()
        mime = {".html": "text/html", ".css": "text/css", ".js": "application/javascript"}.get(path.suffix.lower(), mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", f"{mime}; charset=utf-8" if mime.startswith("text/") or mime == "application/javascript" else mime)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        route = urlparse(self.path).path
        if route == "/api/health":
            self._json({"ok": True, "service": "智能英语作文批改系统", "version": "1.0"})
        elif route == "/api/bootstrap":
            self._json({"ok": True, "samples": public_samples(), "rubric": {"内容与切题": 25, "结构与衔接": 20, "词汇运用": 20, "语法与规范": 25, "篇幅与格式": 10}})
        elif route in {"/", "/index.html"}:
            self._static("index.html")
        else:
            self._static(route.lstrip("/"))

    def do_POST(self) -> None:  # noqa: N802
        route = urlparse(self.path).path
        try:
            payload = self._read_json()
            if route == "/api/grade":
                result = grade_essay(str(payload.get("prompt", "")), str(payload.get("essay", "")), str(payload.get("title", "")))
                self._json({"ok": True, "result": result})
            elif route == "/api/batch":
                results = []
                for item in load_samples():
                    grade = grade_essay(str(item["prompt"]), str(item["essay"]), str(item["title"]))
                    results.append({"id": item["id"], "title": item["title"], "reference_score": item["reference_score"], "score": grade["score"], "band": grade["band"], "issues": grade["statistics"]["issue_count"]})
                self._json({"ok": True, "results": results})
            else:
                self._json({"ok": False, "error": "接口不存在"}, HTTPStatus.NOT_FOUND)
        except ValueError as exc:
            self._json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # 防止单次请求使服务退出
            self._json({"ok": False, "error": f"处理失败：{exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)


def create_server(host: str = "127.0.0.1", port: int = 8000) -> ThreadingHTTPServer:
    return ThreadingHTTPServer((host, port), EssayGraderHandler)


def main() -> None:
    parser = argparse.ArgumentParser(description="智能英语作文批改系统")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    server = create_server(args.host, args.port)
    address = f"http://{args.host}:{server.server_address[1]}"
    print(f"系统已启动：{address}")
    print("按 Ctrl+C 可停止服务。")
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(address)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n系统已停止。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()

