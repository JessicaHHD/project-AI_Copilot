from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any
from urllib.parse import urlparse

from .config import PROJECT_ROOT
from .gui_readers import load_dashboard


def json_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(payload, ensure_ascii=False).encode('utf-8')


def resolve_path(path_text: str) -> Path:
    cleaned = path_text.strip().strip(chr(34))
    path = Path(cleaned)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.resolve()


def open_local_path(path_text: str) -> dict[str, Any]:
    target = resolve_path(path_text)
    if not target.exists():
        return {'ok': False, 'message': f'路径不存在：{target}'}
    if os.name == 'nt':
        os.startfile(str(target))  # type: ignore[attr-defined]
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', str(target)])
    else:
        subprocess.Popen(['xdg-open', str(target)])
    return {'ok': True, 'message': '已打开路径。', 'path': str(target)}


def start_cli() -> dict[str, Any]:
    run_bat = PROJECT_ROOT / 'run.bat'
    if not run_bat.exists():
        return {'ok': False, 'message': f'未找到命令行入口：{run_bat}'}
    if os.name == 'nt':
        subprocess.Popen(['cmd', '/c', 'start', '', str(run_bat)], cwd=str(PROJECT_ROOT))
    else:
        subprocess.Popen([str(run_bat)], cwd=str(PROJECT_ROOT))
    return {'ok': True, 'message': '已启动现有命令行工具入口；GUI 不会自动执行提报/查价流程。'}


class GuiApiHandler(BaseHTTPRequestHandler):
    server_version = 'NewcomerGuiApi/0.1'

    def send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json_bytes(payload)
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self) -> dict[str, Any]:
        length = int(self.headers.get('Content-Length') or 0)
        if not length:
            return {}
        raw = self.rfile.read(length).decode('utf-8')
        return json.loads(raw or '{}')

    def do_OPTIONS(self) -> None:
        self.send_json({'ok': True})

    def do_GET(self) -> None:
        route = urlparse(self.path).path
        if route == '/api/dashboard':
            self.send_json(load_dashboard())
        elif route == '/api/health':
            self.send_json({'ok': True, 'service': 'newcomer-gui-api'})
        else:
            self.send_json({'ok': False, 'message': '接口不存在。'}, 404)

    def do_POST(self) -> None:
        route = urlparse(self.path).path
        try:
            payload = self.read_body()
            if route == '/api/open-path':
                path_text = str(payload.get('path') or '')
                if not path_text:
                    self.send_json({'ok': False, 'message': '缺少 path 参数。'}, 400)
                    return
                self.send_json(open_local_path(path_text))
            elif route == '/api/start-cli':
                self.send_json(start_cli())
            else:
                self.send_json({'ok': False, 'message': '接口不存在。'}, 404)
        except Exception as exc:
            self.send_json({'ok': False, 'message': str(exc)}, 500)

    def log_message(self, format: str, *args: Any) -> None:
        print('[gui-api]', format % args)


def main() -> None:
    parser = argparse.ArgumentParser(description='新人价自动化工作台只读 API')
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8765)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), GuiApiHandler)
    print(f'新人价自动化工作台 API 已启动：http://{args.host}:{args.port}')
    print('当前接口只读取配置、输出文件和日志；不会执行筛品、查价、提报或审核回填。')
    server.serve_forever()


if __name__ == '__main__':
    main()
