#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
法律文件脱敏处理 - 右键菜单启动服务

功能：启动临时 HTTP 服务器，提供 HTML 应用和目标 DOCX 文件，
      使浏览器可以通过 fetch 加载本地文件。

用法：python context_menu_server.py <docx文件路径>
"""

import os
import sys
import json
import socket
import socketserver
import webbrowser
import threading
import urllib.parse
import base64
import mimetypes
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path


def find_free_port():
    """找到一个空闲端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def get_assets_dir():
    """获取 assets 目录路径"""
    # 从脚本位置推导：scripts/../assets/
    return Path(__file__).parent.parent / 'assets'


class RedactionHandler(SimpleHTTPRequestHandler):
    """HTTP 请求处理器"""

    def __init__(self, *args, target_file=None, assets_dir=None, **kwargs):
        self.target_file = target_file
        self.assets_dir = assets_dir
        super().__init__(*args, directory=str(assets_dir), **kwargs)

    def translate_path(self, path):
        """将 URL 路径映射到本地文件路径"""
        # 解析 URL 路径
        parsed = urllib.parse.urlparse(path)
        url_path = parsed.path

        # API 端点不翻译路径
        if url_path.startswith('/api/'):
            return ''

        # 正常静态文件：从 assets 目录提供
        return super().translate_path(url_path)

    def do_GET(self):
        """处理 GET 请求"""
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/api/load-file':
            self._serve_target_file()
        elif parsed.path == '/api/load-settings':
            self._load_settings()
        else:
            super().do_GET()

    def do_POST(self):
        """处理 POST 请求"""
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/api/save-settings':
            self._save_settings()
        else:
            self.send_error(404, 'Not Found')

    def _serve_target_file(self):
        """提供目标 DOCX 文件"""
        if not self.target_file or not os.path.exists(self.target_file):
            self.send_error(404, '目标文件不存在')
            return

        try:
            with open(self.target_file, 'rb') as f:
                data = f.read()

            filename = os.path.basename(self.target_file)
            # URL 编码文件名（HTTP header 只支持 ASCII）
            encoded_filename = urllib.parse.quote(filename)
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Length', str(len(data)))
            # 传递原始文件名（Base64 编码，前端解码使用）
            b64_name = base64.b64encode(filename.encode('utf-8')).decode('ascii')
            self.send_header('X-Filename-B64', b64_name)
            self.send_header('Content-Disposition',
                             f"inline; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}")
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            self.send_error(500, f'Failed to read file: {e}')

    def _save_settings(self):
        """保存用户设置到 user_settings.json"""
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            settings = json.loads(body.decode('utf-8'))

            settings_path = self.assets_dir / 'user_settings.json'
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'status': 'ok'}).encode('utf-8'))
        except Exception as e:
            self.send_error(500, f'Failed to save settings: {e}')

    def _load_settings(self):
        """读取 user_settings.json 并返回"""
        try:
            settings_path = self.assets_dir / 'user_settings.json'
            empty_settings = {
                'whitelist': [],
                'blacklist': [],
                'customTypes': {},
                'disabledRules': []
            }
            if not settings_path.exists():
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(empty_settings, ensure_ascii=False).encode('utf-8'))
                return

            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(settings, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_error(500, f'Failed to load settings: {e}')

    def log_message(self, format, *args):
        """静默日志（不打印到控制台）"""
        pass


class ThreadedHTTPServer(socketserver.ThreadingMixIn, HTTPServer):
    """多线程 HTTP 服务器"""
    daemon_threads = True
    allow_reuse_address = True


def main():
    # 检查参数
    if len(sys.argv) < 2:
        print("用法: python context_menu_server.py <docx文件路径>")
        sys.exit(1)

    file_path = os.path.abspath(sys.argv[1])

    if not os.path.exists(file_path):
        print(f"错误: 文件不存在 - {file_path}")
        sys.exit(1)

    if not file_path.lower().endswith('.docx'):
        print(f"错误: 仅支持 .docx 文件 - {file_path}")
        sys.exit(1)

    # 获取路径
    assets_dir = get_assets_dir()
    if not assets_dir.exists():
        print(f"错误: assets 目录不存在 - {assets_dir}")
        sys.exit(1)

    index_html = assets_dir / 'index.html'
    if not index_html.exists():
        print(f"错误: index.html 不存在 - {index_html}")
        sys.exit(1)

    # 找到空闲端口
    port = find_free_port()

    # 创建服务器
    def handler_factory(*args, **kwargs):
        return RedactionHandler(*args, target_file=file_path, assets_dir=assets_dir, **kwargs)

    server = ThreadedHTTPServer(('127.0.0.1', port), handler_factory)

    # 设置超时自动关闭（30 分钟）
    def auto_shutdown():
        import time
        time.sleep(30 * 60)
        print("超时自动关闭服务器")
        threading.Thread(target=server.shutdown, daemon=True).start()

    shutdown_thread = threading.Thread(target=auto_shutdown, daemon=True)
    shutdown_thread.start()

    # 打开浏览器
    url = f'http://127.0.0.1:{port}/index.html?auto-load=true'
    print(f"正在打开脱敏工具: {url}")
    webbrowser.open(url)

    # 启动服务器（阻塞，直到 shutdown）
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
