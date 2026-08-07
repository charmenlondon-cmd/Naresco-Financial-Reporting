#!/usr/bin/env python3
"""
Dashboard server — serves static files from dashboard/ and exposes
/api/run-pipeline to trigger update-dashboard.bat from the browser.

Usage:  python serve.py [port]   (default port 8000)
"""
import http.server
import subprocess
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
DASHBOARD_DIR = BASE_DIR / 'dashboard'
SOURCE_DIR = BASE_DIR / 'source-files'
BAT_FILE = BASE_DIR / 'update-dashboard.bat'


class DashboardHandler(http.server.SimpleHTTPRequestHandler):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DASHBOARD_DIR), **kwargs)

    def do_POST(self):
        if self.path == '/api/run-pipeline':
            self._run_pipeline()
        else:
            self.send_error(404)

    def _run_pipeline(self):
        xlsx_files = [f for f in SOURCE_DIR.glob('*.xlsx') if f.is_file()]

        if not xlsx_files:
            self._send_json({
                'status': 'no_files',
                'message': 'No files to process at this time.'
            })
            return

        file_list = ', '.join(f.name for f in xlsx_files)
        try:
            result = subprocess.run(
                ['cmd.exe', '/c', str(BAT_FILE)],
                input='\n',
                capture_output=True,
                text=True,
                cwd=str(BASE_DIR),
                timeout=600
            )
            output = (result.stdout + result.stderr).strip()
            if result.returncode == 0:
                self._send_json({
                    'status': 'success',
                    'message': f'Successfully processed: {file_list}',
                    'output': output
                })
            else:
                self._send_json({
                    'status': 'error',
                    'message': 'Pipeline encountered errors — check the log below.',
                    'output': output
                })
        except subprocess.TimeoutExpired:
            self._send_json({
                'status': 'error',
                'message': 'Pipeline timed out after 10 minutes.',
                'output': ''
            })
        except Exception as e:
            self._send_json({
                'status': 'error',
                'message': f'Unexpected error: {e}',
                'output': ''
            })

    def _send_json(self, data):
        body = json.dumps(data).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass


if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    server = http.server.HTTPServer(('localhost', port), DashboardHandler)
    print(f'Dashboard running at http://localhost:{port}')
    print('Press Ctrl+C to stop.')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServer stopped.')
