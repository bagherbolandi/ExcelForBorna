"""Tiny download server for the final workbook (served via sandbox preview)."""
import http.server
import os

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "dist")
PORT = 8080


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/ExcelForBorna-v2.0.xlsx")
            self.end_headers()
            return
        super().do_GET()

    def end_headers(self):
        if self.path.endswith(".xlsx"):
            self.send_header("Content-Disposition",
                             'attachment; filename="ExcelForBorna-v2.0.xlsx"')
        super().end_headers()

    def log_message(self, fmt, *args):
        print("[download]", self.address_string(), fmt % args, flush=True)


if __name__ == "__main__":
    os.chdir(ROOT)
    with http.server.ThreadingHTTPServer(("0.0.0.0", PORT), Handler) as httpd:
        print(f"serving {ROOT} on 0.0.0.0:{PORT}", flush=True)
        httpd.serve_forever()
