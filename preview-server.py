"""Preview server lokal JACNET: menyajikan dashboard-admin-protowifi.html
di http://localhost:8787 dengan auto-reload saat file berubah.
Hanya file dashboard yang disajikan (path lain 404), bind ke 127.0.0.1 saja.
"""
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(BASE, "dashboard-admin-protowifi.html")
PORT = 8787

POLL = """
<script>
(function(){
  var last = null;
  setInterval(function(){
    fetch('/__mtime', {cache:'no-store'}).then(function(r){return r.text()}).then(function(t){
      if (last === null) { last = t; }
      else if (t !== last) { location.reload(); }
    }).catch(function(){});
  }, 1000);
})();
</script>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _send(self, code, body, ctype="text/html; charset=utf-8"):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path.startswith("/__mtime"):
            try:
                mt = str(os.path.getmtime(TARGET))
            except OSError:
                mt = "0"
            self._send(200, mt, "text/plain")
            return
        if self.path in ("/", "/dashboard-admin-protowifi.html"):
            try:
                with open(TARGET, "r", encoding="utf-8", errors="replace") as f:
                    html = f.read()
            except OSError:
                self._send(500, "dashboard file not found", "text/plain")
                return
            if "</body>" in html:
                html = html.replace("</body>", POLL + "</body>", 1)
            else:
                html = html + POLL
            self._send(200, html)
            return
        self._send(404, "not found", "text/plain")


if __name__ == "__main__":
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
