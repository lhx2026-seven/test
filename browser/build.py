"""Rebuild the self-contained HTML using the vendored, pinned JavaScript files."""
from pathlib import Path
r=Path(__file__).parent
html=(r/'ui.html').read_text(encoding='utf-8')
for file in ['vendor/xlsx.full.min.js','vendor/pptxgen.bundle.js','app.js']:
    html+='\n<script>\n'+(r/file).read_text(encoding='utf-8').replace('</script','<\\/script')+'\n</script>'
(r/'銷售簡報工作室-v2.1.html').write_text(html+'</body></html>',encoding='utf-8')
