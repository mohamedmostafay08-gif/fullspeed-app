#!/usr/bin/env python3
"""يبني docs/index.html من apps-script/Index.html (شغّله بعد أي تعديل في Index.html)."""
from pathlib import Path
root = Path(__file__).parent
src = (root / "apps-script" / "Index.html").read_text(encoding="utf-8")
head = """<link rel="manifest" href="manifest.webmanifest">
<meta name="theme-color" content="#080d0a">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Full Speed">
<link rel="icon" href="icons/icon-192.png">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
<script src="config.js"></script>"""
assert "<!--WEB_HEAD-->" in src
(root / "docs" / "index.html").write_text(src.replace("<!--WEB_HEAD-->", head), encoding="utf-8")
print("docs/index.html built")
