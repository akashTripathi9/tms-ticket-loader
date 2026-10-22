"""Wraps ticket-loader.html (artifact source, no doctype/head) into a complete
standalone document that can be opened from disk or hosted anywhere."""
from pathlib import Path
root = Path(__file__).resolve().parent.parent
src = (root / "ticket-loader.html").read_text()
out = root / "dist" / "NUPI Ticket Loader.html"
out.parent.mkdir(exist_ok=True)
out.write_text("<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n</head>\n<body>\n" + src + "\n</body>\n</html>\n")
print(out)
