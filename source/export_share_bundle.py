from pathlib import Path

from build_ai_monitor_dashboard import build_html, build_payload


ROOT = Path("/Users/lilblackmac/Documents/New project")
SHARE_DIR = ROOT / "share" / "ai_monitor_dashboard_v1"
SHARE_HTML = SHARE_DIR / "index.html"


def main() -> None:
    SHARE_DIR.mkdir(parents=True, exist_ok=True)
    payload = build_payload()
    SHARE_HTML.write_text(build_html(payload, logo_src="./assets/logo.jpg"), encoding="utf-8")
    print(SHARE_HTML)


if __name__ == "__main__":
    main()
