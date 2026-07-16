import argparse
import json
import os
import shutil
import sys
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SOURCE_DIR = PROJECT_ROOT / "source"


def load_config(config_path: str | None) -> dict[str, str]:
    if not config_path:
        return {}
    path = Path(config_path).expanduser().resolve()
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return {k: str(v) for k, v in data.items() if v is not None}


def apply_env(config: dict[str, str]) -> None:
    for key, value in config.items():
        os.environ[key] = value


def import_source_module(module_name: str):
    sys.path.insert(0, str(SOURCE_DIR))
    return __import__(module_name)


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def build_dashboard(config_path: str | None) -> Path:
    apply_env(load_config(config_path))
    output = Path(os.environ.get("AI_MONITOR_OUTPUT_HTML", PROJECT_ROOT / "dist" / "ai_monitor_dashboard_v1.html"))
    ensure_parent(output)
    os.environ["AI_MONITOR_OUTPUT_HTML"] = str(output)
    module = import_source_module("build_ai_monitor_dashboard")
    module.main()
    return output


def export_share(config_path: str | None, share_dir: str | None) -> Path:
    apply_env(load_config(config_path))
    target_dir = Path(share_dir or PROJECT_ROOT / "dist" / "site").expanduser().resolve()
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)

    dashboard_path = build_dashboard(config_path)
    shutil.copy2(dashboard_path, target_dir / "index.html")

    assets_src = SOURCE_DIR / "assets"
    if assets_src.exists():
        shutil.copytree(assets_src, target_dir / "assets", dirs_exist_ok=True)

    readme = target_dir / "README.txt"
    readme.write_text("Deploy this folder to any static hosting root.", encoding="utf-8")
    return target_dir


def zip_site(site_dir: str, output: str | None) -> Path:
    site_path = Path(site_dir).expanduser().resolve()
    output_path = Path(output or site_path.parent / f"{site_path.name}.zip").expanduser().resolve()
    ensure_parent(output_path)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in site_path.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(site_path.parent))
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="AI Class Monitor CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build-dashboard", help="Build dashboard HTML")
    p_build.add_argument("--config", help="Path to local JSON config")

    p_share = sub.add_parser("export-share", help="Export static share site")
    p_share.add_argument("--config", help="Path to local JSON config")
    p_share.add_argument("--share-dir", help="Output site directory")

    p_zip = sub.add_parser("zip-site", help="Zip an exported site directory")
    p_zip.add_argument("--site-dir", required=True, help="Existing site directory")
    p_zip.add_argument("--output", help="Zip output path")

    args = parser.parse_args()

    if args.command == "build-dashboard":
        result = build_dashboard(args.config)
    elif args.command == "export-share":
        result = export_share(args.config, args.share_dir)
    else:
        result = zip_site(args.site_dir, args.output)

    print(result)


if __name__ == "__main__":
    main()
