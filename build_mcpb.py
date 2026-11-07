"""Pack the Claude Desktop extension: python build_mcpb.py -> dist/thinkcell-mcp-<version>.mcpb

An .mcpb is a zip with manifest.json at the root. Claude Desktop installs the
Python dependencies itself from pyproject.toml (manifest server.type "uv"), so
only the source goes in. Files matching .mcpbignore are left out.
"""
import fnmatch
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def ignored(rel: str, patterns: list[str]) -> bool:
    for pat in patterns:
        if pat.endswith("/"):  # a folder and everything under it
            if rel.startswith(pat) or f"/{pat}" in f"/{rel}":
                return True
        elif fnmatch.fnmatch(rel, pat) or fnmatch.fnmatch(Path(rel).name, pat):
            return True
    return False


def files_to_pack() -> list[str]:
    lines = (ROOT / ".mcpbignore").read_text(encoding="utf-8").splitlines()
    patterns = [l.strip() for l in lines if l.strip() and not l.startswith("#")]
    patterns += [".mcpbignore", "build_mcpb.py", ".gitignore", ".gitattributes",
                 ".connector_token", "*.pyc"]
    rels = (p.relative_to(ROOT).as_posix() for p in ROOT.rglob("*") if p.is_file())
    return sorted(r for r in rels if not ignored(r, patterns))


def main() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    for key in ("manifest_version", "name", "version", "description", "author", "server"):
        assert key in manifest, f"manifest.json is missing '{key}'"
    files = files_to_pack()
    for needed in ("manifest.json", manifest["server"]["entry_point"], "pyproject.toml",
                   manifest.get("icon", "manifest.json")):
        assert needed in files, f"{needed} would not be packed"
    out = ROOT / "dist" / f"{manifest['name']}-{manifest['version']}.mcpb"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in files:
            zf.write(ROOT / rel, rel)
    print(f"{out} ({len(files)} files, {out.stat().st_size // 1024} KB)")
    for rel in files:
        print("  " + rel)


if __name__ == "__main__":
    main()
