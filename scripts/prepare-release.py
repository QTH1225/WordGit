"""Validate the tagged source and write the matching release notes."""
import argparse
import os
from pathlib import Path
import re
import tomllib

VERSION = r"(?:0|[1-9][0-9]{0,8})\.(?:0|[1-9][0-9]{0,8})\.(?:0|[1-9][0-9]{0,8})(?:-beta\.[1-9][0-9]{0,8})?"


def verify(tag: str, root: Path) -> tuple[str, str]:
    if re.fullmatch("v" + VERSION, tag) is None:
        raise ValueError("Use a release tag such as v0.12.1 or v0.12.1-beta.1.")
    version = tag[1:]

    def read(path: str) -> str:
        return (root / path).read_text(encoding="utf-8-sig")

    cargo_version = tomllib.loads(read("Cargo.toml"))["workspace"]["package"]["version"]
    addin = read("src/AddIn.cs")
    wps = re.search(r'WORDGIT_VERSION\s*=\s*"([^"]+)"', read("addin-wps/js/version.js"))
    assembly = re.search(r'AssemblyInformationalVersion\("([^"]+)"\)', addin)
    if cargo_version != version or not wps or wps[1] != version or not assembly or assembly[1] != version:
        raise ValueError("Cargo, Word informational version, WPS and tag versions must match.")
    numeric = version.split("-", 1)[0] + ".0"
    for attribute in ("AssemblyVersion", "AssemblyFileVersion"):
        if f'{attribute}("{numeric}")' not in addin:
            raise ValueError(f"{attribute} must be {numeric}.")

    changelog = read("CHANGELOG.md")
    headings = list(re.finditer(r"(?m)^## ([^\n]+)$", changelog))
    for index, heading in enumerate(headings):
        if heading[1].strip() in (version, f"[{version}]"):
            end = headings[index + 1].start() if index + 1 < len(headings) else len(changelog)
            notes = changelog[heading.end():end].strip()
            if notes:
                return version, notes
    raise ValueError(f"CHANGELOG.md needs a non-empty ## {version} section.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("tag")
    parser.add_argument("notes", type=Path)
    args = parser.parse_args()
    try:
        version, notes = verify(args.tag, Path(__file__).resolve().parents[1])
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"{error}\n")
    args.notes.write_text(notes + "\n", encoding="utf-8")
    if output := os.environ.get("GITHUB_OUTPUT"):
        with open(output, "a", encoding="utf-8") as file:
            print(f"version={version}", file=file)
            print(f"prerelease={'true' if '-beta.' in version else 'false'}", file=file)
    print(version)


if __name__ == "__main__":
    main()
