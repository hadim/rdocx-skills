"""The schemas of the spec oracle, fetched from their publishers and pinned by SHA-256.

- ECMA-376 5th edition Part 4 (December 2016, the text of ISO/IEC 29500-4:2016): the transitional XSDs.
- ECMA-376 5th edition Part 2 (December 2021, OPC): the content-types and relationships XSDs.
- W3C's xml.xsd, which the transitional schemas import without a location.

``ensure()`` downloads them once into ``.cache/schemas/`` and returns that folder:
``transitional/*.xsd``, ``opc/*.xsd`` and ``xml.xsd``.
"""

from __future__ import annotations

import hashlib
import io
import re
import sys
import urllib.request
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import common  # noqa: E402

SOURCES = {
    "part4": ("https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip",
              "bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd"),
    "part2": ("https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip",
              "1d489dc491168ea1f9e9a59063acc8dd5f02b4ad1d21aa7ec19ba9a58d020c70"),
    "xml": ("https://www.w3.org/2001/xml.xsd",
            "61960fb3131e38022caad5360e2f33a3382578ab3c80cd58bd74320ede61b20c"),
}
DEST = common.CACHE / "schemas"
XML_IMPORT = re.compile(rb'(<xsd:import\s+namespace="http://www\.w3\.org/XML/1998/namespace")\s*/>')


def _fetch(name: str) -> bytes:
    url, sha = SOURCES[name]
    cached = DEST / "downloads" / url.rsplit("/", 1)[1]
    if cached.exists() and hashlib.sha256(cached.read_bytes()).hexdigest() == sha:
        return cached.read_bytes()
    req = urllib.request.Request(url, headers={"User-Agent": "rdocx-skills-benchmark"})
    data = urllib.request.urlopen(req, timeout=120).read()
    got = hashlib.sha256(data).hexdigest()
    if got != sha:
        raise RuntimeError(f"{url}: SHA-256 {got}, expected {sha}")
    cached.parent.mkdir(parents=True, exist_ok=True)
    cached.write_bytes(data)
    return data


def _extract(outer: bytes, inner_name: str, dest: Path, patch=None):
    inner = zipfile.ZipFile(io.BytesIO(zipfile.ZipFile(io.BytesIO(outer)).read(inner_name)))
    dest.mkdir(parents=True, exist_ok=True)
    for n in inner.namelist():
        if n.endswith(".xsd"):
            data = inner.read(n)
            (dest / Path(n).name).write_bytes(patch(data) if patch else data)


def ensure() -> Path:
    if (DEST / "transitional" / "wml.xsd").exists() and (DEST / "opc" / "opc-contentTypes.xsd").exists():
        return DEST
    (DEST).mkdir(parents=True, exist_ok=True)
    (DEST / "xml.xsd").write_bytes(_fetch("xml"))
    # The transitional schemas import the XML namespace without a location: point them at xml.xsd.
    _extract(_fetch("part4"), "OfficeOpenXML-XMLSchema-Transitional.zip", DEST / "transitional",
             patch=lambda b: XML_IMPORT.sub(rb'\1 schemaLocation="../xml.xsd"/>', b))
    _extract(_fetch("part2"), "OpenPackagingConventions-XMLSchema.zip", DEST / "opc")
    return DEST


if __name__ == "__main__":
    print(ensure())
