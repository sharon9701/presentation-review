#!/usr/bin/env python3
"""Build one deterministic, cross-slide evidence manifest for a PPTX.

The script extracts repeatable source-file facts before a visual review:
slide/layout/master usage, grouped objects, repeated chrome, title geometry and
style, fonts, hidden slides/objects, notes, comments, metadata, media, geometry
overflow candidates and position/style drift. Rendering is intentionally left to
the caller because it depends on the target office application and environment.

Alongside the full evidence files, the script writes ``review-index.json``: a
compact, non-substitutive index for locating anomalous pages, hidden content and
objects that deserve extra attention during the default page-by-page deep review.

The default run includes layout/master shapes in the system evidence. Use
``--no-inherited`` only when those layers cannot be read, and record that
degradation in the resulting coverage section.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import re
import statistics
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Callable, Iterable
from zipfile import BadZipFile, ZipFile


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "a": A, "r": R}

LEAF_KINDS = {
    "sp": "sp",
    "pic": "pic",
    "graphicFrame": "graphicFrame",
    "cxnSp": "cxnSp",
}
GEOM_KEYS = ("x", "y", "cx", "cy")
IDENTITY = {"sx": 1.0, "sy": 1.0, "tx": 0.0, "ty": 0.0}
SCHEMA_VERSION = "1.2"


def q(ns: str, local: str) -> str:
    return "{" + ns + "}" + local


def file_sha256(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def local_name(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def rel_path(source: str, target: str) -> str:
    """Resolve an OOXML relationship target to a package path."""
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source)
    return posixpath.normpath(posixpath.join(base, target))


def rels_for(
    source: str,
    names: set[str],
    read_xml: Callable[[str], ET.Element | None],
) -> dict[str, dict[str, str]]:
    folder, filename = posixpath.split(source)
    rel_name = posixpath.join(folder, "_rels", filename + ".rels")
    root = read_xml(rel_name)
    result: dict[str, dict[str, str]] = {}
    if root is None:
        return result
    for item in root:
        rid = item.get("Id")
        if not rid:
            continue
        target = item.get("Target", "")
        result[rid] = {
            "target": rel_path(source, target),
            "type": item.get("Type", ""),
            "target_mode": item.get("TargetMode", ""),
        }
    return result


def _int_attr(node: ET.Element | None, key: str, default: int = 0) -> int:
    if node is None:
        return default
    try:
        return int(node.get(key, default))
    except (TypeError, ValueError):
        return default


def _point(node: ET.Element | None) -> tuple[int, int]:
    return (_int_attr(node, "x"), _int_attr(node, "y"))


def _size(node: ET.Element | None) -> tuple[int, int]:
    return (_int_attr(node, "cx"), _int_attr(node, "cy"))


def direct_xfrm(element: ET.Element, kind: str) -> ET.Element | None:
    """Find the transform belonging to this shape, not a descendant shape."""
    paths = {
        "sp": "./p:spPr/a:xfrm",
        "pic": "./p:spPr/a:xfrm",
        "graphicFrame": "./p:xfrm",
        "cxnSp": "./p:spPr/a:xfrm",
    }
    xfrm = element.find(paths.get(kind, ""), NS)
    if xfrm is not None:
        return xfrm
    # A few producers put the transform directly under a shape-specific
    # property node. Restrict the fallback to direct children.
    for child in list(element):
        for candidate in list(child):
            if local_name(candidate) == "xfrm":
                return candidate
    return None


def direct_geom(element: ET.Element, kind: str) -> dict[str, int] | None:
    xfrm = direct_xfrm(element, kind)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return {
        "x": _int_attr(off, "x"),
        "y": _int_attr(off, "y"),
        "cx": _int_attr(ext, "cx"),
        "cy": _int_attr(ext, "cy"),
    }


def group_transform(element: ET.Element) -> dict[str, int] | None:
    xfrm = element.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    ch_off = xfrm.find("a:chOff", NS)
    ch_ext = xfrm.find("a:chExt", NS)
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return None
    return {
        "off_x": _int_attr(off, "x"),
        "off_y": _int_attr(off, "y"),
        "ext_cx": _int_attr(ext, "cx"),
        "ext_cy": _int_attr(ext, "cy"),
        "ch_off_x": _int_attr(ch_off, "x"),
        "ch_off_y": _int_attr(ch_off, "y"),
        "ch_ext_cx": _int_attr(ch_ext, "cx"),
        "ch_ext_cy": _int_attr(ch_ext, "cy"),
    }


def group_to_transform(group: dict[str, int] | None) -> dict[str, float]:
    if not group:
        return dict(IDENTITY)
    sx = group["ext_cx"] / group["ch_ext_cx"] if group["ch_ext_cx"] else 1.0
    sy = group["ext_cy"] / group["ch_ext_cy"] if group["ch_ext_cy"] else 1.0
    return {
        "sx": sx,
        "sy": sy,
        "tx": group["off_x"] - group["ch_off_x"] * sx,
        "ty": group["off_y"] - group["ch_off_y"] * sy,
    }


def compose_transform(
    parent: dict[str, float], child: dict[str, float]
) -> dict[str, float]:
    """Compose child-local -> parent with parent -> global transforms."""
    return {
        "sx": parent["sx"] * child["sx"],
        "sy": parent["sy"] * child["sy"],
        "tx": parent["sx"] * child["tx"] + parent["tx"],
        "ty": parent["sy"] * child["ty"] + parent["ty"],
    }


def apply_transform(
    local: dict[str, int] | None, transform: dict[str, float]
) -> dict[str, int] | None:
    if local is None:
        return None
    return {
        "x": round(transform["sx"] * local["x"] + transform["tx"]),
        "y": round(transform["sy"] * local["y"] + transform["ty"]),
        "cx": round(transform["sx"] * local["cx"]),
        "cy": round(transform["sy"] * local["cy"]),
    }


def normalized(
    geom: dict[str, int] | None, slide_width: int, slide_height: int
) -> dict[str, float] | None:
    if not geom or not slide_width or not slide_height:
        return None
    return {
        key: round(value / (slide_width if key in ("x", "cx") else slide_height), 6)
        for key, value in geom.items()
    }


def text_of(element: ET.Element) -> str:
    return "".join(t.text or "" for t in element.findall(".//a:t", NS)).strip()


def package_properties(
    names: set[str], read_xml: Callable[[str], ET.Element | None]
) -> dict[str, object]:
    """Extract readable core/app/custom properties without assuming a schema."""
    properties: dict[str, object] = {"core": {}, "app": {}, "custom": {}, "parts": []}
    for path, bucket in (
        ("docProps/core.xml", "core"),
        ("docProps/app.xml", "app"),
        ("docProps/custom.xml", "custom"),
    ):
        if path not in names:
            continue
        properties["parts"].append(path)
        root = read_xml(path)
        if root is None:
            continue
        values: dict[str, object] = {}
        for node in list(root):
            key = local_name(node)
            value = "".join(node.itertext()).strip()
            if key == "property":
                key = node.get("name", key)
                value_node = next(iter(node), None)
                value = "".join(value_node.itertext()).strip() if value_node is not None else ""
            if key:
                values[key] = value
        properties[bucket] = values
    return properties


def package_comment_parts(names: set[str]) -> list[str]:
    """List comment and author parts so callers know whether hidden review text exists."""
    return sorted(
        name
        for name in names
        if name.startswith("ppt/comments")
        or name.startswith("ppt/modernComments")
    )


def package_comment_author_parts(names: set[str]) -> list[str]:
    return sorted(name for name in names if name.startswith("ppt/commentAuthors"))


def font_info(element: ET.Element) -> dict[str, object]:
    families: list[str] = []
    sizes: list[float] = []
    bold = False
    for node in element.findall(".//a:rPr", NS) + element.findall(".//a:defRPr", NS):
        for attr in ("latin", "ea", "cs"):
            value = node.get(attr)
            if value:
                families.append(value)
            child = node.find(f"a:{attr}", NS)
            if child is not None and child.get("typeface"):
                families.append(child.get("typeface", ""))
        raw = node.get("sz")
        if raw:
            try:
                sizes.append(round(int(raw) / 100, 2))
            except ValueError:
                pass
        if node.get("b") in ("1", "true"):
            bold = True
    return {
        "families": sorted(set(families)),
        "size_pt": max(sizes) if sizes else None,
        "bold": bold,
    }


def text_style_info(element: ET.Element) -> dict[str, object]:
    aligns = [
        node.get("algn")
        for node in element.findall(".//a:pPr", NS)
        if node.get("algn")
    ]
    colors: list[str] = []
    for node in element.findall(".//a:srgbClr", NS):
        if node.get("val"):
            colors.append("#" + node.get("val", "").upper())
    for node in element.findall(".//a:schemeClr", NS):
        if node.get("val"):
            colors.append("scheme:" + node.get("val", ""))
    return {"align": aligns[-1] if aligns else None, "colors": sorted(set(colors))}


def alpha_values(element: ET.Element) -> list[int]:
    values: list[int] = []
    for node in element.findall(".//a:alpha", NS):
        try:
            values.append(int(node.get("val", "")))
        except ValueError:
            continue
    return sorted(set(values))


def placeholder_info(element: ET.Element, kind: str) -> tuple[str, str]:
    if kind == "sp":
        ph = element.find("./p:nvSpPr/p:nvPr/p:ph", NS)
    else:
        ph = None
    if ph is None:
        return "", ""
    return ph.get("type", ""), ph.get("idx", "")


def shape_record(
    element: ET.Element,
    kind: str,
    rels: dict[str, dict[str, str]],
    slide_width: int,
    slide_height: int,
    media_hashes: dict[str, str],
    media_visual_hashes: dict[str, str],
    source_layer: str,
    source_xml: str,
    transform: dict[str, float],
    group_path: tuple[str, ...],
    z_order: int,
) -> dict[str, object]:
    nv_paths = {
        "sp": "./p:nvSpPr/p:cNvPr",
        "pic": "./p:nvPicPr/p:cNvPr",
        "graphicFrame": "./p:nvGraphicFramePr/p:cNvPr",
        "cxnSp": "./p:nvCxnSpPr/p:cNvPr",
    }
    nv = element.find(nv_paths[kind], NS)
    name = nv.get("name", "") if nv is not None else ""
    object_id = nv.get("id", "") if nv is not None else ""
    ph_type, ph_idx = placeholder_info(element, kind)
    local = direct_geom(element, kind)
    geom = apply_transform(local, transform)
    text = text_of(element)
    record: dict[str, object] = {
        "kind": kind,
        "name": name,
        "object_id": object_id,
        "hidden": (nv.get("hidden", "0").lower() in ("1", "true", "on")) if nv is not None else False,
        "alt_text": nv.get("descr", "") if nv is not None else "",
        "title_attr": nv.get("title", "") if nv is not None else "",
        "text": text,
        "ph_type": ph_type,
        "ph_idx": ph_idx,
        "local_geom": local,
        "geom": geom,
        "norm_geom": normalized(geom, slide_width, slide_height),
        "font": font_info(element) if text else {"families": [], "size_pt": None, "bold": False},
        "text_style": text_style_info(element) if text else {"align": None, "colors": []},
        "alpha_values": alpha_values(element),
        "z_order": z_order,
        "source_layer": source_layer,
        "source_xml": source_xml,
        "group_path": list(group_path),
        "is_group_child": bool(group_path),
    }
    if kind == "pic":
        blip = element.find(".//a:blip", NS)
        rid = blip.get(q(R, "embed")) if blip is not None else None
        if rid:
            target = rels.get(rid, {}).get("target", "")
            record["rel_id"] = rid
            record["media"] = target if target.startswith("ppt/media/") else None
            record["media_sha256"] = media_hashes.get(target)
            record["media_visual_hash"] = media_visual_hashes.get(target)
    return record


def walk_shapes(
    root: ET.Element,
    source_layer: str,
    source_xml: str,
    rels: dict[str, dict[str, str]],
    slide_width: int,
    slide_height: int,
    media_hashes: dict[str, str],
    media_visual_hashes: dict[str, str],
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    """Return visible leaf shapes and group metadata with global coordinates."""
    tree = root.find("./p:cSld/p:spTree", NS)
    if tree is None:
        return [], []
    leaves: list[dict[str, object]] = []
    groups: list[dict[str, object]] = []

    root_context = compose_transform(IDENTITY, group_to_transform(group_transform(tree)))

    def walk(parent: ET.Element, transform: dict[str, float], path: tuple[str, ...]) -> None:
        for element in list(parent):
            tag = local_name(element)
            if tag == "grpSp":
                nv = element.find("./p:nvGrpSpPr/p:cNvPr", NS)
                name = nv.get("name", "") if nv is not None else ""
                object_id = nv.get("id", "") if nv is not None else ""
                child_transform = compose_transform(transform, group_to_transform(group_transform(element)))
                group_geom = apply_transform(
                    {
                        "x": group_transform(element)["off_x"],
                        "y": group_transform(element)["off_y"],
                        "cx": group_transform(element)["ext_cx"],
                        "cy": group_transform(element)["ext_cy"],
                    }
                    if group_transform(element)
                    else None,
                    transform,
                )
                group_path = path + (name or f"group-{object_id}",)
                groups.append(
                    {
                        "name": name,
                        "object_id": object_id,
                        "source_layer": source_layer,
                        "source_xml": source_xml,
                        "group_path": list(group_path),
                        "geom": group_geom,
                        "norm_geom": normalized(group_geom, slide_width, slide_height),
                    }
                )
                walk(element, child_transform, group_path)
            elif tag in LEAF_KINDS:
                z_order = len(leaves)
                leaves.append(
                    shape_record(
                        element,
                        LEAF_KINDS[tag],
                        rels,
                        slide_width,
                        slide_height,
                        media_hashes,
                        media_visual_hashes,
                        source_layer,
                        source_xml,
                        transform,
                        path,
                        z_order,
                    )
                )

    walk(tree, root_context, ())
    return leaves, groups


def chrome_candidate(shape: dict[str, object]) -> bool:
    geom = shape.get("norm_geom") or {}
    text = str(shape.get("text", ""))
    name = str(shape.get("name", ""))
    top = float(geom.get("y", 1.0))
    left = float(geom.get("x", 1.0))
    right = left + float(geom.get("cx", 0.0))
    if float(geom.get("cx", 0.0)) > 0.8 and float(geom.get("cy", 0.0)) > 0.8:
        return False
    name_or_text = f"{name} {text}".lower()
    logo_words = ("logo", "tcl", "olympic", "五环", "奥运", "鸿鹄")
    return (
        shape.get("kind") == "pic" and top < 0.25 and (left < 0.35 or right > 0.65)
    ) or any(word in name_or_text for word in logo_words)


def title_candidate(shape: dict[str, object]) -> bool:
    if not shape.get("text"):
        return False
    ph = shape.get("ph_type")
    if ph in ("title", "ctrTitle", "subTitle"):
        return True
    geom = shape.get("norm_geom") or {}
    size = (shape.get("font") or {}).get("size_pt")
    return bool(float(geom.get("y", 1.0)) < 0.18 and size and float(size) >= 18)


def visual_hash(data: bytes) -> str | None:
    """Optional average hash; exact SHA remains the fallback identity."""
    try:
        from PIL import Image  # type: ignore
    except ImportError:
        return None
    try:
        image = Image.open(io.BytesIO(data)).convert("L").resize((8, 8))
        pixels = list(image.getdata())
        average = sum(pixels) / len(pixels)
        bits = "".join("1" if pixel >= average else "0" for pixel in pixels)
        return f"{int(bits, 2):016x}"
    except Exception:
        return None


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def visual_identity(shape: dict[str, object]) -> str:
    if shape.get("media_visual_hash"):
        return "image:ahash:" + str(shape["media_visual_hash"])
    if shape.get("media_sha256"):
        return "image:sha256:" + str(shape["media_sha256"])
    text = normalize_text(shape.get("text"))
    name = normalize_text(shape.get("name"))
    return f"{shape.get('kind', '')}|{text or name}"


def cluster(records: list[dict[str, object]], keys: tuple[str, ...]) -> list[dict[str, object]]:
    groups: dict[str, dict[str, object]] = {}
    for record in records:
        values: list[object] = []
        for key in keys:
            if key in GEOM_KEYS:
                values.append((record.get("norm_geom") or {}).get(key))
            else:
                values.append(record.get(key))
        signature = json.dumps(values, ensure_ascii=False, sort_keys=True)
        group = groups.setdefault(signature, {"signature": signature, "pages": [], "count": 0})
        group["pages"].append(record.get("slide"))
        group["count"] += 1
    return sorted(groups.values(), key=lambda item: (-int(item["count"]), str(item["signature"])))


def median_geom(records: list[dict[str, object]]) -> dict[str, float]:
    result: dict[str, float] = {}
    for key in GEOM_KEYS:
        values = [float((r.get("norm_geom") or {}).get(key)) for r in records if (r.get("norm_geom") or {}).get(key) is not None]
        if values:
            result[key] = round(statistics.median(values), 6)
    return result


def style_signature(record: dict[str, object]) -> tuple[object, ...]:
    font = record.get("font") or {}
    text_style = record.get("text_style") or {}
    return (
        tuple(font.get("families", [])),
        font.get("size_pt"),
        font.get("bold"),
        text_style.get("align"),
        tuple(text_style.get("colors", [])),
    )


def drift_deviations(
    records: list[dict[str, object]],
    threshold: float,
    kind: str,
    identity_fn: Callable[[dict[str, object]], str] = visual_identity,
) -> list[dict[str, object]]:
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for record in records:
        grouped[identity_fn(record)].append(record)
    deviations: list[dict[str, object]] = []
    for identity, group in grouped.items():
        if len(group) < 2 or len({item.get("slide") for item in group}) < 2:
            continue
        med = median_geom(group)
        style_values = {style_signature(item) for item in group}
        resources = {item.get("media_sha256") for item in group if item.get("media_sha256")}
        items: list[dict[str, object]] = []
        reasons: set[str] = set()
        for item in group:
            norm = item.get("norm_geom") or {}
            delta = {key: round(float(norm.get(key, 0)) - med.get(key, 0), 6) for key in GEOM_KEYS if key in med}
            position_drift = any(abs(delta.get(key, 0)) > threshold for key in ("x", "y"))
            size_drift = any(abs(delta.get(key, 0)) > threshold for key in ("cx", "cy"))
            if position_drift:
                reasons.add("position")
            if size_drift:
                reasons.add("size")
            if len(style_values) > 1:
                reasons.add("font/style")
            if len(resources) > 1:
                reasons.add("resource")
            if position_drift or size_drift or len(style_values) > 1 or len(resources) > 1:
                items.append(
                    {
                        "slide": item.get("slide"),
                        "source_layer": item.get("source_layer"),
                        "name": item.get("name"),
                        "object_id": item.get("object_id"),
                        "kind": item.get("kind"),
                        "group_path": item.get("group_path", []),
                        "norm_geom": norm,
                        "delta_from_median": delta,
                        "style": item.get("font"),
                        "media_sha256": item.get("media_sha256"),
                    }
                )
        if reasons:
            deviations.append(
                {
                    "kind": kind,
                    "visual_key": identity,
                    "pages": sorted({item.get("slide") for item in group if item.get("slide") is not None}),
                    "count": len(group),
                    "median_norm_geom": med,
                    "reasons": sorted(reasons),
                    "items": items,
                    "status": "unexplained",
                }
            )
    return sorted(deviations, key=lambda item: (-int(item["count"]), str(item["visual_key"])))


def title_role(record: dict[str, object]) -> str:
    geom = record.get("norm_geom") or {}
    x = float(geom.get("x", 0))
    y = float(geom.get("y", 0))
    region = "left" if x < 0.34 else "right" if x > 0.66 else "center"
    vertical = "top" if y < 0.08 else "upper" if y < 0.18 else "other"
    return f"{record.get('ph_type') or 'manual'}|{vertical}|{region}"


def title_style_deviations(records: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for record in records:
        grouped[title_role(record)].append(record)
    result: list[dict[str, object]] = []
    for role, group in grouped.items():
        if len({item.get("slide") for item in group}) < 2:
            continue
        styles = defaultdict(list)
        for item in group:
            styles[json.dumps(style_signature(item), ensure_ascii=False, default=str)].append(item)
        if len(styles) <= 1:
            continue
        result.append(
            {
                "kind": "title_style",
                "role": role,
                "pages": sorted({item.get("slide") for item in group if item.get("slide") is not None}),
                "variants": [
                    {
                        "style": json.loads(signature),
                        "pages": sorted({item.get("slide") for item in items if item.get("slide") is not None}),
                        "texts": [item.get("text") for item in items[:5]],
                    }
                    for signature, items in styles.items()
                ],
                "status": "unexplained",
            }
        )
    return result


def primary_titles(records: list[dict[str, object]]) -> list[dict[str, object]]:
    """Choose one likely main title per slide for low-noise cross-slide checks."""
    by_slide: dict[object, list[dict[str, object]]] = defaultdict(list)
    for record in records:
        by_slide[record.get("slide")].append(record)
    selected: list[dict[str, object]] = []
    for slide, candidates in by_slide.items():
        page_shapes = [item for item in candidates if item.get("source_layer") == "slide"]
        pool = page_shapes or candidates
        if not pool:
            continue
        def rank(item: dict[str, object]) -> tuple[float, float, float, float]:
            font_size = float((item.get("font") or {}).get("size_pt") or 0)
            geom = item.get("norm_geom") or {}
            area = float(geom.get("cx", 0)) * float(geom.get("cy", 0))
            return (font_size, area, -float(geom.get("y", 1)), -float(geom.get("x", 1)))
        selected.append(max(pool, key=rank))
    return selected


def _shape_key(shape: dict[str, object]) -> tuple[object, ...]:
    return (
        shape.get("slide"),
        shape.get("source_layer"),
        shape.get("source_xml"),
        shape.get("object_id"),
        shape.get("name"),
        tuple(shape.get("group_path", [])),
        shape.get("kind"),
    )


def build_review_index(
    all_slides: list[dict[str, object]],
    deviations: list[dict[str, object]],
    geometry_overflow: list[dict[str, object]],
    input_sha256: str,
    input_size_bytes: int,
    drift_threshold: float,
    inherited_enabled: bool,
) -> dict[str, object]:
    """Build a compact navigation index without copying slide text or metadata values."""
    page_reasons: dict[int, set[str]] = defaultdict(set)
    object_reasons: dict[tuple[object, ...], set[str]] = defaultdict(set)
    slide_shapes: dict[int, list[dict[str, object]]] = {
        int(slide["slide"]): list(slide.get("shapes", []))
        for slide in all_slides
    }

    def add_object(shape: dict[str, object], reason: str) -> None:
        slide_number = shape.get("slide")
        if slide_number is None:
            return
        shape_copy = dict(shape)
        shape_copy["slide"] = int(slide_number)
        object_reasons[_shape_key(shape_copy)].add(reason)
        page_reasons[int(slide_number)].add(reason)

    def matches(item: dict[str, object], shape: dict[str, object]) -> bool:
        if item.get("slide") != shape.get("slide"):
            return False
        for field in ("source_layer", "name", "kind", "object_id"):
            value = item.get("shape") if field == "name" and item.get("name") in (None, "") else item.get(field)
            if value not in (None, "") and value != shape.get(field):
                return False
        item_path = tuple(item.get("group_path", []))
        shape_path = tuple(shape.get("group_path", []))
        return not item_path or item_path == shape_path

    # Deviation records are the primary cross-slide anomaly signal.
    for deviation in deviations:
        reason = "cross_slide_" + str(deviation.get("kind", "deviation"))
        pages = [int(page) for page in deviation.get("pages", []) if page is not None]
        for page in pages:
            page_reasons[page].add(reason)
        items = deviation.get("items", [])
        if isinstance(items, list) and items:
            for item in items:
                if not isinstance(item, dict):
                    continue
                candidates = [
                    shape
                    for shape in slide_shapes.get(int(item.get("slide", -1)), [])
                    if matches(item, shape)
                ]
                if candidates:
                    for shape in candidates:
                        add_object(shape, reason)
                else:
                    # Keep a locator even when a producer omitted an object id.
                    page = item.get("slide")
                    if page is not None:
                        object_key = (
                            page,
                            item.get("source_layer"),
                            item.get("source_xml"),
                            item.get("object_id"),
                            item.get("name"),
                            tuple(item.get("group_path", [])),
                            item.get("kind"),
                        )
                        object_reasons[object_key].add(reason)
        elif str(deviation.get("kind")) == "title_style":
            for page in pages:
                for shape in slide_shapes.get(page, []):
                    if title_candidate(shape):
                        add_object(shape, reason)

    # Geometry candidates require visual confirmation and therefore get a deep-review flag.
    for item in geometry_overflow:
        page = item.get("slide")
        if page is None:
            continue
        page_number = int(page)
        page_reasons[page_number].add("geometry_overflow_candidate")
        candidates = [
            shape
            for shape in slide_shapes.get(page_number, [])
            if matches(item, shape)
        ]
        for shape in candidates:
            add_object(shape, "geometry_overflow_candidate")

    # Hidden pages/objects are always indexed; they are not automatically defects.
    for slide in all_slides:
        page = int(slide["slide"])
        if slide.get("hidden"):
            page_reasons[page].add("hidden_slide")
        for shape in slide.get("shapes", []):
            if shape.get("hidden"):
                add_object(shape, "hidden_object")

    # Density is a routing signal only. Thresholds are included for reproducibility.
    shape_counts = [int(slide.get("shape_count", 0)) for slide in all_slides]
    text_counts = [int(slide.get("text_count", 0)) for slide in all_slides]
    text_chars = [
        sum(len(str(shape.get("text", ""))) for shape in slide.get("shapes", []))
        for slide in all_slides
    ]
    median_shapes = statistics.median(shape_counts) if shape_counts else 0
    median_texts = statistics.median(text_counts) if text_counts else 0
    median_chars = statistics.median(text_chars) if text_chars else 0
    shape_limit = max(60, int(round(median_shapes * 1.8)))
    text_limit = max(20, int(round(median_texts * 1.8)))
    chars_limit = max(1200, int(round(median_chars * 1.8)))
    for slide, chars in zip(all_slides, text_chars):
        page = int(slide["slide"])
        if (
            int(slide.get("shape_count", 0)) >= shape_limit
            or int(slide.get("text_count", 0)) >= text_limit
            or chars >= chars_limit
        ):
            page_reasons[page].add("high_density_candidate")
            text_shapes = [shape for shape in slide.get("shapes", []) if shape.get("text")]
            for shape in sorted(text_shapes, key=lambda item: len(str(item.get("text", ""))), reverse=True)[:8]:
                add_object(shape, "high_density_candidate")

    object_index: list[dict[str, object]] = []
    for key, reasons in object_reasons.items():
        page, source_layer, source_xml, object_id, name, group_path, kind = key
        ref: dict[str, object] = {
            "slide": page,
            "source_layer": source_layer,
            "source_xml": source_xml,
            "object_id": object_id,
            "name": name,
            "group_path": list(group_path),
            "kind": kind,
            "reasons": sorted(reasons),
        }
        # Add geometry only; never copy potentially sensitive text into the index.
        for shape in slide_shapes.get(int(page), []):
            if _shape_key(shape) == key:
                ref["geom"] = shape.get("geom")
                ref["norm_geom"] = shape.get("norm_geom")
                ref["text_length"] = len(str(shape.get("text", "")))
                break
        object_index.append(ref)
    object_index.sort(key=lambda item: (int(item.get("slide", 0)), str(item.get("source_layer", "")), str(item.get("name", ""))))

    pages_index: list[dict[str, object]] = []
    anomaly_pages: list[dict[str, object]] = []
    hidden_pages: list[int] = []
    for slide in all_slides:
        page = int(slide["slide"])
        reasons = sorted(page_reasons.get(page, set()))
        if slide.get("hidden"):
            hidden_pages.append(page)
        page_objects = [item for item in object_index if item.get("slide") == page]
        record = {
            "slide": page,
            "hidden": bool(slide.get("hidden")),
            "anomaly": bool(reasons),
            "anomaly_reasons": reasons,
            "deep_review_recommended": bool(reasons or page_objects),
            "deep_review_object_count": len(page_objects),
            "shape_count": int(slide.get("shape_count", 0)),
            "text_count": int(slide.get("text_count", 0)),
            "text_chars": text_chars[page - 1] if 0 < page <= len(text_chars) else 0,
            "notes_present": bool(slide.get("notes")),
            "hidden_object_count": sum(1 for shape in slide.get("shapes", []) if shape.get("hidden")),
        }
        pages_index.append(record)
        if reasons:
            anomaly_pages.append(
                {
                    "slide": page,
                    "reasons": reasons,
                    "deep_review_object_count": len(page_objects),
                }
            )

    cache_key = (
        f"presentation-review:{SCHEMA_VERSION}:{input_sha256}:"
        f"drift={drift_threshold:.6g}:inherited={'on' if inherited_enabled else 'off'}"
    )
    return {
        "schema_version": SCHEMA_VERSION,
        "input": {
            "sha256": input_sha256,
            "size_bytes": input_size_bytes,
        },
        "cache": {
            "key": cache_key,
            "reusable_artifacts": [
                "evidence.json",
                "system-matrix.json",
                "review-index.json",
                "slide_text.txt",
            ],
            "reuse_when": "input sha256, schema version, drift threshold and inherited-layer mode match",
        },
        "deep_review_policy": "默认逐页深审；本索引只用于定位异常和对象，不得据此跳过任何页面或维度。",
        "thresholds": {
            "shape_count": shape_limit,
            "text_count": text_limit,
            "text_chars": chars_limit,
            "shape_multiplier": 1.8,
            "text_multiplier": 1.8,
            "char_multiplier": 1.8,
        },
        "slide_count": len(all_slides),
        "hidden_pages": hidden_pages,
        "anomaly_pages": anomaly_pages,
        "deep_review_objects": object_index,
        "slides": pages_index,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--drift-threshold", type=float, default=0.01, help="Normalized geometry drift threshold (default: 0.01).")
    parser.add_argument("--no-inherited", action="store_true", help="Do not parse layout/master shapes; records this degradation in coverage.")
    args = parser.parse_args()
    if not args.pptx.is_file():
        parser.error(f"input not found: {args.pptx}")
    if args.drift_threshold < 0:
        parser.error("--drift-threshold must be non-negative")
    args.output.mkdir(parents=True, exist_ok=True)
    input_sha256 = file_sha256(args.pptx)
    input_size_bytes = args.pptx.stat().st_size

    errors: list[dict[str, str]] = []
    with ZipFile(args.pptx) as package:
        names = set(package.namelist())

        def read_xml(path: str) -> ET.Element | None:
            if path not in names:
                return None
            try:
                return ET.fromstring(package.read(path))
            except (KeyError, ET.ParseError) as exc:
                errors.append({"source": path, "stage": "xml", "error": str(exc)})
                return None

        def sha(path: str) -> str | None:
            if path not in names:
                return None
            return hashlib.sha256(package.read(path)).hexdigest()

        properties = package_properties(names, read_xml)
        comment_parts = package_comment_parts(names)
        comment_author_parts = package_comment_author_parts(names)

        media_hashes = {
            name: digest
            for name in names
            if name.startswith("ppt/media/") and (digest := sha(name))
        }
        media_visual_hashes: dict[str, str] = {}
        perceptual_available = False
        for media_name in media_hashes:
            try:
                digest = visual_hash(package.read(media_name))
            except KeyError:
                digest = None
            if digest:
                perceptual_available = True
                media_visual_hashes[media_name] = digest

        presentation = read_xml("ppt/presentation.xml")
        if presentation is None:
            raise ValueError("ppt/presentation.xml is missing or invalid")
        size = presentation.find("p:sldSz", NS)
        slide_width = _int_attr(size, "cx")
        slide_height = _int_attr(size, "cy")
        prels = rels_for("ppt/presentation.xml", names, read_xml)
        slide_ids = presentation.find("p:sldIdLst", NS)
        slides: list[dict[str, object]] = []
        for number, node in enumerate(list(slide_ids) if slide_ids is not None else [], 1):
            rid = node.get(q(R, "id"), "")
            target = prels.get(rid, {}).get("target", "")
            root = read_xml(target) if target else None
            hidden = root is not None and root.get("show", "1").lower() in ("0", "false", "off")
            slides.append({"number": number, "xml": target, "slide_id": node.get("id"), "hidden": hidden})

        notes: dict[str, str] = {}
        for note_xml in sorted(
            n for n in names if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml")
        ):
            nrels = rels_for(note_xml, names, read_xml)
            slide_target = next(
                (value["target"] for value in nrels.values() if value["type"].endswith("/slide")),
                None,
            )
            if slide_target:
                root = read_xml(note_xml)
                notes[slide_target] = text_of(root) if root is not None else ""

        master_layout: dict[str, dict[str, str]] = {}
        for slide in slides:
            slide_xml = str(slide["xml"])
            srels = rels_for(slide_xml, names, read_xml) if slide_xml else {}
            layout = next((value["target"] for value in srels.values() if value["type"].endswith("/slideLayout")), "")
            lrels = rels_for(layout, names, read_xml) if layout else {}
            master = next((value["target"] for value in lrels.values() if value["type"].endswith("/slideMaster")), "")
            mrels = rels_for(master, names, read_xml) if master else {}
            theme = next((value["target"] for value in mrels.values() if value["type"].endswith("/theme")), "")
            master_layout[slide_xml] = {"layout": layout, "master": master, "theme": theme}

        layer_cache: dict[tuple[str, str], dict[str, object]] = {}

        def parse_layer(path: str, layer: str) -> dict[str, object]:
            key = (path, layer)
            if key in layer_cache:
                return layer_cache[key]
            root = read_xml(path)
            if root is None:
                result: dict[str, object] = {"shapes": [], "groups": [], "error": "missing or invalid XML"}
                errors.append({"source": path, "stage": layer, "error": "missing or invalid XML"})
            else:
                rels = rels_for(path, names, read_xml)
                shapes, groups = walk_shapes(
                    root,
                    layer,
                    path,
                    rels,
                    slide_width,
                    slide_height,
                    media_hashes,
                    media_visual_hashes,
                )
                result = {"shapes": shapes, "groups": groups}
            layer_cache[key] = result
            return result

        all_slides: list[dict[str, object]] = []
        all_titles: list[dict[str, object]] = []
        chrome: list[dict[str, object]] = []
        geometry_overflow: list[dict[str, object]] = []
        font_counts: Counter[str] = Counter()
        all_layer_fonts: Counter[str] = Counter()
        unique_layer_shapes: list[dict[str, object]] = []

        for slide in slides:
            number = int(slide["number"])
            slide_xml = str(slide["xml"])
            slide_layer = parse_layer(slide_xml, "slide")
            mapping = master_layout.get(slide_xml, {})
            layout_layer = parse_layer(str(mapping.get("layout", "")), "layout") if not args.no_inherited and mapping.get("layout") else {"shapes": [], "groups": []}
            master_layer = parse_layer(str(mapping.get("master", "")), "master") if not args.no_inherited and mapping.get("master") else {"shapes": [], "groups": []}
            slide_shapes = [dict(shape, slide=number) for shape in slide_layer.get("shapes", [])]
            layout_shapes = [dict(shape, slide=number, inherited_for_slide=number) for shape in layout_layer.get("shapes", [])]
            master_shapes = [dict(shape, slide=number, inherited_for_slide=number) for shape in master_layer.get("shapes", [])]

            for shape in slide_shapes:
                for family in (shape.get("font") or {}).get("families", []):
                    font_counts[str(family)] += 1
                geom = shape.get("geom")
                if geom and (
                    geom["x"] < 0
                    or geom["y"] < 0
                    or geom["x"] + geom["cx"] > slide_width
                    or geom["y"] + geom["cy"] > slide_height
                ):
                    geometry_overflow.append(
                        {
                            "slide": number,
                            "source_layer": shape.get("source_layer"),
                            "shape": shape.get("name"),
                            "group_path": shape.get("group_path", []),
                            "kind": shape.get("kind"),
                            "geom": geom,
                        }
                    )

            candidates = slide_shapes + [
                shape
                for shape in layout_shapes + master_shapes
                if not shape.get("ph_type") or shape.get("text")
            ]
            for shape in candidates:
                if title_candidate(shape):
                    all_titles.append(shape)
                if chrome_candidate(shape):
                    chrome.append(shape)

            all_slides.append(
                {
                    "slide": number,
                    "xml": slide_xml,
                    "hidden": slide["hidden"],
                    "notes": notes.get(slide_xml, ""),
                    "master": mapping.get("master", ""),
                    "layout": mapping.get("layout", ""),
                    "theme": mapping.get("theme", ""),
                    "shape_count": len(slide_shapes),
                    "text_count": sum(1 for shape in slide_shapes if shape.get("text")),
                    "placeholder_shape_count": sum(1 for shape in slide_shapes if shape.get("ph_type")),
                    "page_level_shape_count": sum(1 for shape in slide_shapes if not shape.get("ph_type")),
                    "group_count": len(slide_layer.get("groups", [])),
                    "layout_shape_count": len(layout_shapes),
                    "master_shape_count": len(master_shapes),
                    "layout_group_count": len(layout_layer.get("groups", [])),
                    "master_group_count": len(master_layer.get("groups", [])),
                    "page_level_object_names": [
                        shape.get("name") for shape in slide_shapes if not shape.get("ph_type") and shape.get("name")
                    ],
                    "shapes": slide_shapes,
                    "groups": slide_layer.get("groups", []),
                    "inherited_shapes": {
                        "layout": layout_shapes,
                        "master": master_shapes,
                    },
                }
            )

        for (path, layer), parsed in layer_cache.items():
            if layer in ("layout", "master"):
                for shape in parsed.get("shapes", []):
                    unique_layer_shapes.append(shape)
                    for family in (shape.get("font") or {}).get("families", []):
                        all_layer_fonts[str(family)] += 1

        title_clusters = cluster(
            all_titles,
            ("source_layer", "ph_type", "font", "text_style", "x", "y", "cx", "cy"),
        )
        chrome_clusters = cluster(
            chrome,
            ("source_layer", "kind", "media_sha256", "text", "x", "y", "cx", "cy"),
        )
        masters = Counter(str(slide["master"]) for slide in all_slides)
        layouts = Counter(str(slide["layout"]) for slide in all_slides)
        chrome_deviations = drift_deviations(chrome, args.drift_threshold, "repeated_chrome")
        primary_title_records = primary_titles(all_titles)
        title_geometry_deviations = drift_deviations(
            primary_title_records,
            args.drift_threshold,
            "title_geometry",
            identity_fn=lambda record: "title-role:" + title_role(record),
        )
        title_styles = title_style_deviations(primary_title_records)
        deviations = chrome_deviations + title_geometry_deviations + title_styles

        inherited_status = "Not checked" if args.no_inherited else "Available"
        visual_status = "Available" if perceptual_available else "Not available"
        review_index = build_review_index(
            all_slides,
            deviations,
            geometry_overflow,
            input_sha256,
            input_size_bytes,
            args.drift_threshold,
            not args.no_inherited,
        )
        evidence = {
            "schema_version": SCHEMA_VERSION,
            "input": str(args.pptx.resolve()),
            "input_fingerprint": {
                "sha256": input_sha256,
                "size_bytes": input_size_bytes,
            },
            "cache": review_index["cache"],
            "review_index": {
                "path": "review-index.json",
                "status": "Available",
                "schema_version": SCHEMA_VERSION,
            },
            "slide_count": len(slides),
            "slide_size_emu": {"cx": slide_width, "cy": slide_height},
            "slide_size_in": {
                "width": round(slide_width / 914400, 4),
                "height": round(slide_height / 914400, 4),
            },
            "hidden_slides": [slide["number"] for slide in slides if slide["hidden"]],
            "notes": {
                "count": len(notes),
                "slides_with_notes": [slide["number"] for slide in slides if str(slide["xml"]) in notes],
            },
            "comments": {
                "parts": comment_parts,
                "author_parts": comment_author_parts,
                "count": len(comment_parts),
                "status": "Available" if comment_parts else "Not found",
            },
            "metadata": {
                "properties": properties,
                "status": "Available" if properties["parts"] else "Not found",
            },
            "media": {
                "count": len(media_hashes),
                "sha256": media_hashes,
                "visual_hash": media_visual_hashes,
                "visual_similarity": {"status": visual_status, "method": "8x8 average hash"},
            },
            "masters_used": dict(masters),
            "layouts_used": dict(layouts),
            "fonts": dict(font_counts),
            "fonts_all_layers": dict(all_layer_fonts),
            "geometry_overflow_candidates": geometry_overflow,
            # Backward-compatible alias; explicitly states this is not text overflow.
            "overflow_candidates": geometry_overflow,
            "slides": all_slides,
            "errors": errors,
            "system_consistency": {
                "title_candidates": all_titles,
                "title_clusters": title_clusters,
                "repeated_chrome_candidates": chrome,
                "repeated_chrome_clusters": chrome_clusters,
                "master_layout_theme_by_slide": [
                    {
                        "slide": slide["slide"],
                        "master": slide["master"],
                        "layout": slide["layout"],
                        "theme": slide["theme"],
                        "slide_shapes": slide["shape_count"],
                        "layout_shapes": slide["layout_shape_count"],
                        "master_shapes": slide["master_shape_count"],
                        "slide_groups": slide["group_count"],
                        "layout_groups": slide["layout_group_count"],
                        "master_groups": slide["master_group_count"],
                    }
                    for slide in all_slides
                ],
                "inheritance_vs_page_override": {
                    "method": "placeholder and layer provenance are recorded separately; exact application inheritance still requires target-app confirmation",
                    "status": inherited_status,
                    "by_slide": [
                        {
                            "slide": slide["slide"],
                            "placeholder_shape_count": slide["placeholder_shape_count"],
                            "page_level_shape_count": slide["page_level_shape_count"],
                            "layout_shape_count": slide["layout_shape_count"],
                            "master_shape_count": slide["master_shape_count"],
                            "page_level_object_names": slide["page_level_object_names"],
                        }
                        for slide in all_slides
                    ],
                },
                "deviations_and_exceptions": {
                    "status": "Available",
                    "drift_threshold": args.drift_threshold,
                    "items": deviations,
                    "unexplained_count": len(deviations),
                    "note": "离群项仅表示可观察偏差，是否为有意特例需作者确认。",
                },
                "coverage_status": {
                    "grouped_objects": "Available",
                    "layout_shapes": inherited_status,
                    "master_shapes": inherited_status,
                    "drift_outliers": "Available",
                    "perceptual_image_similarity": visual_status,
                },
            },
            "coverage": {
                "full_auto_checked": [
                    "pages",
                    "hidden_slides",
                    "notes",
                    "comments",
                    "metadata",
                    "review-index",
                    "masters",
                    "layouts",
                    "themes",
                    "fonts",
                    "titles",
                    "repeated_chrome",
                    "geometry_overflow_candidates",
                    "media",
                    "grouped_objects",
                    "master_shapes" if not args.no_inherited else "master_shapes_not_requested",
                    "layout_shapes" if not args.no_inherited else "layout_shapes_not_requested",
                    "drift_outliers",
                ],
                "rendering": "not performed by this script",
                "fullsize_pages": [],
                "sampling_rule": "caller must record anomaly/cluster-based visual review",
                "not_checked": [
                    "target-application-playback",
                    "formal-accessibility-conformance",
                    "visual-rendering",
                    "text-overflow-visual",
                ]
                + (["perceptual-image-similarity"] if not perceptual_available else []),
                "status": {
                    "grouped_objects": "Available",
                    "master_shapes": inherited_status,
                    "layout_shapes": inherited_status,
                    "geometry_overflow": "Available",
                    "text_overflow_visual": "Not checked",
                    "perceptual_image_similarity": visual_status,
                },
            },
        }
        (args.output / "evidence.json").write_text(
            json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (args.output / "review-index.json").write_text(
            json.dumps(review_index, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (args.output / "system-matrix.json").write_text(
            json.dumps(evidence["system_consistency"], ensure_ascii=False, indent=2), encoding="utf-8"
        )
        with (args.output / "slide_text.txt").open("w", encoding="utf-8") as stream:
            for slide in all_slides:
                stream.write(
                    f"SLIDE {slide['slide']} hidden={slide['hidden']} master={slide['master']} layout={slide['layout']} shapes={slide['shape_count']} groups={slide['group_count']}\n"
                )
                for shape in slide["shapes"]:
                    if shape.get("text"):
                        stream.write("  " + str(shape["text"]) + "\n")
                if slide["notes"]:
                    stream.write("  [NOTES] " + str(slide["notes"]) + "\n")
                stream.write("\n")

    print(
        json.dumps(
            {
                "output": str(args.output.resolve()),
                "review_index": str((args.output / "review-index.json").resolve()),
                "input_sha256": input_sha256,
                "slide_count": len(slides),
                "hidden_slides": evidence["hidden_slides"],
                "master_count": len(masters),
                "layout_count": len(layouts),
                "title_candidates": len(all_titles),
                "chrome_candidates": len(chrome),
                "grouped_objects": sum(int(slide["group_count"]) for slide in all_slides),
                "inherited_shapes": sum(
                    int(slide["layout_shape_count"]) + int(slide["master_shape_count"])
                    for slide in all_slides
                ),
                "deviations": len(deviations),
                "anomaly_pages": len(review_index["anomaly_pages"]),
                "deep_review_objects": len(review_index["deep_review_objects"]),
                "errors": len(errors),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, ET.ParseError, BadZipFile) as exc:
        print(f"build_evidence.py: {exc}", file=sys.stderr)
        raise SystemExit(2)
