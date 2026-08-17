"""Expert OOXML package-part tools backed by pending session edits."""

from __future__ import annotations

import base64
import binascii
import copy
import hashlib
import io
import json
import mimetypes
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET


_CONTENT_TYPES_PART = "[Content_Types].xml"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_MIME = "application/vnd.openxmlformats-package.content-types+xml"
_RELATIONSHIPS_MIME = "application/vnd.openxmlformats-package.relationships+xml"
_MAX_INFRASTRUCTURE_BYTES = 8 * 1024 * 1024
_MAX_PART_BYTES = 64 * 1024 * 1024
_MAX_TRANSACTION_BYTES = 128 * 1024 * 1024
_MAX_OPERATIONS = 1000
_MAX_READ_BYTES = 256 * 1024
_MAX_LIST_PARTS = 1000
_MAX_RELATIONSHIPS_PER_PART = 100
_RELATIONSHIP_ID_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*$")
_EXTENSION_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+-]*$")
_CONTENT_TYPE_RE = re.compile(r"^[^\s/;]+/[^\s;]+(?:\s*;[^\r\n]*)?$")
_DRIVE_RE = re.compile(r"^[A-Za-z]:")


class PackageToolError(ValueError):
    """Raised when a package operation would create an unsafe or invalid state."""


def _empty_edits() -> dict[str, Any]:
    return {
        "upsert": {},
        "delete": [],
        "relationships": {},
        "content_types": {"defaults": {}, "overrides": {}},
    }


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _ensure_mapping(value: Any, label: str) -> dict:
    if not isinstance(value, dict):
        raise PackageToolError(f"{label} must be an object.")
    return value


def _ensure_list(value: Any, label: str) -> list:
    if not isinstance(value, list):
        raise PackageToolError(f"{label} must be an array.")
    return value


def _reject_unknown_keys(value: dict, allowed: set[str], label: str) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise PackageToolError(f"{label} contains unsupported fields: {unknown}.")


def _normalize_part_path(value: str, *, allow_root: bool = False) -> str:
    if not isinstance(value, str):
        raise PackageToolError("Package part path must be a string.")
    raw = value.strip()
    if allow_root and raw in {"", "/"}:
        return "/"
    if not raw:
        raise PackageToolError("Package part path cannot be empty.")
    if len(raw) > 1024:
        raise PackageToolError("Package part path is too long.")
    if "\x00" in raw or "\\" in raw or _DRIVE_RE.match(raw):
        raise PackageToolError(f"Unsafe package part path: {value!r}.")
    if raw.startswith("//"):
        raise PackageToolError(f"Unsafe package part path: {value!r}.")
    raw = raw.lstrip("/")
    if raw.endswith("/") or "//" in raw:
        raise PackageToolError(f"Package part path must identify a file: {value!r}.")
    segments = raw.split("/")
    decoded_segments = [unquote(segment) for segment in segments]
    if any(segment in {"", ".", ".."} for segment in segments + decoded_segments):
        raise PackageToolError(f"Path traversal is not allowed: {value!r}.")
    if any("/" in segment or "\\" in segment or "\x00" in segment for segment in decoded_segments):
        raise PackageToolError(f"Encoded path separators are not allowed: {value!r}.")
    return "/".join(segments)


def _normalize_prefix(value: str | None) -> str:
    if value is None or value.strip() in {"", "/"}:
        return ""
    raw = value.strip().lstrip("/")
    trailing_slash = raw.endswith("/")
    raw = raw.rstrip("/")
    normalized = _normalize_part_path(raw)
    return normalized + ("/" if trailing_slash else "")


def _is_relationship_part(path: str) -> bool:
    return path == "_rels/.rels" or ("/_rels/" in path and path.endswith(".rels"))


def _is_reserved_structural_part(path: str) -> bool:
    return path == _CONTENT_TYPES_PART or _is_relationship_part(path) or path.endswith(".rels")


def _relationship_part_for_source(source_part: str) -> str:
    source = _normalize_part_path(source_part, allow_root=True)
    if source == "/":
        return "_rels/.rels"
    directory = posixpath.dirname(source)
    filename = posixpath.basename(source)
    prefix = f"{directory}/" if directory else ""
    return f"{prefix}_rels/{filename}.rels"


def _source_for_relationship_part(path: str) -> str:
    normalized = _normalize_part_path(path)
    if normalized == "_rels/.rels":
        return "/"
    marker = "/_rels/"
    if marker not in normalized or not normalized.endswith(".rels"):
        raise PackageToolError(f"Not an OPC relationship part: {path!r}.")
    directory, filename = normalized.split(marker, 1)
    source_filename = filename[:-5]
    return f"{directory}/{source_filename}" if directory else source_filename


def _source_path(data: dict) -> Path:
    source = data.get("source")
    if not source:
        raise PackageToolError("Session does not contain a source OOXML package path.")
    path = Path(str(source))
    if not path.is_file():
        raise PackageToolError(f"Session source package does not exist: {path}.")
    if not zipfile.is_zipfile(path):
        raise PackageToolError(f"Session source is not a ZIP-based OOXML package: {path}.")
    return path


def _validate_xml_bytes(content: bytes, label: str) -> None:
    if len(content) > _MAX_PART_BYTES:
        raise PackageToolError(f"{label} exceeds the {_MAX_PART_BYTES}-byte limit.")
    upper = content[:4096].upper()
    if b"<!DOCTYPE" in upper or b"<!ENTITY" in upper:
        raise PackageToolError(f"{label} contains a prohibited DTD or entity declaration.")
    try:
        ET.fromstring(content)
    except (ET.ParseError, ValueError) as exc:
        raise PackageToolError(f"Invalid XML in {label}: {exc}.") from exc


def _validate_content_type(value: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise PackageToolError("Content type must be a non-empty string.")
    normalized = value.strip()
    if len(normalized) > 512 or not _CONTENT_TYPE_RE.match(normalized):
        raise PackageToolError(f"Invalid content type: {value!r}.")
    return normalized


def _validate_extension(value: str) -> str:
    if not isinstance(value, str):
        raise PackageToolError("Content-type extension must be a string.")
    normalized = value.strip().lstrip(".").lower()
    if not normalized or len(normalized) > 128 or not _EXTENSION_RE.match(normalized):
        raise PackageToolError(f"Invalid content-type extension: {value!r}.")
    return normalized


def _decode_upsert_content(operation: dict, label: str) -> bytes:
    encoding = str(operation.get("encoding", "text")).strip().lower()
    content = operation.get("content")
    if encoding in {"text", "utf8", "utf-8"}:
        if not isinstance(content, str):
            raise PackageToolError(f"{label}.content must be text for encoding={encoding!r}.")
        decoded = content.encode("utf-8")
    elif encoding == "base64":
        if not isinstance(content, str):
            raise PackageToolError(f"{label}.content must be a base64 string.")
        try:
            decoded = base64.b64decode(content, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise PackageToolError(f"Invalid base64 in {label}: {exc}.") from exc
    else:
        raise PackageToolError(f"Unsupported content encoding {encoding!r}; use text or base64.")
    if len(decoded) > _MAX_PART_BYTES:
        raise PackageToolError(f"{label} exceeds the {_MAX_PART_BYTES}-byte per-part limit.")
    return decoded


def _canonical_upsert_entry(content: bytes) -> dict[str, Any]:
    return {
        "data_base64": base64.b64encode(content).decode("ascii"),
        "size": len(content),
        "sha256": hashlib.sha256(content).hexdigest(),
    }


def _decode_canonical_upsert(path: str, entry: dict) -> bytes:
    _reject_unknown_keys(entry, {"data_base64", "size", "sha256"}, f"_package_edits.upsert[{path!r}]")
    encoded = entry.get("data_base64")
    if not isinstance(encoded, str):
        raise PackageToolError(f"Pending upsert {path!r} is missing data_base64.")
    try:
        content = base64.b64decode(encoded, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise PackageToolError(f"Pending upsert {path!r} contains invalid base64.") from exc
    if len(content) > _MAX_PART_BYTES:
        raise PackageToolError(f"Pending upsert {path!r} exceeds the per-part limit.")
    if entry.get("size") != len(content):
        raise PackageToolError(f"Pending upsert {path!r} has an invalid size digest.")
    digest = hashlib.sha256(content).hexdigest()
    if entry.get("sha256") != digest:
        raise PackageToolError(f"Pending upsert {path!r} has an invalid SHA-256 digest.")
    return content


def _validate_relationship_type(value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        raise PackageToolError("Relationship type must be a non-empty absolute URI.")
    normalized = value.strip()
    parsed = urlsplit(normalized)
    if not parsed.scheme or any(character.isspace() for character in normalized):
        raise PackageToolError(f"Relationship type must be an absolute URI: {value!r}.")
    return normalized


def _normalize_target_mode(value: Any) -> str:
    if value is None or str(value).strip().lower() == "internal":
        return "Internal"
    if str(value).strip().lower() == "external":
        return "External"
    raise PackageToolError("Relationship target_mode must be Internal or External.")


def _normalize_relationship_item(item: dict, label: str) -> dict[str, str]:
    _reject_unknown_keys(item, {"id", "type", "target", "target_mode"}, label)
    relationship_id = item.get("id")
    if not isinstance(relationship_id, str) or not _RELATIONSHIP_ID_RE.match(relationship_id):
        raise PackageToolError(f"Invalid relationship ID in {label}: {relationship_id!r}.")
    relationship_type = _validate_relationship_type(item.get("type"))
    target = item.get("target")
    if not isinstance(target, str) or not target.strip() or len(target) > 4096:
        raise PackageToolError(f"Relationship target in {label} must be a non-empty string.")
    target = target.strip()
    if "\x00" in target or "\\" in target:
        raise PackageToolError(f"Unsafe relationship target in {label}: {target!r}.")
    mode = _normalize_target_mode(item.get("target_mode"))
    return {
        "id": relationship_id,
        "type": relationship_type,
        "target": target,
        "target_mode": mode,
    }


def _resolve_internal_target(source_part: str, target: str) -> str:
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc:
        raise PackageToolError(
            f"Internal relationship from {source_part!r} cannot use an absolute URI target {target!r}."
        )
    decoded_path = unquote(parsed.path)
    if not decoded_path or "\\" in decoded_path or "\x00" in decoded_path:
        raise PackageToolError(f"Unsafe internal relationship target: {target!r}.")
    if decoded_path.startswith("/"):
        joined = decoded_path.lstrip("/")
    else:
        base = "" if source_part == "/" else posixpath.dirname(source_part)
        joined = posixpath.join(base, decoded_path)
    normalized = posixpath.normpath(joined)
    if normalized in {"", ".", ".."} or normalized.startswith("../"):
        raise PackageToolError(
            f"Internal relationship target escapes the package root: {target!r}."
        )
    return _normalize_part_path(normalized)


def _serialize_relationships(items: list[dict[str, str]]) -> bytes:
    ET.register_namespace("", _RELATIONSHIPS_NS)
    root = ET.Element(f"{{{_RELATIONSHIPS_NS}}}Relationships")
    for item in items:
        attrs = {"Id": item["id"], "Type": item["type"], "Target": item["target"]}
        if item["target_mode"] == "External":
            attrs["TargetMode"] = "External"
        ET.SubElement(root, f"{{{_RELATIONSHIPS_NS}}}Relationship", attrs)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _parse_relationships(content: bytes, rel_path: str) -> list[dict[str, str]]:
    if len(content) > _MAX_INFRASTRUCTURE_BYTES:
        raise PackageToolError(f"Relationship part {rel_path!r} is too large to inspect safely.")
    _validate_xml_bytes(content, rel_path)
    root = ET.fromstring(content)
    if _local_name(root.tag) != "Relationships":
        raise PackageToolError(f"Relationship part {rel_path!r} has an invalid root element.")
    result: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    for index, element in enumerate(root):
        if _local_name(element.tag) != "Relationship":
            continue
        item = _normalize_relationship_item(
            {
                "id": element.attrib.get("Id"),
                "type": element.attrib.get("Type"),
                "target": element.attrib.get("Target"),
                "target_mode": element.attrib.get("TargetMode"),
            },
            f"{rel_path}.Relationship[{index}]",
        )
        if item["id"] in seen_ids:
            raise PackageToolError(f"Duplicate relationship ID {item['id']!r} in {rel_path!r}.")
        seen_ids.add(item["id"])
        result.append(item)
    return result


def _parse_content_types(content: bytes) -> tuple[dict[str, str], dict[str, str]]:
    if len(content) > _MAX_INFRASTRUCTURE_BYTES:
        raise PackageToolError("[Content_Types].xml is too large to inspect safely.")
    _validate_xml_bytes(content, _CONTENT_TYPES_PART)
    root = ET.fromstring(content)
    if _local_name(root.tag) != "Types":
        raise PackageToolError("[Content_Types].xml has an invalid root element.")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for element in root:
        name = _local_name(element.tag)
        if name == "Default":
            extension = _validate_extension(element.attrib.get("Extension"))
            if extension in defaults:
                raise PackageToolError(f"Duplicate Default content-type mapping for {extension!r}.")
            defaults[extension] = _validate_content_type(element.attrib.get("ContentType"))
        elif name == "Override":
            part_name = _normalize_part_path(element.attrib.get("PartName"))
            if part_name in overrides:
                raise PackageToolError(f"Duplicate Override content-type mapping for {part_name!r}.")
            overrides[part_name] = _validate_content_type(element.attrib.get("ContentType"))
    return defaults, overrides


def _serialize_content_types(defaults: dict[str, str], overrides: dict[str, str]) -> bytes:
    ET.register_namespace("", _CONTENT_TYPES_NS)
    root = ET.Element(f"{{{_CONTENT_TYPES_NS}}}Types")
    for extension, content_type in sorted(defaults.items()):
        ET.SubElement(
            root,
            f"{{{_CONTENT_TYPES_NS}}}Default",
            {"Extension": extension, "ContentType": content_type},
        )
    for part_path, content_type in sorted(overrides.items()):
        ET.SubElement(
            root,
            f"{{{_CONTENT_TYPES_NS}}}Override",
            {"PartName": f"/{part_path}", "ContentType": content_type},
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _normalize_pending_edits(raw: Any) -> dict[str, Any]:
    if raw is None:
        return _empty_edits()
    value = _ensure_mapping(raw, "_package_edits")
    _reject_unknown_keys(value, {"upsert", "delete", "relationships", "content_types"}, "_package_edits")
    normalized = _empty_edits()

    upsert = _ensure_mapping(value.get("upsert", {}), "_package_edits.upsert")
    total_bytes = 0
    for raw_path, raw_entry in upsert.items():
        path = _normalize_part_path(raw_path)
        if _is_reserved_structural_part(path):
            raise PackageToolError(f"Pending direct structural edit is not allowed for {path!r}.")
        entry = _ensure_mapping(raw_entry, f"_package_edits.upsert[{path!r}]")
        content = _decode_canonical_upsert(path, entry)
        total_bytes += len(content)
        normalized["upsert"][path] = _canonical_upsert_entry(content)
    if total_bytes > _MAX_TRANSACTION_BYTES:
        raise PackageToolError("Pending package upserts exceed the transaction byte limit.")

    delete = _ensure_list(value.get("delete", []), "_package_edits.delete")
    normalized_delete: set[str] = set()
    for raw_path in delete:
        path = _normalize_part_path(raw_path)
        if _is_reserved_structural_part(path):
            raise PackageToolError(f"Pending direct structural delete is not allowed for {path!r}.")
        normalized_delete.add(path)
    if normalized_delete & set(normalized["upsert"]):
        overlap = sorted(normalized_delete & set(normalized["upsert"]))
        raise PackageToolError(f"Pending paths cannot be both upserted and deleted: {overlap}.")
    normalized["delete"] = sorted(normalized_delete)

    relationships = _ensure_mapping(value.get("relationships", {}), "_package_edits.relationships")
    for raw_source, raw_items in relationships.items():
        source = _normalize_part_path(raw_source, allow_root=True)
        items = _ensure_list(raw_items, f"_package_edits.relationships[{source!r}]")
        seen_ids: set[str] = set()
        normalized_items = []
        for index, raw_item in enumerate(items):
            item = _normalize_relationship_item(
                _ensure_mapping(raw_item, f"relationship[{index}]"),
                f"relationship[{index}]",
            )
            if item["id"] in seen_ids:
                raise PackageToolError(f"Duplicate relationship ID {item['id']!r} for source {source!r}.")
            seen_ids.add(item["id"])
            normalized_items.append(item)
        normalized["relationships"][source] = normalized_items

    content_types = _ensure_mapping(value.get("content_types", {}), "_package_edits.content_types")
    _reject_unknown_keys(content_types, {"defaults", "overrides"}, "_package_edits.content_types")
    defaults = _ensure_mapping(content_types.get("defaults", {}), "_package_edits.content_types.defaults")
    overrides = _ensure_mapping(content_types.get("overrides", {}), "_package_edits.content_types.overrides")
    for raw_extension, raw_content_type in defaults.items():
        extension = _validate_extension(raw_extension)
        normalized["content_types"]["defaults"][extension] = (
            None if raw_content_type is None else _validate_content_type(raw_content_type)
        )
    for raw_path, raw_content_type in overrides.items():
        path = _normalize_part_path(raw_path)
        normalized["content_types"]["overrides"][path] = (
            None if raw_content_type is None else _validate_content_type(raw_content_type)
        )
    return normalized


class _PackageView:
    def __init__(self, data: dict, edits: dict[str, Any]):
        self.data = data
        self.edits = edits
        self.source_path = _source_path(data)
        self.source_infos = self._load_source_infos()
        self._content_types_cache: tuple[dict[str, str], dict[str, str]] | None = None
        self._relationships_cache: dict[str, list[dict[str, str]]] | None = None

    def _load_source_infos(self) -> dict[str, zipfile.ZipInfo]:
        infos: dict[str, zipfile.ZipInfo] = {}
        with zipfile.ZipFile(self.source_path, "r") as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                path = _normalize_part_path(info.filename)
                if path != info.filename:
                    raise PackageToolError(f"Unsafe or non-canonical ZIP member path: {info.filename!r}.")
                if path in infos:
                    raise PackageToolError(f"Duplicate ZIP member path: {path!r}.")
                infos[path] = info
        if _CONTENT_TYPES_PART not in infos:
            raise PackageToolError("Source package is missing [Content_Types].xml.")
        return infos

    def part_paths(self) -> set[str]:
        paths = set(self.source_infos)
        paths.difference_update(self.edits["delete"])
        paths.update(self.edits["upsert"])
        for source, items in self.edits["relationships"].items():
            rel_path = _relationship_part_for_source(source)
            if items:
                paths.add(rel_path)
            else:
                paths.discard(rel_path)
        paths.add(_CONTENT_TYPES_PART)
        return paths

    def state_for(self, path: str) -> str:
        if path in self.edits["upsert"]:
            return "pending_upsert"
        if path == _CONTENT_TYPES_PART and any(self.edits["content_types"].values()):
            return "pending_generated"
        if _is_relationship_part(path):
            source = _source_for_relationship_part(path)
            if source in self.edits["relationships"]:
                return "pending_generated"
        return "source"

    def _read_source(self, path: str, limit: int | None = None) -> bytes:
        info = self.source_infos.get(path)
        if info is None:
            raise PackageToolError(f"Package part not found: {path!r}.")
        if limit is not None and info.file_size > limit:
            raise PackageToolError(f"Package part {path!r} exceeds the {limit}-byte inspection limit.")
        with zipfile.ZipFile(self.source_path, "r") as archive:
            return archive.read(path)

    def content_types(self) -> tuple[dict[str, str], dict[str, str]]:
        if self._content_types_cache is not None:
            return self._content_types_cache
        defaults, overrides = _parse_content_types(
            self._read_source(_CONTENT_TYPES_PART, _MAX_INFRASTRUCTURE_BYTES)
        )
        defaults = dict(defaults)
        overrides = dict(overrides)
        for extension, content_type in self.edits["content_types"]["defaults"].items():
            if content_type is None:
                defaults.pop(extension, None)
            else:
                defaults[extension] = content_type
        for part_path, content_type in self.edits["content_types"]["overrides"].items():
            if content_type is None:
                overrides.pop(part_path, None)
            else:
                overrides[part_path] = content_type
        self._content_types_cache = defaults, overrides
        return self._content_types_cache

    def _generated_content_types(self) -> bytes:
        defaults, overrides = self.content_types()
        return _serialize_content_types(defaults, overrides)

    def _generated_relationships(self, source: str) -> bytes:
        return _serialize_relationships(self.edits["relationships"][source])

    def full_bytes(self, path: str, limit: int | None = None) -> bytes:
        normalized = _normalize_part_path(path)
        if normalized not in self.part_paths():
            raise PackageToolError(f"Package part not found: {normalized!r}.")
        if normalized in self.edits["upsert"]:
            content = _decode_canonical_upsert(normalized, self.edits["upsert"][normalized])
        elif normalized == _CONTENT_TYPES_PART and any(self.edits["content_types"].values()):
            content = self._generated_content_types()
        elif _is_relationship_part(normalized):
            source = _source_for_relationship_part(normalized)
            if source in self.edits["relationships"]:
                content = self._generated_relationships(source)
            else:
                content = self._read_source(normalized, limit)
        else:
            content = self._read_source(normalized, limit)
        if limit is not None and len(content) > limit:
            raise PackageToolError(f"Package part {normalized!r} exceeds the {limit}-byte inspection limit.")
        return content

    def size(self, path: str) -> int:
        normalized = _normalize_part_path(path)
        if normalized not in self.part_paths():
            raise PackageToolError(f"Package part not found: {normalized!r}.")
        if normalized in self.edits["upsert"]:
            return int(self.edits["upsert"][normalized]["size"])
        if normalized == _CONTENT_TYPES_PART and any(self.edits["content_types"].values()):
            return len(self._generated_content_types())
        if _is_relationship_part(normalized):
            source = _source_for_relationship_part(normalized)
            if source in self.edits["relationships"]:
                return len(self._generated_relationships(source))
        return self.source_infos[normalized].file_size

    def iter_bytes(self, path: str, chunk_size: int = 1024 * 1024) -> Iterable[bytes]:
        normalized = _normalize_part_path(path)
        if normalized not in self.part_paths():
            raise PackageToolError(f"Package part not found: {normalized!r}.")
        generated = (
            normalized in self.edits["upsert"]
            or (normalized == _CONTENT_TYPES_PART and any(self.edits["content_types"].values()))
            or (
                _is_relationship_part(normalized)
                and _source_for_relationship_part(normalized) in self.edits["relationships"]
            )
        )
        if generated:
            content = self.full_bytes(normalized)
            for offset in range(0, len(content), chunk_size):
                yield content[offset : offset + chunk_size]
            return
        with zipfile.ZipFile(self.source_path, "r") as archive:
            with archive.open(normalized, "r") as stream:
                while True:
                    chunk = stream.read(chunk_size)
                    if not chunk:
                        break
                    yield chunk

    def sha256(self, path: str) -> str:
        digest = hashlib.sha256()
        for chunk in self.iter_bytes(path):
            digest.update(chunk)
        return digest.hexdigest()

    def read_range(self, path: str, offset: int, max_bytes: int) -> bytes:
        normalized = _normalize_part_path(path)
        size = self.size(normalized)
        if offset >= size:
            return b""
        generated = (
            normalized in self.edits["upsert"]
            or (normalized == _CONTENT_TYPES_PART and any(self.edits["content_types"].values()))
            or (
                _is_relationship_part(normalized)
                and _source_for_relationship_part(normalized) in self.edits["relationships"]
            )
        )
        if generated:
            return self.full_bytes(normalized)[offset : offset + max_bytes]
        with zipfile.ZipFile(self.source_path, "r") as archive:
            with archive.open(normalized, "r") as stream:
                if offset:
                    stream.seek(offset)
                return stream.read(max_bytes)

    def content_type_for(self, path: str) -> str | None:
        normalized = _normalize_part_path(path)
        if normalized == _CONTENT_TYPES_PART:
            return _CONTENT_TYPES_MIME
        defaults, overrides = self.content_types()
        if normalized in overrides:
            return overrides[normalized]
        filename = posixpath.basename(normalized)
        extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        return defaults.get(extension)

    def relationships(self) -> dict[str, list[dict[str, str]]]:
        if self._relationships_cache is not None:
            return self._relationships_cache
        relationships: dict[str, list[dict[str, str]]] = {}
        for rel_path in sorted(path for path in self.part_paths() if _is_relationship_part(path)):
            source = _source_for_relationship_part(rel_path)
            items = _parse_relationships(
                self.full_bytes(rel_path, _MAX_INFRASTRUCTURE_BYTES), rel_path
            )
            public_items = []
            for item in items:
                normalized = dict(item)
                normalized["source_part"] = source
                normalized["relationship_part"] = rel_path
                normalized["resolved_target"] = (
                    None
                    if item["target_mode"] == "External"
                    else _resolve_internal_target(source, item["target"])
                )
                public_items.append(normalized)
            relationships[source] = public_items
        self._relationships_cache = relationships
        return relationships


def _is_xml_content(path: str, content_type: str | None) -> bool:
    lowered = (content_type or "").lower()
    return path.endswith((".xml", ".rels")) or lowered.endswith("+xml") or "/xml" in lowered


def _validate_candidate(view: _PackageView) -> None:
    paths = view.part_paths()
    defaults, overrides = view.content_types()
    for override_path in overrides:
        if override_path not in paths:
            raise PackageToolError(
                f"Content-type override points to a missing package part: {override_path!r}."
            )
    for path in sorted(paths):
        if path == _CONTENT_TYPES_PART:
            continue
        content_type = view.content_type_for(path)
        if content_type is None:
            raise PackageToolError(f"Package part {path!r} has no effective content-type mapping.")
    for path in view.edits["upsert"]:
        content = view.full_bytes(path)
        if _is_xml_content(path, view.content_type_for(path)):
            _validate_xml_bytes(content, path)
    relationships = view.relationships()
    for source, items in relationships.items():
        if source != "/" and source not in paths:
            raise PackageToolError(f"Relationship source part does not exist: {source!r}.")
        seen_ids: set[str] = set()
        for item in items:
            if item["id"] in seen_ids:
                raise PackageToolError(f"Duplicate relationship ID {item['id']!r} for source {source!r}.")
            seen_ids.add(item["id"])
            if item["target_mode"] == "External":
                continue
            target = item["resolved_target"]
            if target not in paths:
                raise PackageToolError(
                    f"Dangling internal relationship {item['id']!r} from {source!r} to {target!r}."
                )


def _parse_upsert_operations(value: Any) -> tuple[list[tuple[str, bytes]], dict[str, str]]:
    operations = [] if value is None else _ensure_list(value, "upsert")
    if len(operations) > _MAX_OPERATIONS:
        raise PackageToolError("Too many package upsert operations.")
    result: list[tuple[str, bytes]] = []
    content_types: dict[str, str] = {}
    seen: set[str] = set()
    total_bytes = 0
    for index, raw_operation in enumerate(operations):
        operation = _ensure_mapping(raw_operation, f"upsert[{index}]")
        _reject_unknown_keys(
            operation,
            {"path", "content", "encoding", "content_type"},
            f"upsert[{index}]",
        )
        path = _normalize_part_path(operation.get("path"))
        if _is_reserved_structural_part(path):
            raise PackageToolError(
                f"Direct edits to {path!r} are prohibited; use structured relationships/content_types operations."
            )
        if path in seen:
            raise PackageToolError(f"Duplicate upsert operation for {path!r}.")
        seen.add(path)
        content = _decode_upsert_content(operation, f"upsert[{index}]")
        total_bytes += len(content)
        if total_bytes > _MAX_TRANSACTION_BYTES:
            raise PackageToolError("Package transaction exceeds the total byte limit.")
        if path.endswith(".xml"):
            _validate_xml_bytes(content, path)
        if operation.get("content_type") is not None:
            content_types[path] = _validate_content_type(operation["content_type"])
        result.append((path, content))
    return result, content_types


def _parse_delete_operations(value: Any) -> list[str]:
    operations = [] if value is None else _ensure_list(value, "delete")
    if len(operations) > _MAX_OPERATIONS:
        raise PackageToolError("Too many package delete operations.")
    result = []
    seen: set[str] = set()
    for index, raw_path in enumerate(operations):
        path = _normalize_part_path(raw_path)
        if _is_reserved_structural_part(path):
            raise PackageToolError(
                f"Direct deletion of {path!r} is prohibited; use structured relationships/content_types operations."
            )
        if path in seen:
            raise PackageToolError(f"Duplicate delete operation for {path!r} at index {index}.")
        seen.add(path)
        result.append(path)
    return result


def _parse_relationship_operations(value: Any) -> dict[str, list[dict[str, str]]]:
    operations = [] if value is None else _ensure_list(value, "relationships")
    if len(operations) > _MAX_OPERATIONS:
        raise PackageToolError("Too many relationship set operations.")
    result: dict[str, list[dict[str, str]]] = {}
    for index, raw_operation in enumerate(operations):
        operation = _ensure_mapping(raw_operation, f"relationships[{index}]")
        _reject_unknown_keys(operation, {"source_part", "relationships"}, f"relationships[{index}]")
        source = _normalize_part_path(operation.get("source_part"), allow_root=True)
        if source in result:
            raise PackageToolError(f"Duplicate relationship set for source {source!r}.")
        items = _ensure_list(operation.get("relationships"), f"relationships[{index}].relationships")
        if len(items) > _MAX_OPERATIONS:
            raise PackageToolError(f"Too many relationships for source {source!r}.")
        normalized_items = []
        seen_ids: set[str] = set()
        for item_index, raw_item in enumerate(items):
            item = _normalize_relationship_item(
                _ensure_mapping(raw_item, f"relationships[{index}].relationships[{item_index}]"),
                f"relationships[{index}].relationships[{item_index}]",
            )
            if item["id"] in seen_ids:
                raise PackageToolError(f"Duplicate relationship ID {item['id']!r} for source {source!r}.")
            seen_ids.add(item["id"])
            normalized_items.append(item)
        result[source] = normalized_items
    return result


def _parse_content_type_operations(value: Any) -> tuple[dict[str, str | None], dict[str, str | None]]:
    if value is None:
        return {}, {}
    operation = _ensure_mapping(value, "content_types")
    _reject_unknown_keys(operation, {"defaults", "overrides"}, "content_types")
    raw_defaults = _ensure_list(operation.get("defaults", []), "content_types.defaults")
    raw_overrides = _ensure_list(operation.get("overrides", []), "content_types.overrides")
    if len(raw_defaults) + len(raw_overrides) > _MAX_OPERATIONS:
        raise PackageToolError("Too many content-type operations.")
    defaults: dict[str, str | None] = {}
    overrides: dict[str, str | None] = {}
    for index, raw_item in enumerate(raw_defaults):
        item = _ensure_mapping(raw_item, f"content_types.defaults[{index}]")
        _reject_unknown_keys(item, {"extension", "content_type"}, f"content_types.defaults[{index}]")
        extension = _validate_extension(item.get("extension"))
        if extension in defaults:
            raise PackageToolError(f"Duplicate content-type Default mapping for {extension!r}.")
        raw_content_type = item.get("content_type")
        defaults[extension] = None if raw_content_type is None else _validate_content_type(raw_content_type)
    for index, raw_item in enumerate(raw_overrides):
        item = _ensure_mapping(raw_item, f"content_types.overrides[{index}]")
        _reject_unknown_keys(item, {"part_path", "content_type"}, f"content_types.overrides[{index}]")
        path = _normalize_part_path(item.get("part_path"))
        if path == _CONTENT_TYPES_PART:
            raise PackageToolError("[Content_Types].xml cannot have an Override mapping.")
        if path in overrides:
            raise PackageToolError(f"Duplicate content-type Override mapping for {path!r}.")
        raw_content_type = item.get("content_type")
        overrides[path] = None if raw_content_type is None else _validate_content_type(raw_content_type)
    return defaults, overrides


def _dirty_summary(edits: dict[str, Any]) -> dict[str, Any]:
    relationship_count = sum(len(items) for items in edits["relationships"].values())
    upsert_bytes = sum(int(entry["size"]) for entry in edits["upsert"].values())
    return {
        "upsert_count": len(edits["upsert"]),
        "upsert_bytes": upsert_bytes,
        "delete_count": len(edits["delete"]),
        "relationship_source_count": len(edits["relationships"]),
        "relationship_count": relationship_count,
        "content_type_default_edits": len(edits["content_types"]["defaults"]),
        "content_type_override_edits": len(edits["content_types"]["overrides"]),
        "dirty": bool(
            edits["upsert"]
            or edits["delete"]
            or edits["relationships"]
            or edits["content_types"]["defaults"]
            or edits["content_types"]["overrides"]
        ),
    }


def _relationship_indexes(
    view: _PackageView,
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, list[dict[str, Any]]]]:
    outbound = view.relationships()
    inbound: dict[str, list[dict[str, Any]]] = {}
    for items in outbound.values():
        for item in items:
            target = item.get("resolved_target")
            if target is not None:
                inbound.setdefault(target, []).append(item)
    return outbound, inbound


def _bounded_relationships(items: list[dict[str, Any]], limit: int) -> tuple[list[dict[str, Any]], bool]:
    public = [dict(item) for item in items[:limit]]
    return public, len(items) > limit


def _part_summary(
    view: _PackageView,
    path: str,
    outbound: dict[str, list[dict[str, Any]]],
    inbound: dict[str, list[dict[str, Any]]],
    relationship_limit: int,
) -> dict[str, Any]:
    outbound_items, outbound_truncated = _bounded_relationships(outbound.get(path, []), relationship_limit)
    inbound_items, inbound_truncated = _bounded_relationships(inbound.get(path, []), relationship_limit)
    return {
        "path": path,
        "state": view.state_for(path),
        "size": view.size(path),
        "sha256": view.sha256(path),
        "content_type": view.content_type_for(path),
        "outbound_relationship_count": len(outbound.get(path, [])),
        "outbound_relationships": outbound_items,
        "outbound_relationships_truncated": outbound_truncated,
        "inbound_relationship_count": len(inbound.get(path, [])),
        "inbound_relationships": inbound_items,
        "inbound_relationships_truncated": inbound_truncated,
    }


def _view_summary(view: _PackageView) -> dict[str, Any]:
    defaults, overrides = view.content_types()
    relationships = view.relationships()
    return {
        "part_count": len(view.part_paths()),
        "relationship_source_count": len(relationships),
        "relationship_count": sum(len(items) for items in relationships.values()),
        "content_type_default_count": len(defaults),
        "content_type_override_count": len(overrides),
        "pending": _dirty_summary(view.edits),
    }


def _change_report(
    before: _PackageView,
    after: _PackageView,
    impacted_paths: set[str],
    max_summary_items: int,
) -> dict[str, Any]:
    before_outbound, before_inbound = _relationship_indexes(before)
    after_outbound, after_inbound = _relationship_indexes(after)
    changes = []
    for path in sorted(impacted_paths):
        before_exists = path in before.part_paths()
        after_exists = path in after.part_paths()
        before_item = (
            _part_summary(before, path, before_outbound, before_inbound, 20)
            if before_exists
            else None
        )
        after_item = (
            _part_summary(after, path, after_outbound, after_inbound, 20)
            if after_exists
            else None
        )
        if before_item == after_item:
            continue
        if not before_exists:
            change = "added"
        elif not after_exists:
            change = "deleted"
        else:
            change = "modified"
        changes.append({"path": path, "change": change, "before": before_item, "after": after_item})
    return {
        "mutated": bool(changes or before.edits != after.edits),
        "before": _view_summary(before),
        "after": _view_summary(after),
        "changes": changes[:max_summary_items],
        "change_count": len(changes),
        "truncated": len(changes) > max_summary_items,
        "dirty": _dirty_summary(after.edits),
    }


def package_edit_verifier_patterns(
    data: dict,
    baseline_path: str | Path | None = None,
) -> list[str]:
    """Return granular verifier paths for effective pending package changes."""
    edits = _normalize_pending_edits(data.get("_package_edits"))
    if not _dirty_summary(edits)["dirty"]:
        return []

    baseline_data = dict(data)
    if baseline_path is not None:
        baseline_data["source"] = str(Path(baseline_path).expanduser().resolve())
    baseline = _PackageView(baseline_data, _empty_edits())
    effective = _PackageView(data, edits)
    baseline_parts = baseline.part_paths()
    effective_parts = effective.part_paths()
    patterns: list[str] = []

    def add(*values: str) -> None:
        for value in values:
            if value and value not in patterns:
                patterns.append(value)

    direct_paths = set(edits["upsert"]) | set(edits["delete"])
    for path in sorted(direct_paths):
        baseline_exists = path in baseline_parts
        effective_exists = path in effective_parts
        if baseline_exists != effective_exists or (
            baseline_exists
            and effective_exists
            and baseline.sha256(path) != effective.sha256(path)
        ):
            add(f"package/{path}")

    if edits["relationships"]:
        baseline_relationships = baseline.relationships()
        effective_relationships = effective.relationships()
        for source_part in sorted(edits["relationships"]):
            relationship_part = _relationship_part_for_source(source_part)
            baseline_exists = relationship_part in baseline_parts
            effective_exists = relationship_part in effective_parts
            relationships_changed = (
                baseline_relationships.get(source_part, [])
                != effective_relationships.get(source_part, [])
            )
            if baseline_exists != effective_exists or relationships_changed:
                add(f"package/{relationship_part}")
            if relationships_changed:
                add(f"package/relationships/{relationship_part}#*")

    baseline_defaults, baseline_overrides = baseline.content_types()
    effective_defaults, effective_overrides = effective.content_types()
    for extension in sorted(edits["content_types"]["defaults"]):
        if baseline_defaults.get(extension) != effective_defaults.get(extension):
            add(f"package/content_types/Default:{extension}")
    for path in sorted(edits["content_types"]["overrides"]):
        if baseline_overrides.get(path) != effective_overrides.get(path):
            add(f"package/content_types/Override:/{path}")
    return patterns


def _validate_summary_limit(value: int) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or not 1 <= value <= 200:
        raise PackageToolError("max_summary_items must be an integer from 1 to 200.")
    return value


def list_package_parts(
    data: dict,
    prefix: str | None = None,
    max_parts: int = 200,
    max_relationships_per_part: int = 20,
    include_deleted: bool = False,
) -> dict[str, Any]:
    """List effective package parts with hashes, content types and relationship edges."""
    if not isinstance(max_parts, int) or isinstance(max_parts, bool) or not 1 <= max_parts <= _MAX_LIST_PARTS:
        raise PackageToolError(f"max_parts must be an integer from 1 to {_MAX_LIST_PARTS}.")
    if (
        not isinstance(max_relationships_per_part, int)
        or isinstance(max_relationships_per_part, bool)
        or not 0 <= max_relationships_per_part <= _MAX_RELATIONSHIPS_PER_PART
    ):
        raise PackageToolError(
            f"max_relationships_per_part must be from 0 to {_MAX_RELATIONSHIPS_PER_PART}."
        )
    normalized_prefix = _normalize_prefix(prefix)
    edits = _normalize_pending_edits(data.get("_package_edits"))
    view = _PackageView(data, edits)
    _validate_candidate(view)
    outbound, inbound = _relationship_indexes(view)
    paths = sorted(path for path in view.part_paths() if path.startswith(normalized_prefix))
    visible_paths = paths[:max_parts]
    parts = [
        _part_summary(view, path, outbound, inbound, max_relationships_per_part)
        for path in visible_paths
    ]
    if include_deleted:
        for path in edits["delete"]:
            if path.startswith(normalized_prefix) and len(parts) < max_parts:
                parts.append({"path": path, "state": "pending_delete"})
    defaults, overrides = view.content_types()
    return {
        "source": str(view.source_path),
        "prefix": normalized_prefix,
        "part_count": len(paths),
        "returned_part_count": len(parts),
        "truncated": len(paths) > len(visible_paths),
        "content_type_default_count": len(defaults),
        "content_type_override_count": len(overrides),
        "dirty": _dirty_summary(edits),
        "parts": parts,
    }


def read_package_part(
    data: dict,
    part_path: str,
    output_mode: str = "auto",
    offset: int = 0,
    max_bytes: int = 65536,
) -> dict[str, Any]:
    """Read a bounded effective package part as XML, text, raw bytes or base64."""
    path = _normalize_part_path(part_path)
    if not isinstance(offset, int) or isinstance(offset, bool) or offset < 0:
        raise PackageToolError("offset must be a non-negative integer byte offset.")
    if not isinstance(max_bytes, int) or isinstance(max_bytes, bool) or not 1 <= max_bytes <= _MAX_READ_BYTES:
        raise PackageToolError(f"max_bytes must be an integer from 1 to {_MAX_READ_BYTES}.")
    mode = str(output_mode).strip().lower()
    if mode not in {"auto", "xml", "text", "raw", "base64"}:
        raise PackageToolError("output_mode must be auto, xml, text, raw, or base64.")
    edits = _normalize_pending_edits(data.get("_package_edits"))
    view = _PackageView(data, edits)
    _validate_candidate(view)
    content_type = view.content_type_for(path)
    if mode == "auto":
        if _is_xml_content(path, content_type):
            mode = "xml"
        elif (content_type or "").lower().startswith("text/") or any(
            token in (content_type or "").lower() for token in ("json", "javascript", "csv")
        ):
            mode = "text"
        else:
            guessed, _ = mimetypes.guess_type(path)
            mode = "text" if guessed and guessed.startswith("text/") else "base64"
    if mode == "xml":
        _validate_xml_bytes(view.full_bytes(path, _MAX_PART_BYTES), path)
    chunk = view.read_range(path, offset, max_bytes)
    size = view.size(path)
    if mode in {"xml", "text"}:
        content: Any = chunk.decode("utf-8", errors="replace")
        encoding = "utf-8"
    elif mode == "raw":
        content = list(chunk)
        encoding = "byte-array"
    else:
        content = base64.b64encode(chunk).decode("ascii")
        encoding = "base64"
    outbound, inbound = _relationship_indexes(view)
    outbound_items, outbound_truncated = _bounded_relationships(outbound.get(path, []), 50)
    inbound_items, inbound_truncated = _bounded_relationships(inbound.get(path, []), 50)
    return {
        "path": path,
        "state": view.state_for(path),
        "size": size,
        "sha256": view.sha256(path),
        "content_type": content_type,
        "output_mode": mode,
        "encoding": encoding,
        "offset": offset,
        "returned_bytes": len(chunk),
        "truncated": offset + len(chunk) < size,
        "content": content,
        "outbound_relationship_count": len(outbound.get(path, [])),
        "outbound_relationships": outbound_items,
        "outbound_relationships_truncated": outbound_truncated,
        "inbound_relationship_count": len(inbound.get(path, [])),
        "inbound_relationships": inbound_items,
        "inbound_relationships_truncated": inbound_truncated,
    }


def apply_package_transaction(
    data: dict,
    upsert: list[dict] | None = None,
    delete: list[str] | None = None,
    relationships: list[dict] | None = None,
    content_types: dict | None = None,
    max_summary_items: int = 100,
) -> dict[str, Any]:
    """Validate and stage one atomic package edit transaction in the session."""
    summary_limit = _validate_summary_limit(max_summary_items)
    parsed_upserts, upsert_content_types = _parse_upsert_operations(upsert)
    parsed_deletes = _parse_delete_operations(delete)
    parsed_relationships = _parse_relationship_operations(relationships)
    default_updates, override_updates = _parse_content_type_operations(content_types)
    if not any((parsed_upserts, parsed_deletes, parsed_relationships, default_updates, override_updates)):
        raise PackageToolError("Package transaction contains no operations.")
    upsert_paths = {path for path, _ in parsed_upserts}
    overlap = upsert_paths & set(parsed_deletes)
    if overlap:
        raise PackageToolError(f"Paths cannot be upserted and deleted in one transaction: {sorted(overlap)}.")
    duplicate_content_type_paths = set(upsert_content_types) & set(override_updates)
    if duplicate_content_type_paths:
        raise PackageToolError(
            "Content type was specified twice for upsert paths: "
            f"{sorted(duplicate_content_type_paths)}."
        )

    current = _normalize_pending_edits(data.get("_package_edits"))
    before = _PackageView(data, current)
    _validate_candidate(before)
    candidate = copy.deepcopy(current)
    source_paths = set(before.source_infos)
    impacted_paths: set[str] = set()

    for path, content in parsed_upserts:
        candidate["upsert"][path] = _canonical_upsert_entry(content)
        candidate["delete"] = [item for item in candidate["delete"] if item != path]
        impacted_paths.add(path)
    for path, content_type in upsert_content_types.items():
        candidate["content_types"]["overrides"][path] = content_type
        impacted_paths.add(_CONTENT_TYPES_PART)
    for path in parsed_deletes:
        if path not in before.part_paths():
            raise PackageToolError(f"Cannot delete missing package part: {path!r}.")
        candidate["upsert"].pop(path, None)
        if path in source_paths and path not in candidate["delete"]:
            candidate["delete"].append(path)
        candidate["delete"] = sorted(set(candidate["delete"]))
        relationship_part = _relationship_part_for_source(path)
        if relationship_part in source_paths:
            candidate["relationships"][path] = []
        else:
            candidate["relationships"].pop(path, None)
        impacted_paths.update({path, relationship_part})
    for source, items in parsed_relationships.items():
        candidate["relationships"][source] = items
        impacted_paths.add(_relationship_part_for_source(source))
        impacted_paths.add(source)
        for item in items:
            if item["target_mode"] == "Internal":
                impacted_paths.add(_resolve_internal_target(source, item["target"]))
    for extension, content_type in default_updates.items():
        candidate["content_types"]["defaults"][extension] = content_type
        impacted_paths.add(_CONTENT_TYPES_PART)
    for part_path, content_type in override_updates.items():
        candidate["content_types"]["overrides"][part_path] = content_type
        impacted_paths.update({_CONTENT_TYPES_PART, part_path})

    candidate = _normalize_pending_edits(candidate)
    after = _PackageView(data, candidate)
    _validate_candidate(after)
    report = _change_report(before, after, impacted_paths, summary_limit)
    data["_package_edits"] = candidate
    return report


def upsert_package_part(
    data: dict,
    part_path: str,
    content: str,
    encoding: str = "text",
    content_type: str | None = None,
    max_summary_items: int = 100,
) -> dict[str, Any]:
    """Stage one package-part upsert after full package validation."""
    operation = {"path": part_path, "content": content, "encoding": encoding}
    if content_type is not None:
        operation["content_type"] = content_type
    return apply_package_transaction(
        data,
        upsert=[operation],
        max_summary_items=max_summary_items,
    )


def delete_package_part(
    data: dict,
    part_path: str,
    max_summary_items: int = 100,
) -> dict[str, Any]:
    """Stage deletion of one part and remove its explicit content-type override."""
    path = _normalize_part_path(part_path)
    current = _normalize_pending_edits(data.get("_package_edits"))
    view = _PackageView(data, current)
    _validate_candidate(view)
    _, overrides = view.content_types()
    content_types = None
    if path in overrides:
        content_types = {"overrides": [{"part_path": path, "content_type": None}]}
    return apply_package_transaction(
        data,
        delete=[path],
        content_types=content_types,
        max_summary_items=max_summary_items,
    )


def set_package_relationships(
    data: dict,
    source_part: str,
    relationships: list[dict],
    max_summary_items: int = 100,
) -> dict[str, Any]:
    """Replace all relationships for one source part as an atomic session edit."""
    return apply_package_transaction(
        data,
        relationships=[{"source_part": source_part, "relationships": relationships}],
        max_summary_items=max_summary_items,
    )


def set_package_content_types(
    data: dict,
    defaults: list[dict] | None = None,
    overrides: list[dict] | None = None,
    max_summary_items: int = 100,
) -> dict[str, Any]:
    """Stage content-type mapping changes as an atomic session edit."""
    return apply_package_transaction(
        data,
        content_types={"defaults": defaults or [], "overrides": overrides or []},
        max_summary_items=max_summary_items,
    )


def register_package_tools(
    mcp: Any,
    get_session: Callable[[str], dict],
) -> dict[str, Callable[..., str]]:
    """Register package tools without importing the Excel server module."""
    registered: dict[str, Callable[..., str]] = {}

    def expose(function: Callable[..., str]) -> None:
        mcp.tool()(function)
        registered[function.__name__] = function

    def excel_list_package_parts(
        session_key: str,
        prefix: str | None = None,
        max_parts: int = 200,
        max_relationships_per_part: int = 20,
        include_deleted: bool = False,
    ) -> str:
        """List effective OOXML package parts with hashes, content types and relationship edges."""
        return json.dumps(
            list_package_parts(
                get_session(session_key),
                prefix=prefix,
                max_parts=max_parts,
                max_relationships_per_part=max_relationships_per_part,
                include_deleted=include_deleted,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_read_package_part(
        session_key: str,
        part_path: str,
        output_mode: str = "auto",
        offset: int = 0,
        max_bytes: int = 65536,
    ) -> str:
        """Read bounded XML, text, raw-byte-array or base64 content from an effective package part."""
        return json.dumps(
            read_package_part(
                get_session(session_key),
                part_path=part_path,
                output_mode=output_mode,
                offset=offset,
                max_bytes=max_bytes,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_upsert_package_part(
        session_key: str,
        part_path: str,
        content: str,
        encoding: str = "text",
        content_type: str | None = None,
        max_summary_items: int = 100,
    ) -> str:
        """Stage one XML/text or base64 binary package part; save applies it later."""
        return json.dumps(
            upsert_package_part(
                get_session(session_key),
                part_path=part_path,
                content=content,
                encoding=encoding,
                content_type=content_type,
                max_summary_items=max_summary_items,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_delete_package_part(
        session_key: str,
        part_path: str,
        max_summary_items: int = 100,
    ) -> str:
        """Stage deletion of one package part after dangling-reference validation."""
        return json.dumps(
            delete_package_part(
                get_session(session_key),
                part_path=part_path,
                max_summary_items=max_summary_items,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_set_package_relationships(
        session_key: str,
        source_part: str,
        relationships: list[dict],
        max_summary_items: int = 100,
    ) -> str:
        """Replace all relationships for one source part transactionally."""
        return json.dumps(
            set_package_relationships(
                get_session(session_key),
                source_part=source_part,
                relationships=relationships,
                max_summary_items=max_summary_items,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_set_package_content_types(
        session_key: str,
        defaults: list[dict] | None = None,
        overrides: list[dict] | None = None,
        max_summary_items: int = 100,
    ) -> str:
        """Set or remove Default and Override content-type mappings transactionally."""
        return json.dumps(
            set_package_content_types(
                get_session(session_key),
                defaults=defaults,
                overrides=overrides,
                max_summary_items=max_summary_items,
            ),
            ensure_ascii=False,
            indent=2,
        )

    def excel_apply_package_transaction(
        session_key: str,
        upsert: list[dict] | None = None,
        delete: list[str] | None = None,
        relationships: list[dict] | None = None,
        content_types: dict | None = None,
        max_summary_items: int = 100,
    ) -> str:
        """Atomically validate and stage package parts, relationships and content types."""
        return json.dumps(
            apply_package_transaction(
                get_session(session_key),
                upsert=upsert,
                delete=delete,
                relationships=relationships,
                content_types=content_types,
                max_summary_items=max_summary_items,
            ),
            ensure_ascii=False,
            indent=2,
        )

    for function in (
        excel_list_package_parts,
        excel_read_package_part,
        excel_upsert_package_part,
        excel_delete_package_part,
        excel_set_package_relationships,
        excel_set_package_content_types,
        excel_apply_package_transaction,
    ):
        expose(function)
    return registered


__all__ = [
    "PackageToolError",
    "apply_package_transaction",
    "delete_package_part",
    "list_package_parts",
    "package_edit_verifier_patterns",
    "read_package_part",
    "register_package_tools",
    "set_package_content_types",
    "set_package_relationships",
    "upsert_package_part",
]
