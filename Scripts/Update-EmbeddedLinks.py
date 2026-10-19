"""Offline, backup-first remediation of authored SharePoint hyperlink targets."""

import argparse
import csv
import hashlib
import html
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict
from datetime import datetime
from urllib.parse import parse_qsl, quote, unquote, urlsplit, urlunsplit
import xml.etree.ElementTree as ET
import zipfile

HOST = "whallc.sharepoint.com"
ORIGIN = "https://" + HOST
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SUPPORTED = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".pdf"}


def tenant_url(value):
    value = value.strip()
    if re.match(r"^whallc\.sharepoint\.com(?=[/?#]|$)", value, re.I):
        value = "https://" + value
    parsed = urlsplit(value)
    return (parsed if parsed.scheme.lower() in {"http", "https"}
            and parsed.netloc.lower() == HOST else None)


def guid(value):
    value = unquote(value).strip("{}").replace("-", "").lower()
    value = re.sub(r"^[a-z](?=[0-9a-f]{32}$)", "", value)
    return value if re.fullmatch(r"[0-9a-f]{32}", value) else ""


def url_guid(parsed):
    params = {k.lower(): v for k, v in parse_qsl(parsed.query)}
    return guid(params.get("sourcedoc", "")) or guid(params.get("d", ""))


def file_path(parsed):
    params = {k.lower(): v for k, v in parse_qsl(parsed.query)}
    path = unquote(parsed.path)
    for name in ("id", "sourcedoc", "sourcedocument"):
        if params.get(name, "").startswith("/"):
            path = params[name]
            break
    path = re.sub(r"^/:[a-z]:/r/", "/", path, flags=re.I).rstrip("/")
    if "/_layouts/" in path.lower() or re.match(r"^/:[a-z]:/", path, re.I):
        return ""
    # Do not conflate list items/views distinguished by their query parameters.
    if path.lower().endswith(".aspx") or not re.search(r"\.[a-z0-9]{1,8}$", path, re.I):
        return ""
    return path


def exact_key(value):
    parsed = tenant_url(value)
    if not parsed:
        return ""
    # Query values can contain case-sensitive sharing tokens; don't fold them.
    path = unquote(parsed.path).rstrip("/")
    if not re.match(r"^/:[a-z]:/[gs]/", path, re.I):
        path = path.lower()
    return urlunsplit(("https", HOST, path,
                       "&".join(sorted(parsed.query.split("&"))) if parsed.query else "", ""))


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        yield from csv.DictReader(stream)


class Resolver:
    """Only use destinations recorded as found; conflicts are never guessed."""

    def __init__(self, root, allow_view_filter_drop=False):
        self.index = defaultdict(dict)
        self.blocked = defaultdict(set)
        self.new_guids = {}
        self.new_paths = set()
        self.new_urls = set()
        self.allow_view_filter_drop = allow_view_filter_drop
        self.sources = []
        master_rows = []
        for folder in ("Masters", "FixerFiles"):
            for path in sorted((root / folder).glob("*.csv")):
                self.sources.append(path)
                for number, row in enumerate(read_csv(path), 2):
                    master_rows.append((row, f"{path.relative_to(root)}:{number}"))
                    if row.get("NewFound", "").lower() != "yes":
                        continue
                    dest = self.path_destination(row.get("NewFileRef", ""))
                    if dest:
                        self.new_paths.add(unquote(urlsplit(dest).path).lower())
                        for column in ("NewCopyLinkUrl", "NewResolvedUrl"):
                            parsed = tenant_url(row.get(column, ""))
                            if parsed and url_guid(parsed):
                                key = url_guid(parsed)
                                self.new_guids.setdefault(key, set()).add(dest)
        for row, source in master_rows:
            old_values = [row.get(c, "") for c in ("OldCopyLinkUrl", "OldResolvedUrl")]
            if row.get("OldFileRef", "").startswith("/"):
                old_values.append(ORIGIN + row["OldFileRef"])
            if row.get("NewFound", "").lower() == "yes":
                dest = (self.path_destination(row.get("NewFileRef", ""))
                        or self.destination(row.get("NewCopyLinkUrl", ""))
                        or self.destination(row.get("NewResolvedUrl", "")))
                for old in old_values:
                    self.add(old, dest, source)
            else:
                for old in old_values:
                    for key in self.keys(old):
                        self.blocked[key].add(source)
        for path in sorted((root / "EmbeddedLinks").glob("*.csv")):
            self.sources.append(path)
            for number, row in enumerate(read_csv(path), 2):
                old = row.get("EmbeddedURL", "")
                source = f"{path.relative_to(root)}:{number}"
                status = row.get("EmbeddedNewFound", "").lower()
                if status == "yes":
                    self.add(old, self.destination(row.get("EmbeddedNewURL", "")), source)
                elif status == "no":
                    for key in self.keys(old):
                        self.blocked[key].add(source)
        # Older first-pass tables can fill gaps, but cannot override a later
        # master/fixer/scan result, including an explicit "not found".
        for path in sorted((root / "FirstPassDocURLLookupTables").glob("*.csv")):
            self.sources.append(path)
            for number, row in enumerate(read_csv(path), 2):
                if row.get("NewFound", "").lower() != "yes":
                    continue
                olds = [row.get(c, "") for c in ("OldCopyLinkUrl", "OldResolvedUrl")]
                if row.get("OldFileRef", "").startswith("/"):
                    olds.append(ORIGIN + row["OldFileRef"])
                keys = {key for old in olds for key in self.keys(old)}
                if any(key in self.index or key in self.blocked for key in keys):
                    continue
                dest = (self.path_destination(row.get("NewFileRef", ""))
                        or self.destination(row.get("NewCopyLinkUrl", "")))
                for old in olds:
                    self.add(old, dest, f"{path.relative_to(root)}:{number}")
        if not self.sources:
            raise ValueError("No lookup CSVs found under the supplied root.")

    @staticmethod
    def path_destination(path):
        if not path.startswith("/") or path.startswith("//"):
            return ""
        return ORIGIN + quote(unquote(path), safe="/:@-._~!$&'()*+,;=")

    def destination(self, value):
        parsed = tenant_url(value)
        if not parsed:
            return ""
        path = file_path(parsed)
        if path:
            return self.path_destination(path)
        document_guid = url_guid(parsed)
        if document_guid:
            candidates = self.new_guids.get(document_guid, set())
            if len(candidates) == 1:
                return next(iter(candidates))
        return urlunsplit(("https", HOST, parsed.path, parsed.query, parsed.fragment))

    @staticmethod
    def keys(value):
        parsed = tenant_url(value)
        if not parsed:
            return []
        keys = ["exact:" + exact_key(value)]
        if re.match(r"^/:[a-z]:/[gs]/", parsed.path, re.I):
            keys.append("sharing:" + parsed.path)
        document_guid = url_guid(parsed)
        if document_guid:
            keys.append("guid:" + document_guid)
        path = file_path(parsed)
        if path:
            keys.append("path:" + path.lower())
        return keys

    def add(self, old, dest, source):
        if not dest or not tenant_url(dest):
            return
        original = tenant_url(old)
        if (original and not file_path(original) and not url_guid(original)
                and original.path.lower().endswith(".aspx")):
            # Some scan exports confirm the ASPX file but omit its list-item,
            # filtered-view, or other query context. That is not an equivalent
            # hyperlink destination; never silently discard that context.
            old_params = dict(parse_qsl(original.query))
            new_params = dict(parse_qsl(tenant_url(dest).query))
            changed_keys = [key for key, value in old_params.items() if new_params.get(key) != value]
            filter_only = all(re.fullmatch(
                r"filter(?:field|value|type|display)\d+|viewid|view|sortfield|sortdir", key, re.I)
                for key in changed_keys)
            if changed_keys and not (self.allow_view_filter_drop and filter_only):
                for key in self.keys(old):
                    self.blocked[key].add(source + " (destination omits or changes page/view query context)")
                return
        path = file_path(tenant_url(dest))
        if path:
            self.new_paths.add(path.lower())
        self.new_urls.add(exact_key(dest))
        for key in self.keys(old):
            self.index[key].setdefault(dest, set()).add(source)

    def resolve(self, value):
        parsed = tenant_url(value)
        if not parsed:
            return value, "external", ""
        keys = self.keys(value)
        # Combine all applicable evidence, rather than letting an exact match
        # hide a conflicting GUID/path mapping from another table.
        candidates = defaultdict(set)
        for key in keys:
            for dest, sources in self.index.get(key, {}).items():
                candidates[dest].update(sources)
        if len(candidates) > 1 and file_path(parsed):
            files = [dest for dest in candidates if file_path(tenant_url(dest))]
            if len(files) == 1:
                specific = files[0]
                document_path = unquote(urlsplit(specific).path).lower()
                views = [dest for dest in candidates if dest != specific]
                library_views = [re.fullmatch(r"(.*)/forms/[^/]+\.aspx",
                                              unquote(urlsplit(dest).path), re.I) for dest in views]
                if all(view and document_path.startswith(view[1].lower() + "/") for view in library_views):
                    sources = set(candidates[specific])
                    sources.add("Preferred confirmed document over less-specific view in the same library.")
                    candidates = {specific: sources}
        if len(candidates) > 1:
            evidence = "; ".join(f"{dest} [{', '.join(sorted(sources)[:3])}]"
                                 for dest, sources in sorted(candidates.items()))
            return value, "ambiguous", evidence
        if candidates:
            dest, sources = next(iter(candidates.items()))
            if parsed.fragment:
                parts = urlsplit(dest)
                dest = urlunsplit((*parts[:4], parsed.fragment))
            if exact_key(dest) == exact_key(value) and urlsplit(dest).fragment == parsed.fragment:
                return value, "already-current", "; ".join(sorted(sources)[:5])
            return dest, "resolved", "; ".join(sorted(sources)[:5])
        path = re.sub(r"^/:[a-z]:/r/", "/", unquote(parsed.path), flags=re.I)
        if path.lower().rstrip("/") in self.new_paths or exact_key(value) in self.new_urls:
            return value, "already-current", "Verified new destination in lookup tables."
        if path.lower().startswith("/sites/"):
            return value, "new-structure-unverified", "Already under /sites; no confirmed replacement."
        denied = set().union(*(self.blocked.get(key, set()) for key in keys))
        return value, "unresolved", ("Not found in supplied scan: " + "; ".join(sorted(denied)[:5])
                                     if denied else "No confirmed destination in supplied tables.")


def xml_text(value):
    return html.escape(value, quote=False).replace("\r", "&#13;")


class Editor:
    def __init__(self, resolver, relative):
        self.resolver = resolver
        self.relative = relative
        self.rows = []

    def target(self, old, location):
        new, status, evidence = self.resolver.resolve(old)
        if status != "external":
            self.rows.append(dict(File=str(self.relative), Location=location,
                                  OldURL=old, NewURL=new, Status=status, Evidence=evidence))
        return new

    def ooxml(self, path):
        original = path.read_bytes()
        changes = {}
        with zipfile.ZipFile(io.BytesIO(original)) as archive:
            if any(n.lower().startswith("_xmlsignatures/") for n in archive.namelist()):
                raise ValueError("Digitally signed Office package; editing would invalidate its signature.")
            for entry in archive.infolist():
                if entry.filename.endswith(".rels"):
                    data = archive.read(entry)
                    root = ET.fromstring(data)
                    replacements = {}
                    for rel in root:
                        if (rel.tag == f"{{{REL_NS}}}Relationship"
                                and rel.get("Type", "").endswith("/hyperlink")
                                and rel.get("TargetMode", "").lower() == "external"):
                            old = rel.get("Target", "")
                            new = self.target(old, f"{entry.filename}#{rel.get('Id')}")
                            if new != old:
                                replacements[rel.get("Id")] = (old, new)
                    if replacements:
                        text = data.decode("utf-8-sig")

                        def replace_relationship(match):
                            tag = match.group()
                            attrs = {m.group(1): html.unescape(m.group(3)) for m in re.finditer(
                                r"""([\w:]+)\s*=\s*(["'])(.*?)\2""", tag, re.S)}
                            if attrs.get("Id") not in replacements:
                                return tag
                            old, new = replacements[attrs["Id"]]
                            if attrs.get("Target") != old:
                                raise ValueError("Relationship target changed during edit.")
                            return re.sub(r"""(\bTarget\s*=\s*)(["'])(.*?)\2""",
                                          lambda m: m[1] + m[2] + html.escape(new, quote=True) + m[2],
                                          tag, count=1, flags=re.S)

                        updated = re.sub(r"<(?:\w+:)?Relationship\b[^>]*>", replace_relationship, text)
                        changes[entry.filename] = self.xml_bytes(data, updated)
                elif entry.filename.endswith(".xml"):
                    data = archive.read(entry)
                    if b"HYPERLINK" in data.upper() or b"instrText" in data:
                        updated = self.edit_fields(data, entry.filename)
                        if updated != data:
                            changes[entry.filename] = updated
            if not changes:
                return None
            result = io.BytesIO()
            with zipfile.ZipFile(result, "w") as output:
                output.comment = archive.comment
                for entry in archive.infolist():
                    output.writestr(entry, changes.get(entry.filename, archive.read(entry)))
        # Confirm every member, including pictures/styles/macros, survived exactly.
        with zipfile.ZipFile(io.BytesIO(result.getvalue())) as output:
            if output.testzip():
                raise ValueError("Office ZIP integrity check failed.")
            with zipfile.ZipFile(io.BytesIO(original)) as before:
                if before.namelist() != output.namelist():
                    raise ValueError("Office package member list changed.")
                for name in before.namelist():
                    expected = changes.get(name, before.read(name))
                    if output.read(name) != expected:
                        raise ValueError(f"Office package verification failed: {name}")
                    if name.endswith(".xml"):
                        a, b = ET.fromstring(before.read(name)), ET.fromstring(output.read(name))
                        # All visible text and cached formula results must stay identical.
                        visible = {f"{{{W_NS}}}t", f"{{{S_NS}}}t", f"{{{S_NS}}}v",
                                   "{http://schemas.openxmlformats.org/drawingml/2006/main}t"}
                        if [(e.tag, e.text) for e in a.iter() if e.tag in visible] != [
                                (e.tag, e.text) for e in b.iter() if e.tag in visible]:
                            raise ValueError(f"Visible content changed: {name}")
        return result.getvalue()

    @staticmethod
    def xml_bytes(original, text):
        return (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode("utf-8")

    def edit_fields(self, data, location):
        text = data.decode("utf-8-sig")
        ET.fromstring(data)
        edits = []
        # Word simple fields keep the instruction in an attribute.
        for match in re.finditer(r"""<\w+:fldSimple\b[^>]*\b\w+:instr=(["'])(.*?)\1""", text, re.S):
            instruction = html.unescape(match[2])
            updated = self.field_instruction(instruction, location + "#simple-field")
            if updated != instruction:
                edits.append((match.start(2), match.end(2), html.escape(updated, quote=True)))
        # Complex fields can split their HYPERLINK instruction across many runs.
        stack = []
        token_pattern = (r"<(?P<p>\w+):fldChar\b[^>]*/?>|"
                         r"<(?P<q>\w+):instrText\b[^>]*>(?P<code>.*?)</(?P=q):instrText>")
        for match in re.finditer(token_pattern, text, re.S):
            if match["p"]:
                kind = re.search(r"""\bfldCharType=(["'])(.*?)\1""", match.group())
                if kind:
                    if kind[2] == "begin":
                        stack.append([])
                    elif kind[2] == "separate" and stack:
                        self.field_edits(stack[-1], text, location, edits)
                        stack[-1] = []
                    elif kind[2] == "end" and stack:
                        self.field_edits(stack.pop(), text, location, edits)
            elif stack:
                stack[-1].append((match.start("code"), match.end("code")))
        # Excel HYPERLINK formulas: only replace a literal first argument,
        # never the friendly-name argument, cached value, or a computed URL.
        for match in re.finditer(r"<(?:\w+:)?f\b[^>]*>(.*?)</(?:\w+:)?f>", text, re.S):
            formula = html.unescape(match[1])
            updated = re.sub(
                r'(?i)(?<![\w.])((?:_xlfn\.)?HYPERLINK\s*\(\s*)"((?:[^"]|"")*)"',
                lambda m: m[1] + '"' + self.target(
                    m[2].replace('""', '"'), location + "#HYPERLINK-formula").replace('"', '""') + '"',
                formula)
            if formula != updated:
                edits.append((match.start(1), match.end(1), xml_text(updated)))
        for start, end, replacement in sorted(edits, reverse=True):
            text = text[:start] + replacement + text[end:]
        ET.fromstring(text)
        return self.xml_bytes(data, text)

    def field_instruction(self, instruction, location):
        match = re.search(r'(?i)\bHYPERLINK\s+(?:"([^"]*)"|([^\s\\]+))', instruction)
        if not match:
            return instruction
        group = 1 if match[1] is not None else 2
        old = match[group]
        new = self.target(old, location)
        return instruction[:match.start(group)] + new + instruction[match.end(group):]

    def field_edits(self, spans, text, location, edits):
        values = [html.unescape(text[start:end]) for start, end in spans]
        instruction = "".join(values)
        updated = self.field_instruction(instruction, location + "#complex-field")
        if instruction == updated:
            return
        # Retain all run nodes and only distribute the changed URL over their
        # existing instruction text spans; displayed field results are untouched.
        match = re.search(r'(?i)\bHYPERLINK\s+(?:"([^"]*)"|([^\s\\]+))', instruction)
        group = 1 if match[1] is not None else 2
        a, b = match.span(group)
        replacement = updated[a:len(updated) - (len(instruction) - b)] if b < len(instruction) else updated[a:]
        offset = 0
        inserted = False
        for (start, end), value in zip(spans, values):
            lo, hi = max(a - offset, 0), min(b - offset, len(value))
            if lo < hi:
                result = value[:lo] + (replacement if not inserted else "") + value[hi:]
                edits.append((start, end, xml_text(result)))
                inserted = True
            offset += len(value)

    def pdf(self, path):
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject, TextStringObject

        original = path.read_bytes()
        reader = PdfReader(io.BytesIO(original), strict=True)
        if reader.is_encrypted:
            raise ValueError("Encrypted PDF requires an explicitly authorized decryption workflow.")
        fields = reader.get_fields() or {}
        if any(f.get("/FT") == "/Sig" for f in fields.values()):
            raise ValueError("PDF contains a signature field; editing could invalidate a signature.")
        # Incremental updates retain the original PDF bytes and append only
        # modified objects. This also avoids rewriting unrelated damaged objects.
        writer = PdfWriter(io.BytesIO(original), incremental=True)
        changed = False
        for page_number, page in enumerate(writer.pages, 1):
            annotations = page.get("/Annots", [])
            annotations = annotations.get_object() if hasattr(annotations, "get_object") else annotations
            for number, ref in enumerate(annotations, 1):
                annotation = ref.get_object()
                action = annotation.get("/A")
                if annotation.get("/Subtype") != "/Link" or not action:
                    continue
                action = action.get_object()
                if action.get("/S") == "/URI" and "/URI" in action:
                    old = str(action["/URI"])
                    new = self.target(old, f"PDF page {page_number}, annotation {number}")
                    if new != old:
                        action[NameObject("/URI")] = TextStringObject(new)
                        changed = True
        if not changed:
            return None
        result = io.BytesIO()
        writer.write(result)
        if not result.getvalue().startswith(original):
            raise ValueError("Incremental PDF update did not retain the original bytes.")
        checked = PdfReader(io.BytesIO(result.getvalue()), strict=True)
        if len(reader.pages) != len(checked.pages):
            raise ValueError("PDF page count changed.")
        for a, b in zip(reader.pages, checked.pages):
            a_stream = a.get_contents()
            b_stream = b.get_contents()
            if ((a_stream.get_data() if a_stream is not None else b"")
                    != (b_stream.get_data() if b_stream is not None else b"")):
                raise ValueError("PDF page drawing/text stream changed.")
            if a.extract_text() != b.extract_text():
                raise ValueError("PDF extracted visible text changed.")
            if list(a.mediabox) != list(b.mediabox) or list(a.cropbox) != list(b.cropbox):
                raise ValueError("PDF page dimensions changed.")
            a_links, b_links = a.get("/Annots", []), b.get("/Annots", [])
            a_links = a_links.get_object() if hasattr(a_links, "get_object") else a_links
            b_links = b_links.get_object() if hasattr(b_links, "get_object") else b_links
            if len(a_links) != len(b_links):
                raise ValueError("PDF annotation count changed.")
            for aa, bb in zip(a_links, b_links):
                aa, bb = aa.get_object(), bb.get_object()
                if aa.get("/Rect") != bb.get("/Rect"):
                    raise ValueError("PDF clickable region changed.")
                action = aa.get("/A")
                if aa.get("/Subtype") == "/Link" and action:
                    action = action.get_object()
                    if action.get("/S") == "/URI":
                        old = str(action.get("/URI", ""))
                        expected = self.resolver.resolve(old)[0]
                        actual = str(bb["/A"].get_object().get("/URI", ""))
                        if expected != actual:
                            raise ValueError("PDF hyperlink verification failed.")
        return result.getvalue()

    def legacy_word(self, path):
        helper = Path(__file__).with_name("Edit-LegacyWordLinks.ps1")
        with tempfile.TemporaryDirectory(prefix="legacy-linkfix-") as folder:
            folder = Path(folder)
            inspection = folder / "inspection.json"
            command = ["pwsh", "-NoProfile", "-NonInteractive", "-STA", "-File", str(helper),
                       "-InputPath", str(path.resolve())]
            try:
                subprocess.run(command + ["-OutputJson", str(inspection)], check=True, timeout=120,
                               capture_output=True, text=True, encoding="utf-8")
            except subprocess.CalledProcessError as exc:
                raise RuntimeError(f"Word inspection failed: {exc.stderr}") from exc
            snapshot = json.loads(inspection.read_text(encoding="utf-8-sig"))
            replacements = []
            for link in snapshot["Links"]:
                new = self.target(link["Address"], link["Location"])
                if new != link["Address"]:
                    replacements.append(dict(Location=link["Location"], OldURL=link["Address"], NewURL=new))
            if not replacements:
                return None
            mapping = folder / "replacements.json"
            mapping.write_text(json.dumps(replacements), encoding="utf-8")
            output = folder / "edited.doc"
            try:
                subprocess.run(command + ["-OutputJson", str(folder / "verified.json"),
                                           "-OutputPath", str(output), "-ReplacementsJson", str(mapping)],
                               check=True, timeout=120, capture_output=True, text=True, encoding="utf-8")
            except subprocess.CalledProcessError as exc:
                raise RuntimeError(f"Word editing failed: {exc.stderr}") from exc
            return output.read_bytes()


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_csv(path, rows, columns):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--run-dir", type=Path, required=True)
    parser.add_argument("--apply", action="store_true", help="Without this, only produce an audit.")
    parser.add_argument("--legacy-word", action="store_true",
                        help="Use locally installed desktop Word for .doc files (Windows only).")
    parser.add_argument("--allow-view-filter-drop", action="store_true",
                        help="Allow confirmed library/page destinations to drop view/filter/sort parameters.")
    args = parser.parse_args()
    root, input_dir, run_dir = args.root.resolve(), args.input.resolve(), args.run_dir.resolve()
    if not input_dir.is_dir():
        parser.error("Input must be an existing directory.")
    if run_dir == input_dir or input_dir in run_dir.parents:
        parser.error("Run directory/backups must be outside the input tree.")
    if run_dir.exists():
        parser.error("Use a new run directory; never overwrite an existing audit or backup.")
    try:
        import pypdf  # noqa: F401
    except ImportError:
        parser.error("Missing pypdf. Install Scripts\\embedded-link-requirements.txt first.")
    run_dir.mkdir(parents=True)
    resolver = Resolver(root, allow_view_filter_drop=args.allow_view_filter_drop)
    files, links = [], []
    for path in sorted(p for p in input_dir.rglob("*") if p.is_file()):
        relative = path.relative_to(input_dir)
        editor = Editor(resolver, relative)
        row = dict(File=str(relative), Status="", ChangedLinks=0, SHA256Before="",
                   SHA256After="", Backup="", Error="")
        temp = None
        try:
            row["SHA256Before"] = digest(path)
            if path.suffix.lower() == ".doc" and args.legacy_word:
                result = editor.legacy_word(path)
            elif path.suffix.lower() not in SUPPORTED:
                row["Status"] = "unsupported"
                row["Error"] = "Legacy DOC requires Word COM; no binary or format-conversion edits attempted."
            else:
                result = editor.pdf(path) if path.suffix.lower() == ".pdf" else editor.ooxml(path)
            if row["Status"] != "unsupported":
                count = sum(r["Status"] == "resolved" for r in editor.rows)
                row["ChangedLinks"] = count
                row["Status"] = ("updated" if args.apply else "would-update") if result else "unchanged"
                if result and args.apply:
                    backup = run_dir / "backups" / relative
                    backup.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(path, backup)
                    if digest(backup) != row["SHA256Before"] or digest(path) != row["SHA256Before"]:
                        raise ValueError("Source or backup changed before write.")
                    with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".linkfix-", delete=False) as stream:
                        temp = Path(stream.name)
                        stream.write(result)
                    os.replace(temp, path)
                    temp = None
                    row["Backup"] = str(backup)
                row["SHA256After"] = digest(path)
        except Exception as exc:
            row["Status"] = "error"
            row["Error"] = f"{type(exc).__name__}: {exc}"
            for link in editor.rows:
                if link["Status"] == "resolved":
                    link["Status"] = "not-applied-error"
            print(f"ERROR {relative}: {row['Error']}", file=sys.stderr)
        finally:
            if temp is not None:
                temp.unlink(missing_ok=True)
        if not args.apply:
            for link in editor.rows:
                if link["Status"] == "resolved":
                    link["Status"] = "would-update"
        files.append(row)
        links.extend(editor.rows)
        print(f"{row['Status']:14} {row['ChangedLinks']:4} {relative}", flush=True)
        # Persist progress after each document, even if a later document fails.
        write_csv(run_dir / "files.csv", files, list(row))
        write_csv(run_dir / "links.csv", links,
                  ["File", "Location", "OldURL", "NewURL", "Status", "Evidence"])
    manifest = dict(Input=str(input_dir), Applied=args.apply,
                    AllowViewFilterDrop=args.allow_view_filter_drop, Timestamp=datetime.now().isoformat(),
                    Files=len(files), FileStatuses=dict_count(files, "Status"),
                    LinkStatuses=dict_count(links, "Status"),
                    LookupSources=[dict(Path=str(p.relative_to(root)), SHA256=digest(p))
                                   for p in resolver.sources])
    review = {}
    for link in links:
        if link["Status"] not in {"unresolved", "ambiguous", "new-structure-unverified", "not-applied-error"}:
            continue
        key = (link["OldURL"], link["Status"])
        record = review.setdefault(key, dict(OldURL=link["OldURL"], Status=link["Status"], Occurrences=0,
                                              Files=set(), Evidence=link["Evidence"]))
        record["Occurrences"] += 1
        record["Files"].add(link["File"])
    review_rows = [dict(record, Files="; ".join(sorted(record["Files"]))) for record in review.values()]
    write_csv(run_dir / "review.csv", review_rows, ["OldURL", "Status", "Occurrences", "Files", "Evidence"])
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in manifest.items() if k != "LookupSources"}, indent=2))
    return 1 if any(r["Status"] == "error" for r in files) else 0


def dict_count(rows, key):
    result = defaultdict(int)
    for row in rows:
        result[row[key]] += 1
    return dict(result)


if __name__ == "__main__":
    sys.exit(main())
