import csv
import importlib.util
import io
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DictionaryObject, NameObject, TextStringObject, NumberObject

SPEC = importlib.util.spec_from_file_location("linkfix", Path(__file__).with_name("Update-EmbeddedLinks.py"))
fix = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(fix)
OLD = fix.ORIGIN + "/HR/Guide.docx"
NEW = fix.ORIGIN + "/sites/HR/Guide.docx"
OLD_GUID = "01234567-89ab-cdef-0123-456789abcdef"


class LinkTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / "Masters").mkdir()
        self.table = self.root / "Masters" / "HR-master.csv"
        self.write_rows(self.table, [
            dict(OldFileRef="/HR/Guide.docx", OldCopyLinkUrl=OLD,
                 OldResolvedUrl=fix.ORIGIN + "/HR/_layouts/15/Doc.aspx?sourcedoc={" + OLD_GUID + "}",
                 NewFileRef="/sites/HR/Guide.docx", NewFound="Yes")
        ])
        self.resolver = fix.Resolver(self.root)
        self.editor = fix.Editor(self.resolver, Path("fixture.docx"))

    @staticmethod
    def write_rows(path, rows):
        with path.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=sorted(set().union(*(r.keys() for r in rows))))
            writer.writeheader()
            writer.writerows(rows)

    def test_direct_bare_host_and_fragment(self):
        self.assertEqual(self.resolver.resolve(OLD + "#bookmark")[0], NEW + "#bookmark")
        self.assertEqual(self.resolver.resolve(OLD.replace("https://", ""))[0], NEW)

    def test_guid_and_view_url(self):
        url = fix.ORIGIN + "/:w:/r/HR/_layouts/15/Doc.aspx?sourcedoc=%7B" + OLD_GUID.upper() + "%7D"
        self.assertEqual(self.resolver.resolve(url)[0], NEW)
        url = fix.ORIGIN + "/HR/Documents/Forms/AllItems.aspx?id=%2FHR%2FGuide.docx&parent=%2FHR"
        self.assertEqual(self.resolver.resolve(url)[0], NEW)

    def test_external_and_similar_host_untouched(self):
        for url in ["https://example.com/" + OLD, "https://whallc.sharepoint.com.evil.test/HR/Guide.docx",
                    "https://user@whallc.sharepoint.com/HR/Guide.docx"]:
            self.assertEqual(self.resolver.resolve(url), (url, "external", ""))

    def test_missing_is_not_guessed(self):
        url = fix.ORIGIN + "/HR/Missing.docx"
        self.assertEqual(self.resolver.resolve(url)[:2], (url, "unresolved"))

    def test_confirmed_rename_drops_only_explicitly_approved_fragment(self):
        path = self.root / "confirmed.csv"
        old = fix.ORIGIN + "/HR/Old.pdf#search%3Dold"
        new = fix.ORIGIN + "/sites/HR/New.docx"
        self.write_rows(path, [dict(OldURL=old, NewURL=new, Evidence="User confirmed", DropFragment="Yes")])
        resolver = fix.Resolver(self.root, confirmed_mappings=path)
        self.assertEqual(resolver.resolve(old)[:2], (new, "resolved"))
        self.assertEqual(resolver.resolve(new)[1], "already-current")
        self.write_rows(path, [dict(OldURL=old, NewURL=new, Evidence="User confirmed", DropFragment="No")])
        resolver = fix.Resolver(self.root, confirmed_mappings=path)
        self.assertEqual(resolver.resolve(old)[0], new + "#search%3Dold")

    def test_confirmed_mapping_can_override_outdated_scan(self):
        path = self.root / "confirmed.csv"
        self.write_rows(path, [dict(OldURL=OLD, NewURL=fix.ORIGIN + "/sites/HR/Renamed.docx",
                                   Evidence="Live SharePoint metadata", DropFragment="No")])
        resolver = fix.Resolver(self.root, confirmed_mappings=path)
        self.assertEqual(resolver.resolve(OLD)[0], fix.ORIGIN + "/sites/HR/Renamed.docx")

    def test_confirmed_mapping_rejects_invalid_and_conflicting_rows(self):
        path = self.root / "confirmed.csv"
        self.write_rows(path, [dict(OldURL=OLD, NewURL="https://example.com",
                                   Evidence="invalid", DropFragment="No")])
        with self.assertRaisesRegex(ValueError, "Invalid confirmed"):
            fix.Resolver(self.root, confirmed_mappings=path)
        self.write_rows(path, [dict(OldURL=OLD, NewURL=NEW, Evidence="first", DropFragment="No"),
                               dict(OldURL=OLD, NewURL=fix.ORIGIN + "/sites/HR/Other.docx",
                                    Evidence="second", DropFragment="No")])
        with self.assertRaisesRegex(ValueError, "Conflicting confirmed"):
            fix.Resolver(self.root, confirmed_mappings=path)

    def test_current_is_idempotent(self):
        self.assertEqual(self.resolver.resolve(NEW)[:2], (NEW, "already-current"))

    def test_embedded_report_only_destination_is_idempotent(self):
        self.resolver.add(fix.ORIGIN + "/HR/Other.docx",
                          fix.ORIGIN + "/sites/HR/Other.docx", "embedded.csv:2")
        self.assertEqual(self.resolver.resolve(fix.ORIGIN + "/sites/HR/Other.docx")[1], "already-current")

    def test_conflicting_evidence_is_not_applied(self):
        self.resolver.add(OLD, fix.ORIGIN + "/sites/Other/Guide.docx", "conflict.csv:2")
        self.assertEqual(self.resolver.resolve(OLD)[:2], (OLD, "ambiguous"))

    def test_page_mapping_cannot_discard_filter_or_item_context(self):
        for query in ("?FilterField1=Audience&FilterValue1=Clinical&viewid=old-view",
                      "?ID=12", "?RootFolder=%2FHR%2FDocuments"):
            old = fix.ORIGIN + "/HR/Forms/View.aspx" + query
            dest = fix.ORIGIN + "/sites/HR/Forms/View.aspx"
            self.resolver.add(old, dest, "incomplete-report.csv:2")
            self.assertEqual(self.resolver.resolve(old)[1], "unresolved")
        old = fix.ORIGIN + "/HR/Forms/View.aspx?FilterField1=Audience&FilterValue1=Clinical"
        dest = fix.ORIGIN + "/sites/HR/Forms/View.aspx?FilterField1=Audience&FilterValue1=Clinical"
        self.resolver.add(old, dest, "complete-report.csv:2")
        self.assertEqual(self.resolver.resolve(old)[0], dest)

    def test_opt_in_can_drop_view_filters_but_not_item_or_folder_context(self):
        resolver = fix.Resolver(self.root, allow_view_filter_drop=True)
        old = fix.ORIGIN + "/HR/Forms/View.aspx?FilterField1=Audience&FilterValue1=Clinical&viewid=old"
        dest = fix.ORIGIN + "/sites/HR/Forms/View.aspx"
        resolver.add(old, dest, "report.csv:2")
        self.assertEqual(resolver.resolve(old)[0], dest)
        self.assertEqual(resolver.resolve(dest)[1], "already-current")
        for query in ("?ID=12", "?RootFolder=%2FHR%2FDocuments"):
            url = fix.ORIGIN + "/HR/Forms/View.aspx" + query
            resolver.add(url, dest, "report.csv:3")
            self.assertEqual(resolver.resolve(url)[1], "unresolved")

    def test_specific_file_beats_library_view_but_not_another_file(self):
        old = fix.ORIGIN + "/HR/Documents/Forms/AllItems.aspx?id=%2FHR%2FGuide.docx"
        view = fix.ORIGIN + "/sites/HR/Forms/AllItems.aspx"
        self.resolver.add(old, view, "generic-view.csv:2")
        self.assertEqual(self.resolver.resolve(old)[0], NEW)
        self.resolver.add(old, fix.ORIGIN + "/sites/Other/Forms/AllItems.aspx", "other-library.csv:2")
        self.assertEqual(self.resolver.resolve(old)[1], "ambiguous")

    def test_sharing_token_ignores_tracking_but_is_case_sensitive(self):
        url = fix.ORIGIN + "/:w:/g/HR/EAbCdEf"
        self.resolver.add(url + "?e=first", NEW, "sharing.csv:2")
        self.assertEqual(self.resolver.resolve(url + "?e=second")[0], NEW)
        self.assertEqual(self.resolver.resolve(url.lower() + "?e=second")[1], "unresolved")
        self.assertEqual(self.resolver.resolve(url.lower() + "?e=first")[1], "unresolved")

    def test_first_pass_cannot_override_newer_negative(self):
        self.write_rows(self.table, [dict(OldFileRef="/HR/Guide.docx", NewFound="No")])
        folder = self.root / "FirstPassDocURLLookupTables"
        folder.mkdir()
        self.write_rows(folder / "old.csv", [dict(OldFileRef="/HR/Guide.docx", NewFound="Yes",
                                                 NewFileRef="/sites/HR/Guide.docx")])
        resolver = fix.Resolver(self.root)
        self.assertEqual(resolver.resolve(OLD)[1], "unresolved")

    def test_relationship_targets_only_and_visible_text_identical(self):
        path = self.root / "test.docx"
        relationships = f'''<Relationships xmlns="{fix.REL_NS}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{OLD}" TargetMode="External"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{OLD}" TargetMode="External"/>
</Relationships>'''.encode()
        content = f'<w:document xmlns:w="{fix.W_NS}"><w:t>{OLD}</w:t></w:document>'.encode()
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/_rels/document.xml.rels", relationships)
            archive.writestr("word/document.xml", content)
            archive.writestr("word/media/image.png", b"untouched binary")
        result = self.editor.ooxml(path)
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            self.assertEqual(archive.read("word/document.xml"), content)
            rels = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
            self.assertEqual([r.get("Target") for r in rels], [NEW, "https://example.com", OLD])
            self.assertEqual(archive.read("word/media/image.png"), b"untouched binary")

    def test_word_split_instruction_keeps_field_result(self):
        xml = f'''<w:document xmlns:w="{fix.W_NS}"><w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> HYPER</w:instrText></w:r><w:r><w:instrText>LINK "https://whallc.sharepoint.</w:instrText></w:r>
<w:r><w:instrText>com/HR/Guide.docx" \\o "Keep tooltip" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Keep this label</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:document>'''.encode()
        result = self.editor.edit_fields(xml, "word/document.xml")
        root = ET.fromstring(result)
        instruction = "".join(e.text or "" for e in root.iter(f"{{{fix.W_NS}}}instrText"))
        self.assertIn(NEW, instruction)
        self.assertIn('Keep tooltip', instruction)
        self.assertIn(b"<w:t>Keep this label</w:t>", result)

    def test_word_simple_field_preserves_label(self):
        xml = (f'<w:document xmlns:w="{fix.W_NS}"><w:fldSimple w:instr=" HYPERLINK &quot;{OLD}&quot; ">'
               '<w:r><w:t>Keep label</w:t></w:r></w:fldSimple></w:document>').encode()
        result = self.editor.edit_fields(xml, "word/header.xml")
        self.assertIn(NEW.encode(), result)
        self.assertIn(b"<w:t>Keep label</w:t>", result)

    def test_excel_formula_keeps_display_argument_and_cached_value(self):
        xml = (f'<worksheet xmlns="{fix.S_NS}"><c><f>HYPERLINK("{OLD}","{OLD}")</f>'
               '<v>Cached Label</v></c></worksheet>').encode()
        result = self.editor.edit_fields(xml, "xl/worksheets/sheet1.xml")
        self.assertIn(f'HYPERLINK("{NEW}","{OLD}")'.encode(), result)
        self.assertIn(b"<v>Cached Label</v>", result)

    def test_signed_package_is_not_modified(self):
        path = self.root / "signed.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("_xmlsignatures/sig1.xml", b"<Signature/>")
        with self.assertRaisesRegex(ValueError, "signed"):
            self.editor.ooxml(path)

    def test_pdf_incremental_annotations_preserve_external_and_content(self):
        path = self.root / "test.pdf"
        writer = PdfWriter()
        page = writer.add_blank_page(600, 800)
        annotations = ArrayObject()
        for url in (OLD, "https://example.com"):
            annotation = DictionaryObject({
                NameObject("/Type"): NameObject("/Annot"),
                NameObject("/Subtype"): NameObject("/Link"),
                NameObject("/Rect"): ArrayObject([NumberObject(x) for x in (0, 0, 50, 50)]),
                NameObject("/A"): DictionaryObject({NameObject("/S"): NameObject("/URI"),
                                                    NameObject("/URI"): TextStringObject(url)})})
            annotations.append(writer._add_object(annotation))
        page[NameObject("/Annots")] = writer._add_object(annotations)
        with path.open("wb") as stream:
            writer.write(stream)
        original = path.read_bytes()
        result = self.editor.pdf(path)
        self.assertTrue(result.startswith(original))
        reader = PdfReader(io.BytesIO(result))
        targets = [str(a.get_object()["/A"]["/URI"]) for a in reader.pages[0]["/Annots"]]
        self.assertEqual(targets, [NEW, "https://example.com"])
        path.write_bytes(result)
        self.assertIsNone(self.editor.pdf(path))
        newer = fix.ORIGIN + "/sites/HR/Guide-v2.docx"
        self.resolver.add(NEW, newer, "later-migration.csv:2")
        second = self.editor.pdf(path)
        self.assertTrue(second.startswith(result))
        checked = PdfReader(io.BytesIO(second), strict=True)
        self.assertEqual(str(checked.pages[0]["/Annots"][0].get_object()["/A"]["/URI"]), newer)

    def test_pdf_direct_annotations_are_edited_without_changing_shared_action(self):
        path = self.root / "direct.pdf"
        writer = PdfWriter()
        page = writer.add_blank_page(600, 800)
        shared = writer._add_object(DictionaryObject({
            NameObject("/S"): NameObject("/URI"), NameObject("/URI"): TextStringObject(OLD)}))
        annotation = DictionaryObject({
            NameObject("/Subtype"): NameObject("/Link"),
            NameObject("/Rect"): ArrayObject([NumberObject(x) for x in (0, 0, 50, 50)]),
            NameObject("/A"): shared})
        untouched = DictionaryObject({
            NameObject("/Subtype"): NameObject("/Widget"), NameObject("/A"): shared})
        page[NameObject("/Annots")] = ArrayObject([annotation, writer._add_object(untouched)])
        with path.open("wb") as stream:
            writer.write(stream)
        result = self.editor.pdf(path)
        checked = PdfReader(io.BytesIO(result), strict=True)
        annotations = checked.pages[0]["/Annots"]
        self.assertEqual(str(annotations[0]["/A"]["/URI"]), NEW)
        self.assertEqual(str(annotations[1].get_object()["/A"]["/URI"]), OLD)


if __name__ == "__main__":
    unittest.main()
