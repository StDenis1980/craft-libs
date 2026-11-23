//! Fixes from the Tinta Office fork of WordCraft (StDenis1980/wordcraft), moved here with the
//! .docx reader and writer: widow/orphan control a file never sets, partial table cell margins,
//! and the core properties Word writes beyond the usual ones.

use std::io::{Read, Write};

fn docx(styles: Option<&str>) -> Vec<u8> {
    let root = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let doc = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>"#;
    let mut parts = vec![("_rels/.rels", root.to_string()), ("word/document.xml", doc.to_string())];
    if let Some(s) = styles {
        parts.push(("word/_rels/document.xml.rels", doc_rels.to_string()));
        parts.push((
            "word/styles.xml",
            format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{s}</w:styles>"#),
        ));
    }
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, xml) in &parts {
        zw.start_file(*name, zip::write::SimpleFileOptions::default()).unwrap();
        zw.write_all(xml.as_bytes()).unwrap();
    }
    zw.finish().unwrap().into_inner()
}

fn styles_xml(bytes: &[u8]) -> String {
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut s = String::new();
    z.by_name("word/styles.xml").unwrap().read_to_string(&mut s).unwrap();
    s
}

fn widow_control(d: &craft_wordmodel::Document) -> bool {
    let p = d.body[0].as_para().unwrap();
    d.styles.resolve_para(&p.props).widow_control
}

const NORMAL: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#;

#[test]
fn a_docx_that_never_sets_widow_control_has_it_off() {
    for styles in [Some(NORMAL.to_string()), Some(format!("<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>{NORMAL}"))] {
        let d = craft_docx::read(&docx(styles.as_deref())).unwrap();
        assert!(!widow_control(&d), "off when nothing sets it: {styles:?}");
        let bytes = craft_docx::write(&d).unwrap();
        assert!(!styles_xml(&bytes).contains("widowControl"), "written back by omission: {styles:?}");
        assert!(!widow_control(&craft_docx::read(&bytes).unwrap()), "still off after a save: {styles:?}");
    }
}

#[test]
fn widow_control_set_in_the_file_stays_on() {
    for styles in [
        format!("<w:docDefaults><w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>{NORMAL}"),
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:widowControl/></w:pPr></w:style>"#
            .to_string(),
    ] {
        let d = craft_docx::read(&docx(Some(&styles))).unwrap();
        assert!(widow_control(&d), "on: {styles}");
        assert!(widow_control(&craft_docx::read(&craft_docx::write(&d).unwrap()).unwrap()), "on after a save: {styles}");
    }
}

#[test]
fn a_new_document_keeps_widow_control_through_a_save() {
    let mut d = craft_wordmodel::Document::new();
    d.body = vec![craft_wordmodel::para_block(craft_wordmodel::Paragraph::with_text("Text", Default::default()))];
    assert!(widow_control(&d), "the app's default is on");
    let reopened = craft_docx::read(&craft_docx::write(&d).unwrap()).unwrap();
    assert!(widow_control(&reopened), "and stays on in its .docx");
}

/// A table whose `w:tblCellMar` sets only the left side keeps the other sides of its style (or,
/// without one, Word's own 0.08" left and right, in a file that has a default table style as Word
/// writes it; without one, `no_default_margins` gives none).
#[test]
fn a_left_only_table_margin_keeps_the_other_sides() {
    let ns = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let styles = format!(
        r#"<w:styles {ns}><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style><w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:top w:w="20" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"#
    );
    let table = |style: &str| {
        format!(
            r#"<w:tbl><w:tblPr>{style}<w:tblCellMar><w:left w:w="107" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"#
        )
    };
    let document = format!(r#"<w:document {ns}><w:body>{}{}</w:body></w:document>"#, table(r#"<w:tblStyle w:val="Grid"/>"#), table(""));
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, data) in [
        ("_rels/.rels", root),
        ("word/_rels/document.xml.rels", rels),
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", styles.as_str()),
    ] {
        zip.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    let doc = craft_docx::read(&zip.finish().unwrap().into_inner()).unwrap();
    let margins: Vec<_> = doc.body.iter().filter_map(|b| b.as_table()).map(|t| t.props.cell_margins).collect();
    assert_eq!(
        margins,
        vec![Some([1.0, 5.35, 2.0, 8.0]), Some([0.0, 5.35, 0.0, 5.4])],
        "the left side is the table's, the rest the style's or Word's"
    );
}

/// The language, status, identifier, version and last-printed date in `docProps/core.xml` survive
/// a save, as Word keeps them.
#[test]
fn core_properties_beyond_the_usual_ones_survive_a_save() {
    let mut d = craft_wordmodel::Document::new();
    d.body = vec![craft_wordmodel::para_block(craft_wordmodel::Paragraph::with_text("Text", Default::default()))];
    d.core.language = "es-ES".into();
    d.core.content_status = "Final".into();
    d.core.identifier = "EXP-2026/014".into();
    d.core.version = "1.2".into();
    d.core.last_printed = "2026-11-02T09:30:00Z".into();
    let bytes = craft_docx::write(&d).unwrap();
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(&bytes)).unwrap();
    let mut core = String::new();
    z.by_name("docProps/core.xml").unwrap().read_to_string(&mut core).unwrap();
    for tag in ["<dc:language>es-ES</dc:language>", "<cp:contentStatus>Final</cp:contentStatus>", "<cp:version>1.2</cp:version>"] {
        assert!(core.contains(tag), "{tag} in {core}");
    }
    let c = craft_docx::read(&bytes).unwrap().core;
    assert_eq!(
        (c.language.as_str(), c.content_status.as_str(), c.identifier.as_str(), c.version.as_str(), c.last_printed.as_str()),
        ("es-ES", "Final", "EXP-2026/014", "1.2", "2026-11-02T09:30:00Z")
    );
}

/// A file without a font table, with a chart kept as its own part: the first save builds a font
/// table and the next save writes that table as read, in the same place, so the two saves list
/// the parts in the same order (Tinta Office documents compare saves byte for byte).
#[test]
fn a_built_font_table_is_written_in_the_same_place_when_saved_again() {
    let w = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#;
    let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
    let document = format!(
        r#"<w:document {w}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Text</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="5" name="G"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#
    );
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, data) in
        [("_rels/.rels", root), ("word/_rels/document.xml.rels", rels), ("word/document.xml", document.as_str()), ("word/charts/chart1.xml", chart)]
    {
        zw.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        zw.write_all(data.as_bytes()).unwrap();
    }
    let types = |bytes: &[u8]| {
        let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut s = String::new();
        z.by_name("[Content_Types].xml").unwrap().read_to_string(&mut s).unwrap();
        s
    };
    let first = craft_docx::write(&craft_docx::read(&zw.finish().unwrap().into_inner()).unwrap()).unwrap();
    let second = craft_docx::write(&craft_docx::read(&first).unwrap()).unwrap();
    let (a, b) = (types(&first), types(&second));
    assert!(a.contains("/word/fontTable.xml") && a.contains("/word/charts/"), "a font table is built and the chart kept: {a}");
    assert_eq!(a, b, "the second save lists the parts as the first");
}
