# -*- coding: utf-8 -*-
"""Biokimyo savol-javoblaridan formatlangan .docx fayl yaratuvchi skript.
python-docx mavjud bo'lmagani uchun docx (OOXML) qo'lda quriladi."""
import zipfile
from xml.sax.saxutils import escape

from qa_data import TITLE, ITEMS


def run(text, bold=False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return (
        '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(text))
    )


def para(runs, style=None, after=120):
    ppr = "<w:pPr>"
    if style:
        ppr += '<w:pStyle w:val="%s"/>' % style
    ppr += '<w:spacing w:after="%d" w:line="276" w:lineRule="auto"/>' % after
    ppr += "</w:pPr>"
    return "<w:p>" + ppr + "".join(runs) + "</w:p>"


def build_document():
    body = []
    # Title
    body.append(para([run(TITLE, bold=True)], style="Title", after=240))
    body.append(
        para(
            [run("Imtihonda og'zaki javob berishga mo'ljallangan kengaytirilgan javoblar", bold=False)],
            after=300,
        )
    )
    for q, a in ITEMS:
        body.append(para([run(q, bold=True)], after=40))
        body.append(para([run(a, bold=False)], after=200))

    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>" + "".join(body) +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:bottom="1134" w:left="1134" w:right="1134"/>'
        "</w:sectPr></w:body></w:document>"
    )
    return doc


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)

STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    "</w:styles>"
)


def main():
    doc_xml = build_document()
    out = "BIOKIMYO - kengaytirilgan javoblar.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
    print("Yaratildi:", out, "| savollar soni:", len(ITEMS))


if __name__ == "__main__":
    main()
