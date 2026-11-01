import os
import re
import sys

root = sys.argv[1]


def path(name):
    return os.path.join(root, name)


def edit(name, old, new):
    xml = open(path(name), encoding="utf-8").read()
    open(path(name), "w", encoding="utf-8").write(xml.replace(old, new, 1))


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
open(path("word/footnotes.xml"), "w", encoding="utf-8").write(
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes {W}>'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p>'
    '</w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/>'
    '</w:r><w:r><w:t xml:space="preserve"> Source: field survey 2026.</w:t></w:r></w:p></w:footnote></w:footnotes>')
edit("word/_rels/document.xml.rels", "</Relationships>",
     '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
     'footnotes" Target="footnotes.xml"/></Relationships>')
edit("[Content_Types].xml", "</Types>", '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.'
     'openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
# the sentence ends inside the run " sites. Results follow below.": split it and put the mark between
doc = open(path("word/document.xml"), encoding="utf-8").read()
run = re.search(r'<w:r>(<w:rPr/>|<w:rPr>(?:(?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>( sites\.)( Results follow below\.)</w:t></w:r>', doc)
rpr = run.group(1) or ""
new = (f'<w:r>{rpr}<w:t xml:space="preserve">{run.group(2)}</w:t></w:r>'
       '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
       f'<w:r>{rpr}<w:t xml:space="preserve">{run.group(3)}</w:t></w:r>')
open(path("word/document.xml"), "w", encoding="utf-8").write(doc.replace(run.group(0), new))
