package docxlib

import (
	"encoding/xml"
	"testing"
)

const decoded_doc = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14"><w:body><w:p w14:paraId="77CA082D" w14:textId="4AF3264D" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:pPr><w:rPr><w:color w:val="808080"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r><w:t>test</w:t></w:r><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t xml:space="preserve"> font </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>size</w:t></w:r><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve"> color</w:t></w:r></w:p><w:p w14:paraId="6D114165" w14:textId="04580C29" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>New style 1</w:t></w:r></w:p><w:p w14:paraId="40D72B3B" w14:textId="76101901" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>New style 2</w:t></w:r></w:p><w:p w14:paraId="1CA8A9B3" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="44"/></w:rPr><w:t>test font size and color</w:t></w:r></w:p><w:p w14:paraId="0D82FB8B" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:hyperlink r:id="rId4"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>google</w:t></w:r></w:hyperlink></w:p><w:sectPr w:rsidR="00D66E3F"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`
const NUM_PARAGRAPHS = 5

func TestStructure(t *testing.T) {
	doc := Document{
		XMLW:    XMLNS_W,
		XMLR:    XMLNS_R,
		XMLName: xml.Name{Space: XMLNS_W, Local: "document"}}
	err := xml.Unmarshal([]byte(decoded_doc), &doc)
	if err != nil {
		t.Errorf("We expected to be able to decode %s but we didn't",
			decoded_doc)
	}
	if len(doc.Body.Paragraphs) != NUM_PARAGRAPHS {
		t.Errorf("We expected %d paragraphs, we got %d",
			NUM_PARAGRAPHS, len(doc.Body.Paragraphs))
	}
	for _, p := range doc.Body.Paragraphs {
		if len(p.Children()) == 0 {
			t.Errorf("We were not able to parse paragraph %v",
				p)
		}
		for _, child := range p.Children() {
			if child.Link == nil && child.Properties == nil && child.Run == nil {
				t.Errorf("There are children with all fields nil")
			}
			if child.Run != nil && child.Run.Text == nil {
				t.Errorf("We have a run with no text")
			}
		}
	}
}
