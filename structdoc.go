package docxlib

import "encoding/xml"

const (
	XMLNS_W    = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R    = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
	XMLNS_WP   = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
	XMLNS_WP14 = `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`
)

func getAtt(atts []xml.Attr, name string) string {
	for _, at := range atts {
		if at.Name.Local == name {
			return at.Value
		}
	}
	return ""
}

type Body struct {
	XMLName    xml.Name     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main body"`
	Paragraphs []*Paragraph `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
}

type Document struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	XMLW    string   `xml:"xmlns:w,attr"`
	XMLR    string   `xml:"xmlns:r,attr"`
	XMLWP   string   `xml:"xmlns:wp,attr"`
	XMLWP14 string   `xml:"xmlns:wp14,attr"`
	Body    *Body
}
