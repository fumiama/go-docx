package docxlib

import "encoding/xml"

const (
	XMLNS_W = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
)

type Body struct {
	XMLName    xml.Name     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main body"`
	Paragraphs []*Paragraph `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
}

type Document struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	XMLW    string   `xml:"xmlns:w,attr"`
	XMLR    string   `xml:"xmlns:r,attr"`
	Body    *Body
}
