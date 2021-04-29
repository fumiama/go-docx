package docxlib

import "encoding/xml"

type ParagraphChild struct {
	Link *Hyperlink `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink"`
	Run  *Run       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	Data    []ParagraphChild

	file *DocxLib
}
