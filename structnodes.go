package docxlib

import (
	"encoding/xml"
	"io"
)

type ParagraphChild struct {
	Link       *Hyperlink     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	Run        *Run           `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r,omitempty"`
	Properties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	Data    []ParagraphChild

	file *DocxLib
}

func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	children := make([]ParagraphChild, 0)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		switch tt := t.(type) {
		case xml.StartElement:
			var elem ParagraphChild
			if tt.Name.Local == "hyperlink" {
				var value Hyperlink
				d.DecodeElement(&value, &start)
				value.ID = getAtt(tt.Attr, "id")
				elem = ParagraphChild{Link: &value}
			} else if tt.Name.Local == "r" {
				var value Run
				d.DecodeElement(&value, &start)
				elem = ParagraphChild{Run: &value}
			} else if tt.Name.Local == "rPr" {
				var value RunProperties
				d.DecodeElement(&value, &start)
				elem = ParagraphChild{Properties: &value}
			} else {
				continue
			}
			children = append(children, elem)
		}

	}
	*p = Paragraph{Data: children}
	return nil

}
