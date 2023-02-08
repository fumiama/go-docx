package docxlib

import (
	"encoding/xml"
	"io"

	"github.com/golang/glog"
)

type ParagraphChild struct {
	Link       *Hyperlink     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	Run        *Run           `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r,omitempty"`
	Properties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	Data    []ParagraphChild

	file *Docx
}

func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	children := make([]ParagraphChild, 0, 64)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		switch tt := t.(type) {
		case xml.StartElement:
			var elem ParagraphChild
			switch tt.Name.Local {
			case "hyperlink":
				var value Hyperlink
				d.DecodeElement(&value, &start)
				id := getAtt(tt.Attr, "id")
				anchor := getAtt(tt.Attr, "anchor")
				if id != "" {
					value.ID = id
				}
				if anchor != "" {
					value.ID = anchor
				}
				elem.Link = &value
			case "r":
				var value Run
				d.DecodeElement(&value, &start)
				elem.Run = &value
				if value.InstrText == "" && value.Text == nil {
					glog.V(0).Infof("Empty run, we ignore")
					continue
				}
			case "rPr":
				var value RunProperties
				d.DecodeElement(&value, &start)
				elem.Properties = &value
			default:
				continue
			}
			children = append(children, elem)
		}

	}
	*p = Paragraph{Data: children}
	return nil

}
