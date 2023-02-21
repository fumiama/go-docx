package docxlib

import (
	"encoding/xml"
	"io"
)

const (
	XMLNS_W  = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R  = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
	XMLNS_WP = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
	// XMLNS_WP14 = `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`

	XMLNS_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
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
	XMLName    xml.Name     `xml:"w:body"`
	Paragraphs []*Paragraph `xml:"w:p,omitempty"`
}

type Document struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	XMLW    string   `xml:"xmlns:w,attr"`              // cannot be unmarshalled in
	XMLR    string   `xml:"xmlns:r,attr,omitempty"`    // cannot be unmarshalled in
	XMLWP   string   `xml:"xmlns:wp,attr,omitempty"`   // cannot be unmarshalled in
	XMLWP14 string   `xml:"xmlns:wp14,attr,omitempty"` // cannot be unmarshalled in
	Body    *Body

	file *Docx
}

func (doc *Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	if doc.Body == nil {
		doc.Body = &Body{
			Paragraphs: make([]*Paragraph, 0, 64),
		}
	}
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		switch tt := t.(type) {
		case xml.StartElement:
			switch tt.Name.Local {
			case "body":
			case "p":
				var value Paragraph
				d.DecodeElement(&value, &start)
				if len(value.Children) > 0 {
					value.file = doc.file
					doc.Body.Paragraphs = append(doc.Body.Paragraphs, &value)
				}
			default:
				d.Skip()
				continue
			}
		}

	}
	return nil

}
