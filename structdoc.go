package docxlib

import (
	"encoding/xml"
	"io"
	"strings"
	"sync"
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
	mu         sync.Mutex
	Paragraphs []Paragraph `xml:"w:p,omitempty"`

	file *Docx
}

func (b *Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "p" {
				var value Paragraph
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				if len(value.Children) > 0 {
					value.file = b.file
					b.mu.Lock()
					b.Paragraphs = append(b.Paragraphs, value)
					b.mu.Unlock()
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}

	}
	return nil

}

type Document struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	XMLW    string   `xml:"xmlns:w,attr"`            // cannot be unmarshalled in
	XMLR    string   `xml:"xmlns:r,attr,omitempty"`  // cannot be unmarshalled in
	XMLWP   string   `xml:"xmlns:wp,attr,omitempty"` // cannot be unmarshalled in
	// XMLWP14 string   `xml:"xmlns:wp14,attr,omitempty"` // cannot be unmarshalled in

	Body Body `xml:"w:body"`
}

func (doc *Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "body" {
				err = d.DecodeElement(&doc.Body, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			}
		}

	}
	return nil

}
