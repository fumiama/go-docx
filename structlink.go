package docxlib

import (
	"encoding/xml"
	"io"
)

// Hyperlink element contains links
type Hyperlink struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Run     Run
}

func (r *Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.StartElement:
			if tt.Name.Local == "r" {
				d.DecodeElement(&r.Run, &start)
			} else {
				continue
			}
		}

	}
	return nil

}
