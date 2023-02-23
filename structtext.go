package docxlib

import (
	"encoding/xml"
	"io"
)

// Text object contains the actual text
type Text struct {
	XMLName xml.Name `xml:"w:t,omitempty"`

	// XMLSpace string   `xml:"xml:space,attr,omitempty"`

	Text string `xml:",chardata"`
}

// UnmarshalXML ...
func (r *Text) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.CharData); ok {
			r.Text = string(tt) // implicitly copy
		}
	}

	return nil
}
