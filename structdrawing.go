package docxlib

import (
	"encoding/xml"
	"io"
)

// Drawing element contains photos
type Drawing struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main drawing,omitempty"`
	Inline  *WPInline
}

func (r *Drawing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.StartElement:
			switch tt.Name.Local {
			case "inline":
				r.Inline = new(WPInline)
				r.Inline.DistT = getAtt(tt.Attr, "distT")
				r.Inline.DistB = getAtt(tt.Attr, "distB")
				r.Inline.DistL = getAtt(tt.Attr, "distL")
				r.Inline.DistR = getAtt(tt.Attr, "distR")
				r.Inline.AnchorID = getAtt(tt.Attr, "anchorId")
				r.Inline.EditID = getAtt(tt.Attr, "editId")
				d.DecodeElement(r.Inline, &start)
			default:
				continue
			}
		}

	}
	return nil

}

// WPInline wp:inline
type WPInline struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing inline,omitempty"`
	DistT    string   `xml:"distT,attr"`
	DistB    string   `xml:"distB,attr"`
	DistL    string   `xml:"distL,attr"`
	DistR    string   `xml:"distR,attr"`
	AnchorID string   `xml:"wp14:anchorId,attr"`
	EditID   string   `xml:"wp14:editId,attr"`
}

func (r *WPInline) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.StartElement:
			switch tt.Name.Local {
			case "inline":

			default:
				continue
			}
		}

	}
	return nil

}
