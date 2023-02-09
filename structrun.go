package docxlib

import (
	"encoding/xml"
	"io"
)

const (
	HYPERLINK_STYLE = "a1"
)

// Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r,omitempty"`
	RunProperties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
	InstrText     string         `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main instrText,omitempty"`
	Text          *Text
	Drawing       *Drawing
}

// Text object contains the actual text
type Text struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main t"`
	XMLSpace string   `xml:"xml:space,attr,omitempty"`
	Text     string   `xml:",chardata"`
}

// Hyperlink element contains links
type Hyperlink struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Run     Run
}

// Drawing element contains photos
type Drawing struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main drawing,omitempty"`
	Inline  *WPInline
}

// WPInline wp:inline
type WPInline struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing inline,omitempty"`
	DistT   string   `xml:"distT,attr"`
	DistB   string   `xml:"distB,attr"`
	DistL   string   `xml:"distL,attr"`
	DistR   string   `xml:"distR,attr"`
}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName  xml.Name  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
	Color    *Color    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color,omitempty"`
	Size     *Size     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sz,omitempty"`
	RunStyle *RunStyle `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rStyle,omitempty"`
	Style    *Style    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main pStyle,omitempty"`
}

// RunStyle contains styling for a run
type RunStyle struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Style contains styling for a paragraph
type Style struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main pStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Color contains the sound of music. :D
// I'm kidding. It contains the color
type Color struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color"`
	Val     string   `xml:"w:val,attr"`
}

// Size contains the font size
type Size struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sz"`
	Val     int      `xml:"w:val,attr"`
}

func (r *Run) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.StartElement:
			switch tt.Name.Local {
			case "rPr":
				var value RunProperties
				d.DecodeElement(&value, &start)
				r.RunProperties = &value
			case "instrText":
				var value string
				d.DecodeElement(&value, &start)
				r.InstrText = value
			case "t":
				var value Text
				d.DecodeElement(&value, &start)
				r.Text = &value
			case "drawing":
				var value Drawing
				d.DecodeElement(&value, &start)
				r.Drawing = &value
			default:
				continue
			}
		}

	}

	return nil

}
func (r *Text) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.CharData:
			r.Text = string(tt) // implicitly copy
		}

	}

	return nil
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
				d.DecodeElement(r.Inline, &start)
			default:
				continue
			}
		}

	}
	return nil

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
func (r *RunStyle) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.StartElement:
			r.Val = getAtt(tt.Attr, "val")
		}

	}
	return nil

}

func getAtt(atts []xml.Attr, name string) string {
	for _, at := range atts {
		if at.Name.Local == name {
			return at.Value
		}
	}
	return ""
}
