package docxlib

import (
	"encoding/xml"
	"io"
)

const (
	HYPERLINK_STYLE = "a1"
)

// A Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r,omitempty"`
	RunProperties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
	InstrText     string         `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main instrText,omitempty"`
	Text          *Text
}

// The Text object contains the actual text
type Text struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main t"`
	XMLSpace string   `xml:"xml:space,attr,omitempty"`
	Text     string   `xml:",chardata"`
}

// The hyperlink element contains links
type Hyperlink struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Run     Run
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
	var elem Run
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}

		switch tt := t.(type) {
		case xml.CharData:
			var value string
			d.DecodeElement(&value, &start)
			elem.Text = &Text{Text: value}
		case xml.StartElement:
			if tt.Name.Local == "rPr" {
				var value RunProperties
				d.DecodeElement(&value, &start)
				elem.RunProperties = &value
			} else if tt.Name.Local == "instrText" {
				var value string
				d.DecodeElement(&value, &start)
				elem.InstrText = value
			} else {
				continue
			}
		}

	}
	*r = elem
	return nil

}
