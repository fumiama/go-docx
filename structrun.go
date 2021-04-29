package docxlib

import "encoding/xml"

const (
	HYPERLINK_STYLE = "a1"
)

// A Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`
	RunProperties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
	InstrText     string         `xml:"w:instrText,omitempty"`
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
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Run     Run
}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName  xml.Name  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr"`
	Color    *Color    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color,omitempty"`
	Size     *Size     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sz,omitempty"`
	RunStyle *RunStyle `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rStyle,omitempty"`
}

// RunStyle contains styling for a run
type RunStyle struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rStyle"`
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
