package docx

import "encoding/xml"

// RunStyle contains styling for a run
type RunStyle struct {
	XMLName xml.Name `xml:"w:rStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Style contains styling for a paragraph
type Style struct {
	XMLName xml.Name `xml:"w:pStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Color contains the sound of music. :D
// I'm kidding. It contains the color
type Color struct {
	XMLName xml.Name `xml:"w:color,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Size contains the font size
type Size struct {
	XMLName xml.Name `xml:"w:sz,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Justification contains the way of the horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
type Justification struct {
	XMLName xml.Name `xml:"w:jc,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Shade is an element that represents a shading pattern applied to a document element.
type Shade struct {
	XMLName       xml.Name `xml:"w:shd,omitempty"`
	Val           string   `xml:"w:val,attr,omitempty"`
	Color         string   `xml:"w:color,attr,omitempty"`
	Fill          string   `xml:"w:fill,attr,omitempty"`
	ThemeFill     string   `xml:"w:themeFill,attr,omitempty"`
	ThemeFillTint string   `xml:"w:themeFillTint,attr,omitempty"`
}

// UnmarshalXML ...
func (s *Shade) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "color":
			s.Color = attr.Value
		case "fill":
			s.Fill = attr.Value
		case "themeFill":
			s.ThemeFill = attr.Value
		case "themeFillTint":
			s.ThemeFillTint = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}
