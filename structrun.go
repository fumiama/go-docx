package docxlib

import (
	"encoding/xml"
	"io"
)

// Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName       xml.Name       `xml:"w:r,omitempty"`
	RunProperties *RunProperties `xml:"w:rPr,omitempty"`
	FrontTab      []struct {
		XMLName xml.Name `xml:"w:tab,omitempty"`
	}
	InstrText string `xml:"w:instrText,omitempty"`
	Text      *Text
	Drawing   *Drawing
	RearTab   []struct {
		XMLName xml.Name `xml:"w:tab,omitempty"`
	}
}

func (r *Run) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "rPr":
				var value RunProperties
				d.DecodeElement(&value, &tt)
				r.RunProperties = &value
			case "instrText":
				var value string
				d.DecodeElement(&value, &tt)
				r.InstrText = value
			case "t":
				var value Text
				d.DecodeElement(&value, &tt)
				r.Text = &value
			case "drawing":
				var value Drawing
				d.DecodeElement(&value, &tt)
				r.Drawing = &value
			case "tab":
				if r.InstrText == "" && r.Text == nil && r.Drawing == nil {
					r.FrontTab = append(r.FrontTab, struct {
						XMLName xml.Name "xml:\"w:tab,omitempty\""
					}{})
				} else {
					r.RearTab = append(r.RearTab, struct {
						XMLName xml.Name "xml:\"w:tab,omitempty\""
					}{})
				}
			default:
				continue
			}
		}

	}

	return nil

}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName  xml.Name  `xml:"w:rPr,omitempty"`
	Color    *Color    `xml:"w:color,omitempty"`
	Size     *Size     `xml:"w:sz,omitempty"`
	RunStyle *RunStyle `xml:"w:rStyle,omitempty"`
	Style    *Style    `xml:"w:pStyle,omitempty"`
}

func (r *RunProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "color":
				var value Color
				value.Val = getAtt(tt.Attr, "val")
				r.Color = &value
			case "sz":
				var value Size
				value.Val = getAtt(tt.Attr, "val")
				r.Size = &value
			case "rStyle":
				var value RunStyle
				value.Val = getAtt(tt.Attr, "val")
				r.RunStyle = &value
			case "pStyle":
				var value Style
				value.Val = getAtt(tt.Attr, "val")
				r.Style = &value
			default:
				continue
			}
		}

	}

	return nil

}

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
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

// Size contains the font size
type Size struct {
	XMLName xml.Name `xml:"w:sz"`
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
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}
