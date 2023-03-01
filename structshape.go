package docx

import (
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

// WPSWordprocessingShape is a container for a WordprocessingML DrawingML shape.
type WPSWordprocessingShape struct {
	XMLName xml.Name `xml:"wps:wsp,omitempty"`
	CNvCnPr *WPSCNvCnPr
	SpPr    *WPSSpPr
	BodyPr  *WPSBodyPr
}

// UnmarshalXML ...
func (w *WPSWordprocessingShape) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "cNvCnPr":
				w.CNvCnPr = new(WPSCNvCnPr)
				err = d.DecodeElement(w.CNvCnPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "spPr":
				w.SpPr = new(WPSSpPr)
				err = d.DecodeElement(w.SpPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "bodyPr":
				w.BodyPr = new(WPSBodyPr)
				err = d.DecodeElement(w.BodyPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// WPSCNvCnPr represents the non-visual drawing properties of a connector.
type WPSCNvCnPr struct {
	XMLName        xml.Name  `xml:"wps:cNvCnPr,omitempty"`
	ConnShapeLocks *struct{} `xml:"a:cxnSpLocks,omitempty"`
}

// UnmarshalXML ...
func (w *WPSCNvCnPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "cxnSpLocks":
				w.ConnShapeLocks = &struct{}{}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

// WPSSpPr is a container element that represents the visual properties of a shape.
type WPSSpPr struct {
	XMLName xml.Name `xml:"wps:spPr,omitempty"`
	BWMode  string   `xml:"bwMode,attr"`

	Xfrm     AXfrm
	PrstGeom APrstGeom
	NoFill   *struct{} `xml:"a:noFill,omitempty"`
	Elems    []interface{}
}

// UnmarshalXML ...
func (w *WPSSpPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "bwMode":
			w.BWMode = attr.Value
		default:
			// ignore other attributes
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

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "xfrm":
				err = d.DecodeElement(&w.Xfrm, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "prstGeom":
				w.PrstGeom.Prst = getAtt(tt.Attr, "prst")
			case "noFill":
				w.NoFill = &struct{}{}
			case "ln":
				var ln ALine
				err = d.DecodeElement(&ln, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &ln)
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

// ALine represents a line element in a Word document.
type ALine struct {
	XMLName  xml.Name `xml:"a:ln,omitempty"`
	W        int64    `xml:"w,attr"`
	Cap      string   `xml:"cap,attr,omitempty"`
	Compound string   `xml:"cmpd,attr,omitempty"`
	Align    string   `xml:"algn,attr,omitempty"`

	SolidFill *ASolidFill
	PrstDash  *APrstDash
	Miter     *AMiter
	Round     *struct{} `xml:"a:round,omitempty"`
	HeadEnd   *AHeadEnd
	TailEnd   *ATailEnd
}

// UnmarshalXML ...
func (l *ALine) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "w":
			l.W, err = strconv.ParseInt(attr.Value, 10, 64)
			if err != nil {
				return err
			}
		case "cap":
			l.Cap = attr.Value
		case "cmpd":
			l.Compound = attr.Value
		case "algn":
			l.Align = attr.Value
		default:
			// ignore other attributes
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

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "solidFill":
				l.SolidFill = new(ASolidFill)
				err = d.DecodeElement(l.SolidFill, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "prstDash":
				var value APrstDash
				value.Val = getAtt(tt.Attr, "val")
				l.PrstDash = &value
			case "miter":
				var value AMiter
				value.Limit = getAtt(tt.Attr, "lim")
				l.Miter = &value
			case "round":
				l.Round = &struct{}{}
			case "headEnd":
				l.HeadEnd = new(AHeadEnd)
				err = d.DecodeElement(l.HeadEnd, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tailEnd":
				l.TailEnd = new(ATailEnd)
				err = d.DecodeElement(l.TailEnd, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

// ASolidFill represents a solid fill of a shape or chart element.
type ASolidFill struct {
	XMLName xml.Name `xml:"a:solidFill,omitempty"`
	SrgbClr *ASrgbClr
}

// UnmarshalXML ...
func (s *ASolidFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "srgbClr":
				s.SrgbClr = new(ASrgbClr)
				for _, attr := range tt.Attr {
					if attr.Name.Local == "val" {
						s.SrgbClr.Val = attr.Value
						break
					}
				}
				err = d.Skip() // skip unsupported elements
				if err != nil {
					return err
				}
			default:
				err = d.Skip() // skip unsupported elements
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

// ASrgbClr represents an sRGB color.
type ASrgbClr struct {
	XMLName xml.Name `xml:"a:srgbClr,omitempty"`
	Val     string   `xml:"val,attr"`
}

// APrstDash ...
type APrstDash struct {
	XMLName xml.Name `xml:"a:prstDash,omitempty"`
	Val     string   `xml:"val,attr"`
}

// AMiter ...
type AMiter struct {
	XMLName xml.Name `xml:"a:miter,omitempty"`
	Limit   string   `xml:"lim,attr"`
}

// AHeadEnd ...
type AHeadEnd struct {
	XMLName xml.Name `xml:"a:headEnd,omitempty"`
	Type    string   `xml:"type,attr,omitempty"`
	W       string   `xml:"w,attr,omitempty"`
	Len     string   `xml:"len,attr,omitempty"`
}

// UnmarshalXML ...
func (r *AHeadEnd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			r.Type = attr.Value
		case "w":
			r.W = attr.Value
		case "len":
			r.Len = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}

// ATailEnd ...
type ATailEnd struct {
	XMLName xml.Name `xml:"a:tailEnd,omitempty"`
	Type    string   `xml:"type,attr,omitempty"`
	W       string   `xml:"w,attr,omitempty"`
	Len     string   `xml:"len,attr,omitempty"`
}

// UnmarshalXML ...
func (r *ATailEnd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			r.Type = attr.Value
		case "w":
			r.W = attr.Value
		case "len":
			r.Len = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}

// WPSBodyPr represents the body properties for a WordprocessingML DrawingML shape.
type WPSBodyPr struct {
	XMLName xml.Name `xml:"wps:bodyPr,omitempty"`
}
