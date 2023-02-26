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
	Ln       *ALine
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
				w.Ln = &ALine{}
				err = d.DecodeElement(&w.Ln, &tt)
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

// ALine represents a line element in a Word document.
type ALine struct {
	XMLName xml.Name `xml:"a:ln,omitempty"`
	W       int64    `xml:"w,attr"`

	SolidFill *ASolidFill
	Round     *struct{} `xml:"a:round,omitempty"`
	HeadEnd   *struct{} `xml:"a:headEnd,omitempty"`
	TailEnd   *struct{} `xml:"a:tailEnd,omitempty"`
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
			case "round":
				l.Round = &struct{}{}
			case "headEnd":
				l.HeadEnd = &struct{}{}
			case "tailEnd":
				l.TailEnd = &struct{}{}
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

// WPSBodyPr represents the body properties for a WordprocessingML DrawingML shape.
type WPSBodyPr struct {
	XMLName xml.Name `xml:"wps:bodyPr,omitempty"`
}
