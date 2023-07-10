/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

// WordprocessingShape is a container for a WordprocessingML DrawingML shape.
type WordprocessingShape struct {
	XMLName xml.Name             `xml:"wps:wsp,omitempty"`
	CNvPr   *NonVisualProperties `xml:"wps:cNvPr,omitempty"`
	CNvCnPr *WPSCNvCnPr
	CNvSpPr *WPSCNvSpPr
	SpPr    *ShapeProperties `xml:"wps:spPr,omitempty"`
	TextBox *WPSTextBox
	BodyPr  *WPSBodyPr

	file *Docx
}

// UnmarshalXML ...
func (w *WordprocessingShape) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "cNvPr":
				w.CNvPr = new(NonVisualProperties)
				err = d.DecodeElement(w.CNvPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "cNvCnPr":
				w.CNvCnPr = new(WPSCNvCnPr)
				err = d.DecodeElement(w.CNvCnPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "cNvSpPr":
				w.CNvSpPr = new(WPSCNvSpPr)
				err = d.DecodeElement(w.CNvSpPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "spPr":
				w.SpPr = new(ShapeProperties)
				err = d.DecodeElement(w.SpPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "txbx":
				var value WPSTextBox
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.TextBox = &value
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
func (w *WPSCNvCnPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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

// WPSCNvSpPr represents the non-visual properties of a WordArt object.
type WPSCNvSpPr struct {
	XMLName xml.Name `xml:"wps:cNvSpPr,omitempty"`
	TxBox   int      `xml:"txBox,attr,omitempty"`

	SPLocks *ASPLocks
}

// UnmarshalXML ...
func (w *WPSCNvSpPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "txBox":
			w.TxBox, err = GetInt(attr.Value)
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
			case "spLocks":
				var value ASPLocks
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.SPLocks = &value
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

// ASPLocks represents the locks applied to a shape.
type ASPLocks struct {
	XMLName            xml.Name `xml:"a:spLocks,omitempty"`
	NoChangeArrowheads int      `xml:"noChangeArrowheads,attr,omitempty"`
}

// UnmarshalXML ...
func (l *ASPLocks) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "noChangeArrowheads":
			l.NoChangeArrowheads, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// ABlipFill represents a fill that contains a reference to an image.
type ABlipFill struct {
	XMLName      xml.Name `xml:"a:blipFill,omitempty"`
	DPI          int      `xml:"dpi,attr"`
	RotWithShape int      `xml:"rotWithShape,attr"`

	Blip    *ABlip
	SrcRect *ASrcRect
	Tile    *ATile
}

// UnmarshalXML ...
func (r *ABlipFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "dpi":
			r.DPI, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		case "rotWithShape":
			r.RotWithShape, err = GetInt(attr.Value)
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
			case "blip":
				var value ABlip
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Blip = &value
			case "srcRect":
				r.SrcRect = new(ASrcRect)
			case "tile":
				var value ATile
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Tile = &value
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

// ASrcRect represents the source rectangle of a tiled image fill.
type ASrcRect struct {
	XMLName xml.Name `xml:"a:srcRect,omitempty"`
}

// ATile represents the tiling information of a fill or border
type ATile struct {
	XMLName xml.Name `xml:"a:tile,omitempty"`
	TX      int64    `xml:"tx,attr"`
	TY      int64    `xml:"ty,attr"`
	SX      int64    `xml:"sx,attr"`
	SY      int64    `xml:"sy,attr"`
	Flip    string   `xml:"flip,attr"`
	Algn    string   `xml:"algn,attr"`
}

// UnmarshalXML ...
func (t *ATile) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "tx":
			t.TX, err = GetInt64(attr.Value)
		case "ty":
			t.TY, err = GetInt64(attr.Value)
		case "sx":
			t.SX, err = GetInt64(attr.Value)
		case "sy":
			t.SY, err = GetInt64(attr.Value)
		case "flip":
			t.Flip = attr.Value
		case "algn":
			t.Algn = attr.Value
		default:
			// ignore other attributes
		}
		if err != nil {
			return err
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// ALine represents a line element in a Word document.
type ALine struct {
	XMLName  xml.Name `xml:"a:ln,omitempty"`
	W        int64    `xml:"w,attr,omitempty"`
	Cap      string   `xml:"cap,attr,omitempty"`
	Compound string   `xml:"cmpd,attr,omitempty"`
	Align    string   `xml:"algn,attr,omitempty"`

	NoFill    *struct{} `xml:"a:noFill,omitempty"`
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
			l.W, err = GetInt64(attr.Value)
		case "cap":
			l.Cap = attr.Value
		case "cmpd":
			l.Compound = attr.Value
		case "algn":
			l.Align = attr.Value
		default:
			// ignore other attributes
		}
		if err != nil {
			return err
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
			case "noFill":
				l.NoFill = &struct{}{}
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
func (s *ASolidFill) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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

// WPSTextBox ...
type WPSTextBox struct {
	XMLName xml.Name `xml:"wps:txbx,omitempty"`
	Content *WTextBoxContent

	file *Docx
}

// UnmarshalXML ...
func (b *WPSTextBox) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "txbxContent":
				var value WTextBoxContent
				value.file = b.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Content = &value
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

// WTextBoxContent ...
type WTextBoxContent struct {
	XMLName    xml.Name    `xml:"w:txbxContent,omitempty"`
	Paragraphs []Paragraph `xml:"w:p,omitempty"`

	file *Docx
}

// UnmarshalXML ...
func (c *WTextBoxContent) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "p":
				var value Paragraph
				value.file = c.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Paragraphs = append(c.Paragraphs, value)
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

// WPSBodyPr represents the body properties for a WordprocessingML DrawingML shape.
type WPSBodyPr struct {
	XMLName   xml.Name `xml:"wps:bodyPr,omitempty"`
	Rot       int      `xml:"rot,attr"`
	Vert      string   `xml:"vert,attr,omitempty"`
	Wrap      string   `xml:"wrap,attr,omitempty"`
	LIns      int64    `xml:"lIns,attr"`
	TIns      int64    `xml:"tIns,attr"`
	RIns      int64    `xml:"rIns,attr"`
	BIns      int64    `xml:"bIns,attr"`
	Anchor    string   `xml:"anchor,attr,omitempty"`
	AnchorCtr int      `xml:"anchorCtr,attr"`
	Upright   int      `xml:"upright,attr"`

	NoAutofit *struct{} `xml:"a:noAutofit,omitempty"`
}

// UnmarshalXML ...
func (r *WPSBodyPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rot":
			r.Rot, _ = GetInt(attr.Value)
		case "vert":
			r.Vert = attr.Value
		case "wrap":
			r.Wrap = attr.Value
		case "lIns":
			r.LIns, _ = GetInt64(attr.Value)
		case "tIns":
			r.TIns, _ = GetInt64(attr.Value)
		case "rIns":
			r.RIns, _ = GetInt64(attr.Value)
		case "bIns":
			r.BIns, _ = GetInt64(attr.Value)
		case "anchor":
			r.Anchor = attr.Value
		case "anchorCtr":
			r.AnchorCtr, _ = GetInt(attr.Value)
		case "upright":
			r.Upright, _ = GetInt(attr.Value)
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
			case "noAutofit":
				r.NoAutofit = &struct{}{}
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
