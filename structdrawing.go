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
	"crypto/md5"
	"encoding/hex"
	"encoding/xml"
	"io"
	"strconv"
	"strings"
	"sync/atomic"
)

//nolint:revive,stylecheck
const (
	// A4_EMU_MAX_WIDTH is the max display width of an A4 paper
	A4_EMU_MAX_WIDTH = 5274310
)

//nolint:revive,stylecheck
const (
	XMLNS_DRAWINGML_MAIN    = `http://schemas.openxmlformats.org/drawingml/2006/main`
	XMLNS_DRAWINGML_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
)

// Drawing element contains photos
type Drawing struct {
	XMLName xml.Name `xml:"w:drawing,omitempty"`
	Inline  *WPInline
	Anchor  *WPAnchor

	file *Docx
}

// UnmarshalXML ...
func (r *Drawing) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "inline":
				r.Inline = new(WPInline)
				r.Inline.file = r.file
				err = d.DecodeElement(r.Inline, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "anchor":
				r.Anchor = new(WPAnchor)
				r.Anchor.file = r.file
				err = d.DecodeElement(r.Anchor, &tt)
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

func (r *Drawing) copymedia(to *Docx) *Drawing {
	if r.Inline != nil {
		return &Drawing{
			Inline: r.Inline.copymedia(to),
			file:   to,
		}
	}
	if r.Anchor != nil {
		return &Drawing{
			Anchor: r.Anchor.copymedia(to),
			file:   to,
		}
	}
	return &Drawing{file: to}
}

// WPInline is an element that represents an inline image within a text paragraph.
//
// It contains information about the image's size and position,
// as well as any non-visual properties associated with the image.
// The <wp:inline> element can contain child elements such as <wp:extent> to specify
// the dimensions of the image and <wp:cNvGraphicFramePr> to specify the non-visual
// properties of the image. Inline images are often used in documents where the images
// are meant to be treated as part of the text flow, such as in a newsletter or a product brochure.
type WPInline struct {
	XMLName xml.Name `xml:"wp:inline,omitempty"`
	DistT   int64    `xml:"distT,attr"`
	DistB   int64    `xml:"distB,attr"`
	DistL   int64    `xml:"distL,attr"`
	DistR   int64    `xml:"distR,attr"`
	// AnchorID string   `xml:"wp14:anchorId,attr,omitempty"`
	// EditID   string   `xml:"wp14:editId,attr,omitempty"`

	Extent            *WPExtent
	EffectExtent      *WPEffectExtent
	DocPr             *WPDocPr
	CNvGraphicFramePr *WPCNvGraphicFramePr
	Graphic           *AGraphic

	file *Docx
}

// UnmarshalXML ...
func (r *WPInline) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "distT":
			r.DistT, err = GetInt64(attr.Value)
		case "distB":
			r.DistB, err = GetInt64(attr.Value)
		case "distL":
			r.DistL, err = GetInt64(attr.Value)
		case "distR":
			r.DistR, err = GetInt64(attr.Value)
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
			case "extent":
				r.Extent = new(WPExtent)
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "cx":
						r.Extent.CX, err = GetInt64(v.Value)
					case "cy":
						r.Extent.CY, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
			case "effectExtent":
				r.EffectExtent = new(WPEffectExtent)
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "l":
						r.EffectExtent.L, err = GetInt64(v.Value)
					case "t":
						r.EffectExtent.T, err = GetInt64(v.Value)
					case "r":
						r.EffectExtent.R, err = GetInt64(v.Value)
					case "b":
						r.EffectExtent.B, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
			case "docPr":
				r.DocPr = new(WPDocPr)
				err = d.DecodeElement(r.DocPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "cNvGraphicFramePr":
				var value WPCNvGraphicFramePr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.CNvGraphicFramePr = &value
			case "graphic":
				var value AGraphic
				value.file = r.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Graphic = &value
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

// String ...
func (r *WPInline) String() string {
	sb := strings.Builder{}
	if r.Graphic.GraphicData.Pic != nil {
		sb.WriteString("![inlnim ")
		switch {
		case r.DocPr != nil:
			sb.WriteString(r.DocPr.Name)
		case r.Graphic.GraphicData.Pic.NonVisualPicProperties != nil:
			sb.WriteString(r.Graphic.GraphicData.Pic.NonVisualPicProperties.NonVisualDrawingProperties.Name)
		default:
			sb.WriteString(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
		}
		sb.WriteString("](")
		if r.Graphic.GraphicData.Pic.BlipFill != nil {
			tgt, err := r.file.ReferTarget(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
			if err != nil {
				sb.WriteString(err.Error())
			} else {
				h := md5.Sum(r.file.Media(tgt[6:]).Data)
				sb.WriteString(hex.EncodeToString(h[:]))
			}
		}
		sb.WriteByte(')')
		return sb.String()
	}
	if r.Graphic.GraphicData.Shape != nil {
		sb.WriteString("![inlnsp ")
		switch {
		case r.DocPr != nil:
			sb.WriteString(r.DocPr.Name)
		case r.Graphic.GraphicData.Shape.CNvPr != nil:
			sb.WriteString(r.Graphic.GraphicData.Shape.CNvPr.Name)
		case r.Graphic.GraphicData.Shape.SpPr != nil:
			sb.WriteString(r.Graphic.GraphicData.Shape.SpPr.PrstGeom.Prst)
		default:
			sb.WriteString("nil")
		}
		sb.WriteString("](")
		if r.Graphic.GraphicData.Shape.SpPr != nil {
			sb.WriteString(r.Graphic.GraphicData.Shape.SpPr.PrstGeom.Prst)
		}
		sb.WriteByte(')')
		return sb.String()
	}
	if r.Graphic.GraphicData.Canvas != nil {
		sb.WriteString("![inlncv ")
		if r.DocPr != nil {
			sb.WriteString(r.DocPr.Name)
		} else {
			sb.WriteString("nil")
		}
		sb.WriteString("]()")
		return sb.String()
	}
	return "![inln?](unknown)"
}

func (r *WPInline) copymedia(to *Docx) *WPInline { //nolint: dupl
	if r.Graphic.GraphicData.Pic != nil {
		if r.Graphic.GraphicData.Pic.BlipFill != nil {
			tgt, err := r.file.ReferTarget(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
			if err != nil {
				return nil
			}
			format := tgt[strings.LastIndex(tgt, ".")+1:]
			idn := int(atomic.AddUintptr(&to.docID, 1))
			id := int(to.IncreaseID("图片"))
			ids := strconv.Itoa(id)
			m := r.file.Media(tgt[6:])
			if m == nil {
				return nil
			}
			rid := to.addImage(format, m.Data)
			inln := *r
			grph := *r.Graphic
			inln.Graphic = &grph
			grphdata := *r.Graphic.GraphicData
			grph.GraphicData = &grphdata
			pic := *r.Graphic.GraphicData.Pic
			grphdata.Pic = &pic
			grphdata.file = to
			grph.file = to
			inln.file = to

			inln.DocPr = &WPDocPr{
				ID:   idn,
				Name: "图片 " + ids,
			}
			pic.NonVisualPicProperties = &PICNonVisualPicProperties{
				NonVisualDrawingProperties: NonVisualProperties{
					ID:   id,
					Name: "图片 " + ids,
				},
				CNvPicPr: r.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr,
			}
			pic.BlipFill = &PICBlipFill{
				Blip: ABlip{
					Embed:  rid,
					Cstate: r.Graphic.GraphicData.Pic.BlipFill.Blip.Cstate,
				},
				Stretch: r.Graphic.GraphicData.Pic.BlipFill.Stretch,
			}
			return &inln
		}
		return nil
	}
	if r.Graphic.GraphicData.Shape != nil { // shape has no media
		return r
	}
	if r.Graphic.GraphicData.Canvas != nil { //TODO: copy canvas media
		return r
	}
	return nil
}

// WPExtent represents the extent of a drawing in a Word document.
//
//	CX CY 's unit is English Metric Units, which is 1/914400 inch
type WPExtent struct {
	XMLName xml.Name `xml:"wp:extent,omitempty"`
	CX      int64    `xml:"cx,attr"`
	CY      int64    `xml:"cy,attr"`
}

// UnmarshalXML ...
func (r *WPExtent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "cx":
			r.CX, err = GetInt64(attr.Value)
		case "cy":
			r.CY, err = GetInt64(attr.Value)
		}
		if err != nil {
			return err
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// WPEffectExtent represents the effect extent of a drawing in a Word document.
type WPEffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent,omitempty"`
	L       int64    `xml:"l,attr"`
	T       int64    `xml:"t,attr"`
	R       int64    `xml:"r,attr"`
	B       int64    `xml:"b,attr"`
}

// UnmarshalXML ...
func (r *WPEffectExtent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "l":
			r.L, err = GetInt64(attr.Value)
		case "t":
			r.T, err = GetInt64(attr.Value)
		case "r":
			r.R, err = GetInt64(attr.Value)
		case "b":
			r.B, err = GetInt64(attr.Value)
		}
		if err != nil {
			return err
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// WPDocPr represents the document properties of a drawing in a Word document.
type WPDocPr struct {
	XMLName xml.Name `xml:"wp:docPr,omitempty"`
	ID      int      `xml:"id,attr"`
	Name    string   `xml:"name,attr,omitempty"`
}

// UnmarshalXML ...
func (r *WPDocPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			id, err := GetInt(attr.Value)
			if err != nil {
				return err
			}
			r.ID = id
		case "name":
			r.Name = attr.Value

		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}

// WPCNvGraphicFramePr represents the non-visual properties of a graphic frame.
type WPCNvGraphicFramePr struct {
	XMLName xml.Name `xml:"wp:cNvGraphicFramePr,omitempty"`
	Locks   AGraphicFrameLocks
}

// UnmarshalXML ...
func (w *WPCNvGraphicFramePr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "graphicFrameLocks":
				w.Locks.XMLA = getAtt(tt.Attr, "a")
				v := getAtt(tt.Attr, "noChangeAspect")
				if v == "" {
					continue
				}
				w.Locks.NoChangeAspect, err = GetInt(v)
				if err != nil {
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

// AGraphicFrameLocks represents the locks applied to a graphic frame.
type AGraphicFrameLocks struct {
	XMLName        xml.Name `xml:"a:graphicFrameLocks,omitempty"`
	XMLA           string   `xml:"xmlns:a,attr,omitempty"`
	NoChangeAspect int      `xml:"noChangeAspect,attr,omitempty"`
}

// AGraphic represents a graphic in a Word document.
type AGraphic struct {
	XMLName     xml.Name `xml:"a:graphic,omitempty"`
	XMLA        string   `xml:"xmlns:a,attr,omitempty"`
	GraphicData *AGraphicData

	file *Docx
}

// UnmarshalXML ...
func (a *AGraphic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "a":
			a.XMLA = attr.Value
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
			case "graphicData":
				var value AGraphicData
				value.file = a.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.URI = getAtt(tt.Attr, "uri")
				a.GraphicData = &value
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

// AGraphicData represents the data of a graphic in a Word document.
type AGraphicData struct {
	XMLName xml.Name `xml:"a:graphicData,omitempty"`
	URI     string   `xml:"uri,attr"`

	Pic    *Picture
	Shape  *WordprocessingShape
	Canvas *WordprocessingCanvas
	Group  *WordprocessingGroup

	file *Docx
}

// UnmarshalXML ...
func (a *AGraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "uri":
			a.URI = attr.Value
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
			case "pic":
				var value Picture
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.XMLPIC = getAtt(tt.Attr, "pic")
				a.Pic = &value
			case "wsp":
				var value WordprocessingShape
				value.file = a.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				a.Shape = &value
			case "wpc":
				var value WordprocessingCanvas
				value.file = a.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				a.Canvas = &value
			case "wgp":
				var value WordprocessingGroup
				value.file = a.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				a.Group = &value
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

// Picture represents a picture in a Word document.
type Picture struct {
	XMLName                xml.Name `xml:"pic:pic,omitempty"`
	XMLPIC                 string   `xml:"xmlns:pic,attr,omitempty"`
	NonVisualPicProperties *PICNonVisualPicProperties
	BlipFill               *PICBlipFill
	SpPr                   *PICSpPr
}

// UnmarshalXML ...
func (p *Picture) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "nvPicPr":
				var value PICNonVisualPicProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.NonVisualPicProperties = &value
			case "blipFill":
				var value PICBlipFill
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.BlipFill = &value
			case "spPr":
				var value PICSpPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.SpPr = &value
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

// PICNonVisualPicProperties represents the non-visual properties of a picture in a Word document.
type PICNonVisualPicProperties struct {
	XMLName                    xml.Name            `xml:"pic:nvPicPr,omitempty"`
	NonVisualDrawingProperties NonVisualProperties `xml:"pic:cNvPr,omitempty"`
	CNvPicPr                   PicCNvPicPr
}

// UnmarshalXML ...
func (p *PICNonVisualPicProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				v := getAtt(tt.Attr, "id")
				if v == "" {
					continue
				}
				p.NonVisualDrawingProperties.ID, err = GetInt(v)
				if err != nil {
					return err
				}
				p.NonVisualDrawingProperties.Name = getAtt(tt.Attr, "name")
			case "cNvPicPr":
				err = d.DecodeElement(&p.CNvPicPr, &tt)
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

// PicCNvPicPr represents the non-visual properties of a picture.
type PicCNvPicPr struct {
	XMLName xml.Name `xml:"pic:cNvPicPr,omitempty"`
	Locks   *APicLocks
}

// UnmarshalXML ...
func (p *PicCNvPicPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	// Loop through XML tokens
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		// Switch based on token type
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "picLocks":
				var value APicLocks
				v := getAtt(tt.Attr, "noChangeAspect")
				if v == "" {
					continue
				}
				value.NoChangeAspect, err = GetInt(v)
				if err != nil {
					return err
				}
				p.Locks = &value
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

// APicLocks represents the locks applied to a picture.
type APicLocks struct {
	XMLName        xml.Name `xml:"a:picLocks,omitempty"`
	NoChangeAspect int      `xml:"noChangeAspect,attr"`
}

// PICBlipFill represents the blip fill of a picture in a Word document.
type PICBlipFill struct {
	XMLName xml.Name `xml:"pic:blipFill,omitempty"`
	Blip    ABlip
	Stretch AStretch
}

// UnmarshalXML ...
func (p *PICBlipFill) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				err = d.DecodeElement(&p.Blip, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "stretch":
				err = d.DecodeElement(&p.Stretch, &tt)
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

// ABlip represents the blip of a picture in a Word document.
type ABlip struct {
	XMLName     xml.Name `xml:"a:blip,omitempty"`
	Embed       string   `xml:"r:embed,attr"`
	Cstate      string   `xml:"cstate,attr,omitempty"`
	AlphaModFix *AAlphaModFix
}

// UnmarshalXML ...
func (a *ABlip) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "embed":
			a.Embed = attr.Value
		case "cstate":
			a.Cstate = attr.Value
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
			case "alphaModFix":
				var value AAlphaModFix
				v := getAtt(tt.Attr, "amt")
				if v == "" {
					continue
				}
				value.Amount, err = GetInt(v)
				if err != nil {
					return err
				}
				a.AlphaModFix = &value
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

// AAlphaModFix ...
type AAlphaModFix struct {
	XMLName xml.Name `xml:"a:alphaModFix,omitempty"`
	Amount  int      `xml:"amt,attr"`
}

// AStretch ...
type AStretch struct {
	XMLName  xml.Name `xml:"a:stretch,omitempty"`
	FillRect *AFillRect
}

// UnmarshalXML ...
func (s *AStretch) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "fillRect":
				var value AFillRect
				/*err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}*/
				s.FillRect = &value
				err = d.Skip() // skip unparsed innerxml
				if err != nil {
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

// AFillRect ...
type AFillRect struct {
	XMLName xml.Name `xml:"a:fillRect,omitempty"`
}

// PICSpPr is a struct representing the <pic:spPr> element in OpenXML,
// which describes the shape properties for a picture.
type PICSpPr struct {
	XMLName  xml.Name `xml:"pic:spPr,omitempty"`
	Xfrm     AXfrm
	PrstGeom *APrstGeom
}

// UnmarshalXML ...
func (p *PICSpPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				err = d.DecodeElement(&p.Xfrm, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "prstGeom":
				var value APrstGeom
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.PrstGeom = &value
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

// AXfrm is a struct representing the <a:xfrm> element in OpenXML,
// which describes the position and size of a shape.
type AXfrm struct {
	XMLName xml.Name `xml:"a:xfrm,omitempty"`
	Rot     int64    `xml:"rot,attr,omitempty"`
	FlipH   int      `xml:"flipH,attr,omitempty"`
	FlipV   int      `xml:"flipV,attr,omitempty"`
	Off     AOff     `xml:"a:off,omitempty"`
	Ext     AExt     `xml:"a:ext,omitempty"`
	ChOff   *AOff    `xml:"a:chOff,omitempty"`
	ChExt   *AExt    `xml:"a:chExt,omitempty"`
}

// UnmarshalXML ...
func (a *AXfrm) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rot":
			a.Rot, err = GetInt64(attr.Value)
		case "flipH":
			a.FlipH, err = GetInt(attr.Value)
		case "flipV":
			a.FlipV, err = GetInt(attr.Value)
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
			case "off":
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "x":
						a.Off.X, err = GetInt64(v.Value)
					case "y":
						a.Off.Y, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
			case "ext":
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "cx":
						a.Ext.CX, err = GetInt64(v.Value)
					case "cy":
						a.Ext.CY, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
			case "chOff":
				var value AOff
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "x":
						value.X, err = GetInt64(v.Value)
					case "y":
						value.Y, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
				a.ChOff = &value
			case "chExt":
				var value AExt
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "cx":
						value.CX, err = GetInt64(v.Value)
					case "cy":
						value.CY, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
				a.ChExt = &value
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

// AOff is a struct representing the <a:off> / <a:chOff> element in OpenXML,
// which describes the offset of a shape from its original position.
type AOff struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// AExt is a struct representing the <a:ext> / <a:chExt> element in OpenXML,
// which describes the size of a shape.
type AExt struct {
	CX int64 `xml:"cx,attr"`
	CY int64 `xml:"cy,attr"`
}

// APrstGeom is a struct representing the <a:prstGeom> element in OpenXML,
// which describes the preset shape geometry for a shape.
type APrstGeom struct {
	XMLName xml.Name  `xml:"a:prstGeom,omitempty"`
	Prst    string    `xml:"prst,attr"`
	AvLst   *struct{} `xml:"a:avLst,omitempty"`
}

// UnmarshalXML ...
func (a *APrstGeom) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "prst":
			a.Prst = attr.Value
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
			case "avLst":
				a.AvLst = &struct{}{}
				err = d.Skip() // skip innerxml
				if err != nil {
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

/*
// AAvLst is a struct representing the <a:avLst> element in OpenXML,
// which describes the adjustments to the shape's preset geometry.
type AAvLst struct {
	XMLName xml.Name `xml:"a:avLst,omitempty"`
	RawXML  string   `xml:",innerxml"`
}

// UnmarshalXML ...
func (a *AAvLst) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) (err error) {
	var content []byte

	if content, err = xml.Marshal(start); err != nil {
		return err
	}

	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if end, ok := t.(xml.EndElement); ok && end == start.End() {
			break
		}

		b, err := xml.Marshal(t)
		if err != nil {
			return err
		}

		content = append(content, b...)
	}

	a.RawXML = BytesToString(content)

	return nil
}*/

// WPAnchor is an element that represents an anchored object in a Word document.
//
// It allows for the positioning of a drawing object relative to a specific location
// in the text of the document. The <wp:anchor> element contains child elements that
// specify the dimensions and position of the anchored object, as well as the non-visual
// properties of the object. The <wp:anchor> element can contain the <wp:docPr> element,
// which contains the non-visual properties of the anchored object, such as its ID and name,
// as well as the <a:graphic> element, which specifies the visual properties of the object,
// such as its shape and fill.
type WPAnchor struct {
	XMLName        xml.Name `xml:"wp:anchor,omitempty"`
	DistT          int64    `xml:"distT,attr"`
	DistB          int64    `xml:"distB,attr"`
	DistL          int64    `xml:"distL,attr"`
	DistR          int64    `xml:"distR,attr"`
	SimplePos      int      `xml:"simplePos,attr"`
	RelativeHeight int      `xml:"relativeHeight,attr"`
	BehindDoc      int      `xml:"behindDoc,attr"`
	Locked         int      `xml:"locked,attr"`
	LayoutInCell   int      `xml:"layoutInCell,attr"`
	AllowOverlap   int      `xml:"allowOverlap,attr"`

	SimplePosXY       *WPSimplePos
	PositionH         *WPPositionH
	PositionV         *WPPositionV
	Extent            *WPExtent
	EffectExtent      *WPEffectExtent
	WrapNone          *struct{} `xml:"wp:wrapNone,omitempty"`
	WrapSquare        *WPWrapSquare
	DocPr             *WPDocPr
	CNvGraphicFramePr *WPCNvGraphicFramePr
	Graphic           *AGraphic

	file *Docx
}

// UnmarshalXML ...
func (r *WPAnchor) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, tt := range start.Attr {
		switch tt.Name.Local {
		case "distT":
			r.DistT, err = GetInt64(tt.Value)
		case "distB":
			r.DistB, err = GetInt64(tt.Value)
		case "distL":
			r.DistL, err = GetInt64(tt.Value)
		case "distR":
			r.DistR, err = GetInt64(tt.Value)
		case "simplePos":
			r.SimplePos, err = GetInt(tt.Value)
		case "relativeHeight":
			r.RelativeHeight, err = GetInt(tt.Value)
		case "behindDoc":
			r.BehindDoc, err = GetInt(tt.Value)
		case "locked":
			r.Locked, err = GetInt(tt.Value)
		case "layoutInCell":
			r.LayoutInCell, err = GetInt(tt.Value)
		case "allowOverlap":
			r.AllowOverlap, err = GetInt(tt.Value)
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
			case "simplePos":
				r.SimplePosXY = new(WPSimplePos)
				for _, v := range tt.Attr {
					switch v.Name.Local {
					case "x":
						r.SimplePosXY.X, err = GetInt64(v.Value)
					case "y":
						r.SimplePosXY.Y, err = GetInt64(v.Value)
					}
					if err != nil {
						return err
					}
				}
			case "positionH":
				r.PositionH = new(WPPositionH)
				// r.PositionH.RelativeFrom = getAtt(tt.Attr, "relativeFrom")
				err = d.DecodeElement(&r.PositionH, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "positionV":
				r.PositionV = new(WPPositionV)
				// r.PositionV.RelativeFrom = getAtt(tt.Attr, "relativeFrom")
				err = d.DecodeElement(&r.PositionV, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "extent":
				r.Extent = new(WPExtent)
				err = d.DecodeElement(&r.Extent, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "effectExtent":
				r.EffectExtent = new(WPEffectExtent)
				err = d.DecodeElement(&r.EffectExtent, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "wrapNone":
				r.WrapNone = &struct{}{}
			case "wrapSquare":
				r.WrapSquare = new(WPWrapSquare)
				r.WrapSquare.WrapText = getAtt(tt.Attr, "wrapText")
			case "docPr":
				r.DocPr = new(WPDocPr)
				err = d.DecodeElement(r.DocPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "cNvGraphicFramePr":
				r.CNvGraphicFramePr = new(WPCNvGraphicFramePr)
				err = d.DecodeElement(r.CNvGraphicFramePr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "graphic":
				r.Graphic = new(AGraphic)
				r.Graphic.file = r.file
				err = d.DecodeElement(&r.Graphic, &tt)
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

// String ...
func (r *WPAnchor) String() string {
	sb := strings.Builder{}
	if r.Graphic != nil && r.Graphic.GraphicData != nil {
		if r.Graphic.GraphicData.Pic != nil {
			sb.WriteString("![anchim ")
			switch {
			case r.DocPr != nil:
				sb.WriteString(r.DocPr.Name)
			case r.Graphic.GraphicData.Pic.NonVisualPicProperties != nil:
				sb.WriteString(r.Graphic.GraphicData.Pic.NonVisualPicProperties.NonVisualDrawingProperties.Name)
			default:
				sb.WriteString(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
			}
			sb.WriteString("](")
			if r.Graphic.GraphicData.Pic.BlipFill != nil {
				tgt, err := r.file.ReferTarget(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
				if err != nil {
					sb.WriteString(err.Error())
				} else {
					h := md5.Sum(r.file.Media(tgt[6:]).Data)
					sb.WriteString(hex.EncodeToString(h[:]))
				}
			}
			sb.WriteByte(')')
			return sb.String()
		}
		if r.Graphic.GraphicData.Shape != nil {
			sb.WriteString("![anchsp ")
			switch {
			case r.DocPr != nil:
				sb.WriteString(r.DocPr.Name)
			case r.Graphic.GraphicData.Shape.CNvPr != nil:
				sb.WriteString(r.Graphic.GraphicData.Shape.CNvPr.Name)
			case r.Graphic.GraphicData.Shape.SpPr != nil:
				sb.WriteString(r.Graphic.GraphicData.Shape.SpPr.PrstGeom.Prst)
			default:
				sb.WriteString("nil")
			}
			sb.WriteString("](")
			if r.Graphic.GraphicData.Shape.SpPr != nil {
				sb.WriteString(r.Graphic.GraphicData.Shape.SpPr.PrstGeom.Prst)
			}
			sb.WriteByte(')')
			return sb.String()
		}
		if r.Graphic.GraphicData.Canvas != nil {
			sb.WriteString("![anchcv ")
			if r.DocPr != nil {
				sb.WriteString(r.DocPr.Name)
			} else {
				sb.WriteString("nil")
			}
			sb.WriteString("]()")
			return sb.String()
		}
	}
	return "![anch?](unknown)"
}

func (r *WPAnchor) copymedia(to *Docx) *WPAnchor { //nolint: dupl
	if r.Graphic.GraphicData.Pic != nil {
		if r.Graphic.GraphicData.Pic.BlipFill != nil {
			tgt, err := r.file.ReferTarget(r.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
			if err != nil {
				return nil
			}
			format := tgt[strings.LastIndex(tgt, ".")+1:]
			idn := int(atomic.AddUintptr(&to.docID, 1))
			id := int(to.IncreaseID("图片"))
			ids := strconv.Itoa(id)
			m := r.file.Media(tgt[6:])
			if m == nil {
				return nil
			}
			rid := to.addImage(format, m.Data)
			anch := *r
			grph := *r.Graphic
			anch.Graphic = &grph
			grphdata := *r.Graphic.GraphicData
			grph.GraphicData = &grphdata
			pic := *r.Graphic.GraphicData.Pic
			grphdata.Pic = &pic
			grphdata.file = to
			grph.file = to
			anch.file = to

			anch.DocPr = &WPDocPr{
				ID:   idn,
				Name: "图片 " + ids,
			}
			pic.NonVisualPicProperties = &PICNonVisualPicProperties{
				NonVisualDrawingProperties: NonVisualProperties{
					ID:   id,
					Name: "图片 " + ids,
				},
				CNvPicPr: r.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr,
			}
			pic.BlipFill = &PICBlipFill{
				Blip: ABlip{
					Embed:  rid,
					Cstate: r.Graphic.GraphicData.Pic.BlipFill.Blip.Cstate,
				},
				Stretch: r.Graphic.GraphicData.Pic.BlipFill.Stretch,
			}
			return &anch
		}
		return nil
	}
	if r.Graphic.GraphicData.Shape != nil { // shape has no media
		return r
	}
	if r.Graphic.GraphicData.Canvas != nil { //TODO: copy canvas media
		return r
	}
	return nil
}

// WPSimplePos represents the position of an object in a Word document.
type WPSimplePos struct {
	XMLName xml.Name `xml:"wp:simplePos,omitempty"`
	X       int64    `xml:"x,attr"`
	Y       int64    `xml:"y,attr"`
}

// WPPositionH represents the horizontal position of an object in a Word document.
type WPPositionH struct {
	XMLName      xml.Name `xml:"wp:positionH,omitempty"`
	RelativeFrom string   `xml:"relativeFrom,attr"`
	PosOffset    int64    `xml:"wp:posOffset"`
}

// UnmarshalXML ...
func (r *WPPositionH) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "relativeFrom" {
			r.RelativeFrom = attr.Value
			break
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
			case "posOffset":
				err = d.DecodeElement(&r.PosOffset, &tt)
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

// WPPositionV represents the vertical position of an object in a Word document.
type WPPositionV struct {
	XMLName      xml.Name `xml:"wp:positionV,omitempty"`
	RelativeFrom string   `xml:"relativeFrom,attr"`
	PosOffset    int64    `xml:"wp:posOffset"`
}

// UnmarshalXML ...
func (r *WPPositionV) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "relativeFrom" {
			r.RelativeFrom = attr.Value
			break
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
			case "posOffset":
				err = d.DecodeElement(&r.PosOffset, &tt)
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

// WPWrapSquare represents the square wrapping of an object in a Word document.
type WPWrapSquare struct {
	XMLName  xml.Name `xml:"wp:wrapSquare,omitempty"`
	WrapText string   `xml:"wrapText,attr"`
}
