package docxlib

import (
	"encoding/xml"
	"io"
	"strconv"
)

const (
	// A4_EMU_MAX_WIDTH is the max display width of an A4 paper
	A4_EMU_MAX_WIDTH = 5274310
)

const (
	XMLNS_DRAWINGML_MAIN    = `http://schemas.openxmlformats.org/drawingml/2006/main`
	XMLNS_DRAWINGML_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
)

// Drawing element contains photos
type Drawing struct {
	XMLName xml.Name `xml:"w:drawing,omitempty"`
	Inline  *WPInline
}

func (r *Drawing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "inline":
				r.Inline = new(WPInline)
				r.Inline.DistT, err = strconv.Atoi(getAtt(tt.Attr, "distT"))
				if err != nil {
					return err
				}
				r.Inline.DistB, err = strconv.Atoi(getAtt(tt.Attr, "distB"))
				if err != nil {
					return err
				}
				r.Inline.DistL, err = strconv.Atoi(getAtt(tt.Attr, "distL"))
				if err != nil {
					return err
				}
				r.Inline.DistR, err = strconv.Atoi(getAtt(tt.Attr, "distR"))
				if err != nil {
					return err
				}
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
	XMLName  xml.Name `xml:"wp:inline,omitempty"`
	DistT    int      `xml:"distT,attr"`
	DistB    int      `xml:"distB,attr"`
	DistL    int      `xml:"distL,attr"`
	DistR    int      `xml:"distR,attr"`
	AnchorID string   `xml:"wp14:anchorId,attr,omitempty"`
	EditID   string   `xml:"wp14:editId,attr,omitempty"`

	Extent            *WPExtent
	EffectExtent      *WPEffectExtent
	DocPr             *WPDocPr
	CNvGraphicFramePr *WPCNvGraphicFramePr
	Graphic           *AGraphic
}

func (r *WPInline) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "extent":
				r.Extent = new(WPExtent)
				r.Extent.CX, err = strconv.Atoi(getAtt(tt.Attr, "cx"))
				if err != nil {
					return err
				}
				r.Extent.CY, err = strconv.Atoi(getAtt(tt.Attr, "cy"))
				if err != nil {
					return err
				}
			case "effectExtent":
				r.EffectExtent = new(WPEffectExtent)
				r.EffectExtent.L, err = strconv.Atoi(getAtt(tt.Attr, "l"))
				if err != nil {
					return err
				}
				r.EffectExtent.T, err = strconv.Atoi(getAtt(tt.Attr, "t"))
				if err != nil {
					return err
				}
				r.EffectExtent.R, err = strconv.Atoi(getAtt(tt.Attr, "r"))
				if err != nil {
					return err
				}
				r.EffectExtent.B, err = strconv.Atoi(getAtt(tt.Attr, "b"))
				if err != nil {
					return err
				}
			case "docPr":
				r.DocPr = new(WPDocPr)
				r.DocPr.ID = getAtt(tt.Attr, "id")
				r.DocPr.Name = getAtt(tt.Attr, "name")
				r.DocPr.Macro = getAtt(tt.Attr, "macro")
				r.DocPr.Hidden = getAtt(tt.Attr, "hidden")
			case "cNvGraphicFramePr":
				var value WPCNvGraphicFramePr
				d.DecodeElement(&value, &start)
				r.CNvGraphicFramePr = &value
			case "graphic":
				var value AGraphic
				d.DecodeElement(&value, &start)
				value.XMLA = getAtt(tt.Attr, "a")
				r.Graphic = &value
			default:
				continue
			}
		}

	}
	return nil

}

// WPExtent represents the extent of a drawing in a Word document.
//
//	CX CY 's unit is English Metric Units, which is 1/914400 inch
type WPExtent struct {
	XMLName xml.Name `xml:"wp:extent,omitempty"`
	CX      int      `xml:"cx,attr"`
	CY      int      `xml:"cy,attr"`
}

// WPEffectExtent represents the effect extent of a drawing in a Word document.
type WPEffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent,omitempty"`
	L       int      `xml:"l,attr"`
	T       int      `xml:"t,attr"`
	R       int      `xml:"r,attr"`
	B       int      `xml:"b,attr"`
}

// WPDocPr represents the document properties of a drawing in a Word document.
type WPDocPr struct {
	XMLName xml.Name `xml:"wp:docPr,omitempty"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr,omitempty"`
	Macro   string   `xml:"macro,attr,omitempty"`
	Hidden  string   `xml:"hidden,attr,omitempty"`
}

// WPCNvGraphicFramePr represents the non-visual properties of a graphic frame.
type WPCNvGraphicFramePr struct {
	XMLName xml.Name `xml:"wp:cNvGraphicFramePr,omitempty"`
	Locks   *AGraphicFrameLocks
}

func (w *WPCNvGraphicFramePr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "graphicFrameLocks":
				var value AGraphicFrameLocks
				d.DecodeElement(&value, &start)
				value.NoChangeAspect, err = strconv.Atoi(getAtt(tt.Attr, "noChangeAspect"))
				if err != nil {
					return err
				}
				w.Locks = &value
			default:
				continue
			}
		}

	}
	return nil
}

// AGraphicFrameLocks represents the locks applied to a graphic frame.
type AGraphicFrameLocks struct {
	XMLName        xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicFrameLocks,omitempty"`
	NoChangeAspect int      `xml:"noChangeAspect,attr"`
}

// AGraphic represents a graphic in a Word document.
type AGraphic struct {
	XMLName     xml.Name `xml:"a:graphic,omitempty"`
	XMLA        string   `xml:"xmlns:a,attr,omitempty"`
	GraphicData *AGraphicData
}

func (a *AGraphic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "graphicData":
				var value AGraphicData
				d.DecodeElement(&value, &start)
				value.URI = getAtt(tt.Attr, "uri")
				a.GraphicData = &value
			default:
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
	Pic     *PICPic
}

func (a *AGraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "pic":
				var value PICPic
				d.DecodeElement(&value, &start)
				value.XMLPIC = getAtt(tt.Attr, "pic")
				a.Pic = &value
			default:
				continue
			}
		}

	}
	return nil
}

// PICPic represents a picture in a Word document.
type PICPic struct {
	XMLName                xml.Name `xml:"pic:pic,omitempty"`
	XMLPIC                 string   `xml:"xmlns:pic,attr,omitempty"`
	NonVisualPicProperties *PICNonVisualPicProperties
	BlipFill               *PICBlipFill
	SpPr                   *PICSpPr
}

func (p *PICPic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "nvPicPr":
				var value PICNonVisualPicProperties
				d.DecodeElement(&value, &start)
				p.NonVisualPicProperties = &value
			case "blipFill":
				var value PICBlipFill
				d.DecodeElement(&value, &start)
				p.BlipFill = &value
			case "spPr":
				var value PICSpPr
				d.DecodeElement(&value, &start)
				p.SpPr = &value
			default:
				continue
			}
		}

	}
	return nil
}

// PICNonVisualPicProperties represents the non-visual properties of a picture in a Word document.
type PICNonVisualPicProperties struct {
	XMLName                    xml.Name `xml:"pic:nvPicPr,omitempty"`
	NonVisualDrawingProperties PICNonVisualDrawingProperties
}

func (p *PICNonVisualPicProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "cNvPr":
				p.NonVisualDrawingProperties.ID = getAtt(tt.Attr, "id")
			default:
				continue
			}
		}

	}
	return nil
}

// PICNonVisualDrawingProperties represents the non-visual drawing properties of a picture in a Word document.
type PICNonVisualDrawingProperties struct {
	XMLName xml.Name `xml:"pic:cNvPr,omitempty"`
	ID      string   `xml:"id,attr"`
}

// PICBlipFill represents the blip fill of a picture in a Word document.
type PICBlipFill struct {
	XMLName xml.Name `xml:"pic:blipFill,omitempty"`
	Blip    ABlip
	Stretch AStretch
}

func (p *PICBlipFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "blip":
				p.Blip.Embed = getAtt(tt.Attr, "embed")
				p.Blip.Cstate = getAtt(tt.Attr, "cstate")
			case "stretch":
				d.DecodeElement(&p.Stretch, &start)
			default:
				continue
			}
		}

	}
	return nil
}

// ABlip represents the blip of a picture in a Word document.
type ABlip struct {
	XMLName xml.Name `xml:"a:blip,omitempty"`
	Embed   string   `xml:"r:embed,attr"`
	Cstate  string   `xml:"cstate,attr"`
}

// AStretch ...
type AStretch struct {
	XMLName  xml.Name `xml:"a:stretch,omitempty"`
	FillRect AFillRect
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
	PrstGeom APrstGeom
}

func (p *PICSpPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "xfrm":
				d.DecodeElement(&p.Xfrm, &start)
			case "prstGeom":
				d.DecodeElement(&p.PrstGeom, &start)
				p.PrstGeom.Prst = getAtt(tt.Attr, "prst")
			default:
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
	Off     AOff
	Ext     AExt
}

func (a *AXfrm) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "off":
				a.Off.X, err = strconv.Atoi(getAtt(tt.Attr, "x"))
				if err != nil {
					return err
				}
				a.Off.Y, err = strconv.Atoi(getAtt(tt.Attr, "y"))
				if err != nil {
					return err
				}
			case "ext":
				a.Ext.CX, err = strconv.Atoi(getAtt(tt.Attr, "cx"))
				if err != nil {
					return err
				}
				a.Ext.CY, err = strconv.Atoi(getAtt(tt.Attr, "cy"))
				if err != nil {
					return err
				}
			default:
				continue
			}
		}

	}
	return nil
}

// AOff is a struct representing the <a:off> element in OpenXML,
// which describes the offset of a shape from its original position.
type AOff struct {
	XMLName xml.Name `xml:"a:off,omitempty"`
	X       int      `xml:"x,attr"`
	Y       int      `xml:"y,attr"`
}

// AExt is a struct representing the <a:ext> element in OpenXML,
// which describes the size of a shape.
type AExt struct {
	XMLName xml.Name `xml:"a:ext,omitempty"`
	CX      int      `xml:"cx,attr"`
	CY      int      `xml:"cy,attr"`
}

// APrstGeom is a struct representing the <a:prstGeom> element in OpenXML,
// which describes the preset shape geometry for a shape.
type APrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom,omitempty"`
	Prst    string   `xml:"prst,attr"`
	AvLst   AAvLst
}

// AAvLst is a struct representing the <a:avLst> element in OpenXML,
// which describes the adjustments to the shape's preset geometry.
type AAvLst struct {
	XMLName xml.Name `xml:"a:avLst,omitempty"`
	RawXML  string   `xml:",innerxml"`
}

func (a *AAvLst) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
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
}
