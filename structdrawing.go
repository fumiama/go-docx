package docxlib

import (
	"encoding/xml"
	"io"
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

		switch tt := t.(type) {
		case xml.StartElement:
			switch tt.Name.Local {
			case "inline":
				r.Inline = new(WPInline)
				r.Inline.DistT = getAtt(tt.Attr, "distT")
				r.Inline.DistB = getAtt(tt.Attr, "distB")
				r.Inline.DistL = getAtt(tt.Attr, "distL")
				r.Inline.DistR = getAtt(tt.Attr, "distR")
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
	DistT    string   `xml:"distT,attr"`
	DistB    string   `xml:"distB,attr"`
	DistL    string   `xml:"distL,attr"`
	DistR    string   `xml:"distR,attr"`
	AnchorID string   `xml:"wp14:anchorId,attr"`
	EditID   string   `xml:"wp14:editId,attr"`

	Extent       *WPExtent
	EffectExtent *WPEffectExtent
	DocPr        *WPDocPr
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
			case "extent":
				r.Extent.CX = getAtt(tt.Attr, "cx")
				r.Extent.CY = getAtt(tt.Attr, "cy")
			case "effectExtent":
				r.EffectExtent.L = getAtt(tt.Attr, "l")
				r.EffectExtent.T = getAtt(tt.Attr, "t")
				r.EffectExtent.R = getAtt(tt.Attr, "r")
				r.EffectExtent.B = getAtt(tt.Attr, "b")
			case "docPr":
				r.DocPr.ID = getAtt(tt.Attr, "id")
				r.DocPr.Name = getAtt(tt.Attr, "name")
				r.DocPr.Macro = getAtt(tt.Attr, "macro")
				r.DocPr.Hidden = getAtt(tt.Attr, "hidden")
			default:
				continue
			}
		}

	}
	return nil

}

// WPExtent represents the extent of a drawing in a Word document.
type WPExtent struct {
	XMLName xml.Name `xml:"wp:extent,omitempty"`
	CX      string   `xml:"cx,attr"`
	CY      string   `xml:"cy,attr"`
}

// WPEffectExtent represents the effect extent of a drawing in a Word document.
type WPEffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent,omitempty"`
	L       string   `xml:"l,attr"`
	T       string   `xml:"t,attr"`
	R       string   `xml:"r,attr"`
	B       string   `xml:"b,attr"`
}

// WPDocPr represents the document properties of a drawing in a Word document.
type WPDocPr struct {
	XMLName xml.Name `xml:"wp:docPr,omitempty"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr,omitempty"`
	Macro   string   `xml:"macro,attr,omitempty"`
	Hidden  string   `xml:"hidden,attr,omitempty"`
}

// AGraphic represents a graphic in a Word document.
type AGraphic struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic,omitempty"`
	GraphicData *AGraphicData
}

// AGraphicData represents the data of a graphic in a Word document.
type AGraphicData struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicData,omitempty"`
	URI     string   `xml:"uri,attr"`
}
