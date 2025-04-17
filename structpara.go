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
	"reflect"
	"strings"
)

// ParagraphProperties <w:pPr>
type ParagraphProperties struct {
	XMLName        xml.Name `xml:"w:pPr,omitempty"`
	Tabs           *Tabs
	Spacing        *Spacing
	NumProperties  *NumProperties
	Ind            *Ind
	Justification  *Justification
	Shade          *Shade
	Kern           *Kern
	Style          *Style
	TextAlignment  *TextAlignment
	AdjustRightInd *AdjustRightInd
	SnapToGrid     *SnapToGrid
	Kinsoku        *Kinsoku
	OverflowPunct  *OverflowPunct

	RunProperties *RunProperties
}

// UnmarshalXML ...
func (p *ParagraphProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "tabs":
				var value Tabs
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Tabs = &value
			case "spacing":
				var value Spacing
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Spacing = &value
			case "ind":
				var value Ind
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Ind = &value
			case "jc":
				p.Justification = &Justification{Val: getAtt(tt.Attr, "val")}
			case "shd":
				var value Shade
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Shade = &value
			case "kern":
				var value Kern
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt64(v)
				if err != nil {
					return err
				}
				p.Kern = &value
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.RunProperties = &value
			case "pStyle":
				p.Style = &Style{Val: getAtt(tt.Attr, "val")}
			case "numPr":
				var value NumProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.NumProperties = &value
			case "textAlignment":
				p.TextAlignment = &TextAlignment{Val: getAtt(tt.Attr, "val")}
			case "adjustRightInd":
				var value AdjustRightInd
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.AdjustRightInd = &value
			case "snapToGrid":
				var value SnapToGrid
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.SnapToGrid = &value
			case "kinsoku":
				var value Kinsoku
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.Kinsoku = &value
			case "overflowPunct":
				var value OverflowPunct
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt(v)
				if err != nil {
					return err
				}
				p.OverflowPunct = &value

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

// Paragraph <w:p>
type Paragraph struct {
	XMLName xml.Name `xml:"w:p,omitempty"`

	// RsidR        string `xml:"w:rsidR,attr,omitempty"`
	// RsidRPr      string `xml:"w:rsidRPr,attr,omitempty"`
	// RsidRDefault string `xml:"w:rsidRDefault,attr,omitempty"`
	// RsidP        string `xml:"w:rsidP,attr,omitempty"`

	Properties *ParagraphProperties
	Children   []interface{}

	file *Docx
}

func (p *Paragraph) String() string {
	sb := strings.Builder{}
	for _, c := range p.Children {
		switch o := c.(type) {
		case *Hyperlink:
			id := o.ID
			text := o.Run.InstrText
			link, err := p.file.ReferTarget(id)
			sb.WriteString("[")
			sb.WriteString(text)
			sb.WriteString("](")
			if err != nil {
				sb.WriteString(id)
			} else {
				sb.WriteString(link)
			}
			sb.WriteByte(')')
		case *Run:
			for _, c := range o.Children {
				switch x := c.(type) {
				case *Text:
					sb.WriteString(x.Text)
				case *Tab:
					sb.WriteByte('\t')
				case *BarterRabbet:
					sb.WriteByte('\n')
				case *Drawing:
					if x.Inline != nil {
						sb.WriteString(x.Inline.String())
						continue
					}
					if x.Anchor != nil {
						sb.WriteString(x.Anchor.String())
						continue
					}
				}
			}
		default:
			continue
		}
	}
	return sb.String()
}

// UnmarshalXML ...
func (p *Paragraph) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	/*for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			p.RsidR = attr.Value
		case "rsidRPr":
			p.RsidRPr = attr.Value
		case "rsidRDefault":
			p.RsidRDefault = attr.Value
		case "rsidP":
			p.RsidP = attr.Value
		default:
			// ignore other attributes
		}
	}*/
	children := make([]interface{}, 0, 64)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			var elem interface{}
			switch tt.Name.Local {
			case "hyperlink":
				var value Hyperlink
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				id := getAtt(tt.Attr, "id")
				anchor := getAtt(tt.Attr, "anchor")
				if id != "" {
					value.ID = id
				}
				if anchor != "" {
					value.ID = anchor
				}
				elem = &value
			case "r":
				var value Run
				value.file = p.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				elem = &value
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				elem = &value
			case "pPr":
				var value ParagraphProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.Properties = &value
				continue
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
			children = append(children, elem)
		}
	}
	p.Children = children
	return nil
}

// KeepElements keep named elems amd removes others
//
// names: *docx.Hyperlink *docx.Run *docx.RunProperties
func (p *Paragraph) KeepElements(name ...string) {
	items := make([]interface{}, 0, len(p.Children))
	namemap := make(map[string]struct{}, len(name)*2)
	for _, n := range name {
		namemap[n] = struct{}{}
	}
	for _, item := range p.Children {
		_, ok := namemap[reflect.ValueOf(item).Type().String()]
		if ok {
			items = append(items, item)
		}
	}
	p.Children = items
}

// DropCanvas drops all canvases in paragraph
func (p *Paragraph) DropCanvas() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShape drops all shapes in paragraph
func (p *Paragraph) DropShape() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropGroup drops all groups in paragraph
func (p *Paragraph) DropGroup() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Group != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Group != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShapeAndCanvas drops all shapes and canvases in paragraph
func (p *Paragraph) DropShapeAndCanvas() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil || d.Inline.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil || d.Anchor.Graphic.GraphicData.Canvas != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropShapeAndCanvasAndGroup drops all shapes, canvases and groups in paragraph
func (p *Paragraph) DropShapeAndCanvasAndGroup() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Shape != nil || d.Inline.Graphic.GraphicData.Canvas != nil || d.Inline.Graphic.GraphicData.Group != nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Shape != nil || d.Anchor.Graphic.GraphicData.Canvas != nil || d.Anchor.Graphic.GraphicData.Group != nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}

// DropNilPicture drops all drawings with nil picture in paragraph
func (p *Paragraph) DropNilPicture() {
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			nrc := make([]interface{}, 0, len(r.Children))
			for _, rc := range r.Children {
				if d, ok := rc.(*Drawing); ok {
					if d.Inline == nil && d.Anchor == nil {
						continue
					}
					if (d.Inline != nil && d.Inline.Graphic == nil) || (d.Anchor != nil && d.Anchor.Graphic == nil) {
						continue
					}
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData == nil {
						continue
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData == nil {
						continue
					}
					if d.Inline != nil && d.Inline.Graphic != nil && d.Inline.Graphic.GraphicData != nil {
						if d.Inline.Graphic.GraphicData.Pic == nil {
							continue
						}
					}
					if d.Anchor != nil && d.Anchor.Graphic != nil && d.Anchor.Graphic.GraphicData != nil {
						if d.Anchor.Graphic.GraphicData.Pic == nil {
							continue
						}
					}
				}
				nrc = append(nrc, rc)
			}
			r.Children = nrc
		}
	}
}
