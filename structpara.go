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
	"strconv"
	"strings"
)

// ParagraphProperties <w:pPr>
type ParagraphProperties struct {
	XMLName        xml.Name `xml:"w:pPr,omitempty"`
	Spacing        *Spacing
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
func (p *ParagraphProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
				value.Val, err = strconv.ParseInt(v, 10, 64)
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
			case "textAlignment":
				p.TextAlignment = &TextAlignment{Val: getAtt(tt.Attr, "val")}
			case "adjustRightInd":
				var value AdjustRightInd
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = strconv.Atoi(v)
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
				value.Val, err = strconv.Atoi(v)
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
				value.Val, err = strconv.Atoi(v)
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
				value.Val, err = strconv.Atoi(v)
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

	RsidR        string `xml:"w:rsidR,attr,omitempty"`
	RsidRPr      string `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string `xml:"w:rsidRDefault,attr,omitempty"`
	RsidP        string `xml:"w:rsidP,attr,omitempty"`

	Properties *ParagraphProperties
	Children   []interface{} // Children will generate an unnecessary tag <Children> ... </Children> and we skip it by a self-defined xml.Marshaler

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
func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
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
	}
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
