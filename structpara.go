package docxlib

import (
	"crypto/md5"
	"encoding/hex"
	"encoding/xml"
	"io"
	"strings"
)

// ParagraphProperties <w:pPr>
type ParagraphProperties struct {
	XMLName       xml.Name       `xml:"w:pPr,omitempty"`
	Justification *Justification `xml:"w:jc,omitempty"`
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
			case "jc":
				p.Justification = &Justification{Val: getAtt(tt.Attr, "val")}
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
			switch {
			case o.Text != nil:
				sb.WriteString(o.Text.Text)
			case o.Drawing != nil:
				if o.Drawing.Inline != nil {
					sb.WriteString("![inline image: ")
					tgt, err := p.file.ReferTarget(o.Drawing.Inline.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
					if err != nil {
						sb.WriteString(err.Error())
					} else {
						h := md5.Sum(p.file.Media(tgt[6:]).Data)
						sb.WriteString(hex.EncodeToString(h[:]))
					}
					sb.WriteString("]()")
					continue
				}
				if o.Drawing.Anchor != nil {
					sb.WriteString("![anchor image: ")
					tgt, err := p.file.ReferTarget(o.Drawing.Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
					if err != nil {
						sb.WriteString(err.Error())
					} else {
						h := md5.Sum(p.file.Media(tgt[6:]).Data)
						sb.WriteString(hex.EncodeToString(h[:]))
					}
					sb.WriteString("]()")
				}
			}
		case *RunProperties:
			sb.WriteString("<prop>") //TODO: implement
		default:
			continue
		}
	}
	return sb.String()
}

// MarshalXML ...
func (p *Paragraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	err := e.EncodeToken(start)
	if err != nil {
		return err
	}
	if p.Properties != nil {
		err = e.Encode(p.Properties)
		if err != nil {
			return err
		}
	}
	for _, c := range p.Children {
		err = e.Encode(c)
		if err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
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
