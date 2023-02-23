package docxlib

import (
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
				continue
			}
		}

	}
	return nil

}

// Paragraph <w:p>
type Paragraph struct {
	// XMLName    xml.Name `xml:"w:p,omitempty"`
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
			sb.WriteString(text)
			sb.WriteByte('(')
			if err != nil {
				sb.WriteString(id)
			} else {
				sb.WriteString(link)
			}
			sb.WriteByte(')')
		case *Run:
			sb.WriteString("run") //TODO: implement
		case *RunProperties:
			sb.WriteString("prop") //TODO: implement
		default:
			continue
		}
	}
	return sb.String()
}

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
				continue
			}
			children = append(children, elem)
		}

	}
	p.Children = children
	return nil

}
