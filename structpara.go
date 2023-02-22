package docxlib

import (
	"encoding/xml"
	"io"
	"strings"
)

type ParagraphProperties struct {
	XMLName       xml.Name       `xml:"w:pPr,omitempty"`
	Justification *Justification `xml:"w:jc,omitempty"`
}

func (p *ParagraphProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "jc":
				p.Justification = &Justification{Val: getAtt(tt.Attr, "val")}
			default:
				continue
			}
		}

	}
	return nil

}

type ParagraphChild struct {
	Link       *Hyperlink     `xml:"w:hyperlink,omitempty"`
	Run        *Run           `xml:"w:r,omitempty"`
	Properties *RunProperties `xml:"w:rPr,omitempty"`
}

type Paragraph struct {
	XMLName    xml.Name `xml:"w:p,omitempty"`
	Properties *ParagraphProperties
	Children   []ParagraphChild // Children will generate an unnecessary tag <Children> ... </Children> and we skip it by a self-defined xml.Marshaler

	file *Docx
}

func (p *Paragraph) String() string {
	sb := strings.Builder{}
	for _, c := range p.Children {
		switch {
		case c.Link != nil:
			id := c.Link.ID
			text := c.Link.Run.InstrText
			link, err := p.file.ReferTarget(id)
			sb.WriteString(text)
			sb.WriteByte('(')
			if err != nil {
				sb.WriteString(id)
			} else {
				sb.WriteString(link)
			}
			sb.WriteByte(')')
		case c.Run != nil:
			sb.WriteString("run") //TODO: implement
		case c.Properties != nil:
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
		switch {
		case c.Link != nil:
			err = e.Encode(c.Link)
		case c.Run != nil:
			err = e.Encode(c.Run)
		case c.Properties != nil:
			err = e.Encode(c.Properties)
		default:
			continue
		}
		if err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	children := make([]ParagraphChild, 0, 64)
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
			var elem ParagraphChild
			switch tt.Name.Local {
			case "hyperlink":
				var value Hyperlink
				d.DecodeElement(&value, &tt)
				id := getAtt(tt.Attr, "id")
				anchor := getAtt(tt.Attr, "anchor")
				if id != "" {
					value.ID = id
				}
				if anchor != "" {
					value.ID = anchor
				}
				elem.Link = &value
			case "r":
				var value Run
				d.DecodeElement(&value, &tt)
				elem.Run = &value
			case "rPr":
				var value RunProperties
				d.DecodeElement(&value, &tt)
				elem.Properties = &value
			case "pPr":
				var value ParagraphProperties
				d.DecodeElement(&value, &tt)
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
