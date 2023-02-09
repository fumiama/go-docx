package docxlib

import "strings"

// AddText adds text to paragraph
func (p *Paragraph) AddText(text string) *Run {
	t := &Text{
		Text: text,
	}

	run := &Run{
		Text:          t,
		RunProperties: &RunProperties{},
	}

	p.Children = append(p.Children, ParagraphChild{Run: run})

	return run
}

func (p *Paragraph) String() string {
	sb := strings.Builder{}
	for _, c := range p.Children {
		switch {
		case c.Link != nil:
			id := c.Link.ID
			text := c.Link.Run.InstrText
			link, err := p.file.Refer(id)
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
