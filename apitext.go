package docxlib

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
