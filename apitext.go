package docxlib

import "encoding/xml"

// AddTab adds tab to para
func (p *Paragraph) AddTab() *Run {
	run := &Run{
		RunProperties: &RunProperties{},
		FrontTab: []struct {
			XMLName xml.Name "xml:\"w:tab,omitempty\""
		}{{}},
	}
	p.Children = append(p.Children, ParagraphChild{Run: run})
	return run
}

// AddText adds text to paragraph
func (p *Paragraph) AddText(text string) *Run {
	if text == "\t" {
		return p.AddTab()
	}

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
