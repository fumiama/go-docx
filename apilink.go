package docxlib

const (
	HYPERLINK_STYLE = "a1"
)

// AddLink adds an hyperlink to paragraph
func (p *Paragraph) AddLink(text string, link string) *Hyperlink {
	rId := p.file.addLinkRelation(link)
	hyperlink := &Hyperlink{
		ID: rId,
		Run: Run{
			RunProperties: &RunProperties{
				RunStyle: &RunStyle{
					Val: HYPERLINK_STYLE,
				},
			},
			InstrText: text,
		},
	}

	p.Children = append(p.Children, ParagraphChild{Link: hyperlink})

	return hyperlink
}
