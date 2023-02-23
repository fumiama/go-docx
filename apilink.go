package docxlib

//nolint:revive,stylecheck
const (
	HYPERLINK_STYLE = "a3"
)

// AddLink adds an hyperlink to paragraph
func (p *Paragraph) AddLink(text string, link string) *Hyperlink {
	rid := p.file.addLinkRelation(link)
	hyperlink := &Hyperlink{
		ID: rid,
		Run: Run{
			RunProperties: &RunProperties{
				RunStyle: &RunStyle{
					Val: HYPERLINK_STYLE,
				},
			},
			InstrText: text,
		},
	}

	p.Children = append(p.Children, hyperlink)

	return hyperlink
}
