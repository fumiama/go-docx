package docxlib

import (
	"strconv"
	"sync/atomic"
)

const (
	HYPERLINK_STYLE = "a1"
)

// when adding an hyperlink we need to store a reference in the relationship field
func (f *Docx) addLinkRelation(link string) string {
	rel := &Relationship{
		ID:         "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rId, 1))),
		Type:       REL_HYPERLINK,
		Target:     link,
		TargetMode: REL_TARGETMODE,
	}

	f.DocRelation.Relationships = append(f.DocRelation.Relationships, rel)

	return rel.ID
}

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
