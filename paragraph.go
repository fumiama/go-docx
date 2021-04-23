package docxlib

import (
	"encoding/xml"
)

type ParagraphChild struct {
	Link *Hyperlink `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink"`
	Run  *Run       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	Data    []ParagraphChild

	file *Docx
}

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	p := &Paragraph{
		Data: make([]ParagraphChild, 0),
		file: f,
	}

	f.Document.Body.Paragraphs = append(f.Document.Body.Paragraphs, p)
	return p
}

func (f *Docx) Paragraphs() []*Paragraph {
	return f.Document.Body.Paragraphs
}

func (p *Paragraph) Runs() (ret []*Run) {
	data := p.Data
	for _, d := range data {
		if d.Run != nil {
			ret = append(ret, d.Run)
		}
	}
	return
}
