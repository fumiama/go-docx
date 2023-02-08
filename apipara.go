package docxlib

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	p := &Paragraph{
		Data: make([]ParagraphChild, 0, 64),
		file: f,
	}

	f.Document.Body.Paragraphs = append(f.Document.Body.Paragraphs, p)

	return p
}

func (f *Docx) Paragraphs() []*Paragraph {
	return f.Document.Body.Paragraphs
}

func (p *Paragraph) Children() (ret []ParagraphChild) {
	return p.Data
}
