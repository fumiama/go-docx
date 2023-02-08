package docxlib

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	p := &Paragraph{
		Children: make([]ParagraphChild, 0, 64),
		file:     f,
	}

	f.Document.Body.Paragraphs = append(f.Document.Body.Paragraphs, p)

	return p
}
