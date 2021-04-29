package docxlib

// AddParagraph adds a new paragraph
func (f *DocxLib) AddParagraph() *Paragraph {
	p := &Paragraph{
		Data: make([]ParagraphChild, 0),
		file: f,
	}

	f.Document.Body.Paragraphs = append(f.Document.Body.Paragraphs, p)
	return p
}

func (f *DocxLib) Paragraphs() []*Paragraph {
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
