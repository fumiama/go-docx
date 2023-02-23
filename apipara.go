package docxlib

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	f.Document.Body.mu.Lock()
	defer f.Document.Body.mu.Unlock()
	f.Document.Body.Paragraphs = append(f.Document.Body.Paragraphs, Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     f,
	})

	return &f.Document.Body.Paragraphs[len(f.Document.Body.Paragraphs)-1]
}

// Justification allows to set para's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (p *Paragraph) Justification(val string) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	p.Properties.Justification = &Justification{Val: val}
	return p
}
