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
