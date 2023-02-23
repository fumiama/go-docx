package docxlib

import "unsafe"

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	f.Document.Body.mu.Lock()
	defer f.Document.Body.mu.Unlock()
	f.Document.Body.Items = append(f.Document.Body.Items, Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     f,
	})

	p := f.Document.Body.Items[len(f.Document.Body.Items)-1]

	return *(**Paragraph)(unsafe.Add(unsafe.Pointer(&p), unsafe.Sizeof(uintptr(0))))
}

// AddParagraph adds a new paragraph
func (c *WTableCell) AddParagraph() *Paragraph {
	c.mu.Lock()
	defer c.mu.Unlock()

	c.Paragraphs = append(c.Paragraphs, Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     c.file,
	})

	return &c.Paragraphs[len(c.Paragraphs)-1]
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
