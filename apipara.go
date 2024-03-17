/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	p := &Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     f,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, p)
	return p
}

// AddParagraph adds a new paragraph
func (c *WTableCell) AddParagraph() *Paragraph {
	c.Paragraphs = append(c.Paragraphs, &Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     c.file,
	})

	return c.Paragraphs[len(c.Paragraphs)-1]
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

// AddPageBreaks adds a pagebreaks
func (p *Paragraph) AddPageBreaks() *Run {
	c := make([]interface{}, 1, 64)
	c[0] = &BarterRabbet{
		Type: "page",
	}
	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}
	p.Children = append(p.Children, run)
	return run
}
