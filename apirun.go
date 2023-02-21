package docxlib

import "encoding/xml"

// Color allows to set run color
func (r *Run) Color(color string) *Run {
	r.RunProperties.Color = &Color{
		Val: color,
	}

	return r
}

// Size allows to set run size
func (r *Run) Size(size string) *Run {
	r.RunProperties.Size = &Size{
		Val: size,
	}

	return r
}

// Justification allows to set run's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (r *Run) Justification(val string) *Run {
	r.RunProperties.Justification = &Justification{
		Val: val,
	}

	return r
}

// AddTab add a tab in front of the run
func (r *Run) AddTab() *Run {
	r.FrontTab = append(r.FrontTab, struct {
		XMLName xml.Name "xml:\"w:tab,omitempty\""
	}{})
	return r
}

// AppendTab add a tab after the run
func (r *Run) AppendTab() *Run {
	r.RearTab = append(r.RearTab, struct {
		XMLName xml.Name "xml:\"w:tab,omitempty\""
	}{})
	return r
}
