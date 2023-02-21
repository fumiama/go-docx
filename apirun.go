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
