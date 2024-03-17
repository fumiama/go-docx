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

import (
	"encoding/xml"
	"io"
	"reflect"
	"strings"
)

// Tabs ...
type Tabs struct {
	XMLName xml.Name `xml:"w:tabs,omitempty"`
	Tabs    []*Tab
}

// UnmarshalXML ...
func (tb *Tabs) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "tab" {
				var value Tab
				err := d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				tb.Tabs = append(tb.Tabs, &value)
			}
		}
	}

	return nil
}

// Tab is the literal tab
type Tab struct {
	XMLName  xml.Name `xml:"w:tab,omitempty"`
	Val      string   `xml:"w:val,attr,omitempty"`
	Position int      `xml:"w:pos,attr,omitempty"`
}

// UnmarshalXML ...
func (t *Tab) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			t.Val = attr.Value
		case "pos":
			if attr.Value == "" {
				continue
			}
			t.Position, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// BarterRabbet is <br> , if with type=page , add pagebreaks
type BarterRabbet struct {
	XMLName xml.Name `xml:"w:br,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML ...
func (f *BarterRabbet) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "type" {
			f.Type = attr.Value
			break
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}

// Text object contains the actual text
type Text struct {
	XMLName xml.Name `xml:"w:t,omitempty"`

	XMLSpace string `xml:"xml:space,attr,omitempty"`

	Text string `xml:",chardata"`
}

// UnmarshalXML ...
func (r *Text) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "space":
			r.XMLSpace = attr.Value
		default:
			// ignore other attributes
		}
	}
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.CharData); ok {
			r.Text = string(tt) // implicitly copy
		}
	}

	return nil
}

// RunMergeRule compares two runs and decides whether they can be merged
type RunMergeRule func(r1, r2 *Run) bool

// MergeAllRuns ...
func MergeAllRuns(_, _ *Run) bool {
	return true
}

// MergeSamePropRuns merges runs with the same properties
func MergeSamePropRuns(r1, r2 *Run) bool {
	if r1 == nil || r2 == nil {
		return false
	}
	if r1.RunProperties == r2.RunProperties {
		return true
	}
	if r1.RunProperties == nil && r2.RunProperties != nil {
		return false
	}
	if r1.RunProperties != nil && r2.RunProperties == nil {
		return false
	}
	rr1 := reflect.ValueOf(r1.RunProperties).Elem()
	rr2 := reflect.ValueOf(r2.RunProperties).Elem()
	for i := 1; i < rr1.NumField(); i++ {
		x1 := rr1.Field(i)
		x2 := rr2.Field(i)
		if x1.IsZero() && x2.IsZero() {
			continue
		}
		if x1.IsZero() && !x2.IsZero() {
			return false
		}
		if !x1.IsZero() && x2.IsZero() {
			return false
		}
		xx1 := x1.Elem()
		if xx1.NumField() <= 1 {
			continue
		}
		xx2 := x2.Elem()
		for j := 1; j < xx1.NumField(); j++ {
			if !xx1.Field(j).Equal(xx2.Field(j)) {
				return false
			}
		}
	}
	return true
}

// MergeSamePropRunsOf merges runs with the same properties of names
func MergeSamePropRunsOf(name ...string) RunMergeRule {
	return func(r1, r2 *Run) bool {
		if r1 == nil || r2 == nil {
			return false
		}
		if r1.RunProperties == r2.RunProperties {
			return true
		}
		if r1.RunProperties == nil && r2.RunProperties != nil {
			return false
		}
		if r1.RunProperties != nil && r2.RunProperties == nil {
			return false
		}
		rr1 := reflect.ValueOf(r1.RunProperties).Elem()
		rr2 := reflect.ValueOf(r2.RunProperties).Elem()
		for _, n := range name {
			x1 := rr1.FieldByName(n)
			x2 := rr2.FieldByName(n)
			if x1.IsZero() && x2.IsZero() {
				continue
			}
			if x1.IsZero() && !x2.IsZero() {
				return false
			}
			if !x1.IsZero() && x2.IsZero() {
				return false
			}
			xx1 := x1.Elem()
			if xx1.NumField() <= 1 {
				continue
			}
			xx2 := x2.Elem()
			for j := 1; j < xx1.NumField(); j++ {
				if !xx1.Field(j).Equal(xx2.Field(j)) {
					return false
				}
			}
		}
		return true
	}
}

// MergeText will merge contiguous run texts in a paragraph into one run
//
//	note: np is not a deep-copy
func (p *Paragraph) MergeText(canmerge RunMergeRule) (np Paragraph) {
	var prevrun *Run
	np = *p
	np.Children = make([]interface{}, 0, 64)
	for _, c := range p.Children {
		switch o := c.(type) {
		case *Run:
			r := *o
			r.Children = make([]interface{}, 0, 16)
			t := &Text{}
			for _, c := range o.Children {
				switch x := c.(type) {
				case *Text:
					if x.Text != "" {
						t.Text += x.Text
					}
				default:
					if t.Text != "" {
						r.Children = append(r.Children, t)
						t = &Text{}
					}
					r.Children = append(r.Children, x)
				}
			}
			if t.Text != "" {
				r.Children = append(r.Children, t)
			}
			if prevrun != nil && canmerge(prevrun, &r) {
				var prevtext *Text
				noappend := false
				if len(prevrun.Children) == 0 {
					prevtext = &Text{}
				} else {
					i := len(prevrun.Children) - 1
					if t, ok := prevrun.Children[i].(*Text); ok {
						prevtext = t
						noappend = true
					} else {
						prevtext = &Text{}
					}
				}
				for _, c := range r.Children {
					switch x := c.(type) {
					case *Text:
						if x.Text != "" {
							prevtext.Text += x.Text
						}
					default:
						if prevtext.Text != "" {
							if noappend {
								noappend = false
							} else {
								prevrun.Children = append(prevrun.Children, t)
							}
							prevtext = &Text{}
						}
						prevrun.Children = append(prevrun.Children, x)
					}
				}
				if prevtext.Text != "" && !noappend {
					prevrun.Children = append(prevrun.Children, t)
				}
			} else {
				prevrun = &r
				np.Children = append(np.Children, &r)
			}
		default:
			np.Children = append(np.Children, o)
		}
	}
	return
}
