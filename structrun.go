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
	"strings"
)

// Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName xml.Name `xml:"w:r,omitempty"`

	RunProperties *RunProperties `xml:"w:rPr,omitempty"`

	InstrText string `xml:"w:instrText,omitempty"`

	Children []interface{}
}

// UnmarshalXML ...
func (r *Run) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		var child interface{}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.RunProperties = &value
				continue
			case "instrText":
				var value string
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.InstrText = value
				continue
			case "t":
				var value Text
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				child = &value
			case "drawing":
				var value Drawing
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				child = &value
			case "tab":
				child = &WTab{}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
			r.Children = append(r.Children, child)
		}
	}

	return nil
}

// WTab is the literal tab
type WTab struct {
	XMLName xml.Name `xml:"w:tab,omitempty"`
}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName  xml.Name `xml:"w:rPr,omitempty"`
	Color    *Color
	Size     *Size
	RunStyle *RunStyle
	Style    *Style
	Shade    *Shade
}

// UnmarshalXML ...
func (r *RunProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "color":
				var value Color
				value.Val = getAtt(tt.Attr, "val")
				r.Color = &value
			case "sz":
				var value Size
				value.Val = getAtt(tt.Attr, "val")
				r.Size = &value
			case "rStyle":
				var value RunStyle
				value.Val = getAtt(tt.Attr, "val")
				r.RunStyle = &value
			case "pStyle":
				var value Style
				value.Val = getAtt(tt.Attr, "val")
				r.Style = &value
			case "shd":
				var value Shade
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Shade = &value
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}

	return nil
}
