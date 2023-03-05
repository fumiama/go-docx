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
	"strconv"
	"strings"
)

// Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName xml.Name `xml:"w:r,omitempty"`
	RsidRPr string   `xml:"w:rsidRPr,attr,omitempty"`

	RunProperties *RunProperties `xml:"w:rPr,omitempty"`

	InstrText string `xml:"w:instrText,omitempty"`

	Children []interface{}

	file *Docx
}

// UnmarshalXML ...
func (r *Run) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidRPr":
			r.RsidRPr = attr.Value
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

		if tt, ok := t.(xml.StartElement); ok {
			child, err := r.parse(d, tt)
			if err != nil {
				return err
			}
			if child != nil {
				r.Children = append(r.Children, child)
			}
		}
	}

	return nil
}

func (r *Run) parse(d *xml.Decoder, tt xml.StartElement) (child interface{}, err error) {
	switch tt.Name.Local {
	case "rPr":
		var value RunProperties
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		r.RunProperties = &value
		return nil, nil
	case "instrText":
		var value string
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		r.InstrText = value
		return nil, nil
	case "t":
		var value Text
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		child = &value
	case "drawing":
		var value Drawing
		value.file = r.file
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		child = &value
	case "tab":
		child = &Tab{}
	case "AlternateContent":
	altcont:
		for {
			tok, err1 := d.Token()
			if err1 == io.EOF {
				break
			}
			if err1 != nil {
				return nil, err1
			}

			if ttt, ok := tok.(xml.StartElement); ok && ttt.Name.Local == "Choice" {
				for _, attr := range ttt.Attr {
					if attr.Name.Local == "Requires" {
						if attr.Value == "wps" || attr.Value == "wpc" || attr.Value == "wpg" {
							tok, err = d.Token() // go into choice
							if err != nil {
								return nil, err
							}
							if ttt, ok := tok.(xml.StartElement); ok {
								child, err = r.parse(d, ttt)
							}
							break altcont
						}
						break
					}
				}
			}
			if et, ok := tok.(xml.EndElement); ok {
				if et.Name.Local == "AlternateContent" {
					break
				}
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return nil, err
			}
		}
	default:
		err = d.Skip() // skip unsupported tags
	}
	return
}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName   xml.Name `xml:"w:rPr,omitempty"`
	Fonts     *RunFonts
	Bold      *Bold
	ICs       *struct{} `xml:"w:iCs,omitempty"`
	Italic    *Italic
	Underline *Underline
	Highlight *Highlight
	Color     *Color
	Size      *Size
	SizeCs    *SizeCs
	RunStyle  *RunStyle
	Style     *Style
	Shade     *Shade
	Kern      *Kern
	VertAlign *VertAlign
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
			case "rFonts":
				var value RunFonts
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Fonts = &value
			case "b":
				r.Bold = &Bold{}
			case "iCs":
				r.ICs = &struct{}{}
			case "i":
				r.Italic = &Italic{}
			case "u":
				var value Underline
				value.Val = getAtt(tt.Attr, "val")
				r.Underline = &value
			case "highlight":
				var value Highlight
				value.Val = getAtt(tt.Attr, "val")
				r.Highlight = &value
			case "color":
				var value Color
				value.Val = getAtt(tt.Attr, "val")
				r.Color = &value
			case "sz":
				var value Size
				value.Val = getAtt(tt.Attr, "val")
				r.Size = &value
			case "szCs":
				var value SizeCs
				value.Val = getAtt(tt.Attr, "val")
				r.SizeCs = &value
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
			case "kern":
				var value Kern
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = strconv.ParseInt(v, 10, 64)
				if err != nil {
					return err
				}
				r.Kern = &value
			case "vertAlign":
				var value VertAlign
				value.Val = getAtt(tt.Attr, "val")
				r.VertAlign = &value
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

// RunFonts specifies the fonts used in the text of a run.
type RunFonts struct {
	XMLName xml.Name `xml:"w:rFonts,omitempty"`
	ASCII   string   `xml:"w:ascii,attr,omitempty"`
	HAnsi   string   `xml:"w:hAnsi,attr,omitempty"`
	Hint    string   `xml:"w:hint,attr,omitempty"`
}

// UnmarshalXML ...
func (f *RunFonts) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "ascii":
			f.ASCII = attr.Value
		case "hAnsi":
			f.HAnsi = attr.Value
		case "hint":
			f.Hint = attr.Value
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}
