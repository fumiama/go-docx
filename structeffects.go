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

import "encoding/xml"

// RunStyle contains styling for a run
type RunStyle struct {
	XMLName xml.Name `xml:"w:rStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Style contains styling for a paragraph
type Style struct {
	XMLName xml.Name `xml:"w:pStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Color contains the sound of music. :D
// I'm kidding. It contains the color
type Color struct {
	XMLName xml.Name `xml:"w:color,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Size contains the font size
type Size struct {
	XMLName xml.Name `xml:"w:sz,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// SizeCs contains the cs font size
type SizeCs struct {
	XMLName xml.Name `xml:"w:szCs,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Bold ...
type Bold struct {
	XMLName xml.Name `xml:"w:b,omitempty"`
}

// Italic ...
type Italic struct {
	XMLName xml.Name `xml:"w:i,omitempty"`
}

// Underline ...
type Underline struct {
	XMLName xml.Name `xml:"w:u,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Highlight ...
type Highlight struct {
	XMLName xml.Name `xml:"w:highlight,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Kern ...
type Kern struct {
	XMLName xml.Name `xml:"w:kern,omitempty"`
	Val     int64    `xml:"w:val,attr"`
}

// Justification contains the way of the horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
type Justification struct {
	XMLName xml.Name `xml:"w:jc,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// TextAlignment ...
type TextAlignment struct {
	XMLName xml.Name `xml:"w:textAlignment,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// VertAlign ...
type VertAlign struct {
	XMLName xml.Name `xml:"w:vertAlign,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Shade is an element that represents a shading pattern applied to a document element.
type Shade struct {
	XMLName       xml.Name `xml:"w:shd,omitempty"`
	Val           string   `xml:"w:val,attr,omitempty"`
	Color         string   `xml:"w:color,attr,omitempty"`
	Fill          string   `xml:"w:fill,attr,omitempty"`
	ThemeFill     string   `xml:"w:themeFill,attr,omitempty"`
	ThemeFillTint string   `xml:"w:themeFillTint,attr,omitempty"`
}

// UnmarshalXML ...
func (s *Shade) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "color":
			s.Color = attr.Value
		case "fill":
			s.Fill = attr.Value
		case "themeFill":
			s.ThemeFill = attr.Value
		case "themeFillTint":
			s.ThemeFillTint = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}

// AdjustRightInd ...
type AdjustRightInd struct {
	XMLName xml.Name `xml:"w:adjustRightInd,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

// SnapToGrid ...
type SnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

// Kinsoku ...
type Kinsoku struct {
	XMLName xml.Name `xml:"w:kinsoku,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

// OverflowPunct ...
type OverflowPunct struct {
	XMLName xml.Name `xml:"w:overflowPunct,omitempty"`
	Val     int      `xml:"w:val,attr"`
}
