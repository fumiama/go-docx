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

// WordprocessingCanvas ...
type WordprocessingCanvas struct {
	XMLName    xml.Name `xml:"wpc:wpc,omitempty"`
	Background *WPCBackground
	Whole      *WPCWhole

	Items []interface{}

	file *Docx
}

// UnmarshalXML ...
func (c *WordprocessingCanvas) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) (err error) {
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
			case "bg":
				c.Background = new(WPCBackground)
				err = d.DecodeElement(c.Background, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "whole":
				c.Whole = new(WPCWhole)
				err = d.DecodeElement(c.Whole, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "wsp":
				var value WordprocessingShape
				value.file = c.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Items = append(c.Items, &value)
			case "pic":
				var value Picture
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.XMLPIC = getAtt(tt.Attr, "pic")
				c.Items = append(c.Items, &value)
			case "wgp":
				var value WordprocessingGroup
				value.file = c.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Items = append(c.Items, &value)
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

// WPCBackground ...
type WPCBackground struct {
	XMLName xml.Name  `xml:"wpc:bg,omitempty"`
	NoFill  *struct{} `xml:"a:noFill,omitempty"`
}

// UnmarshalXML ...
func (b *WPCBackground) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) (err error) {
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
			case "noFill":
				b.NoFill = &struct{}{}
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

// WPCWhole ...
type WPCWhole struct {
	XMLName xml.Name `xml:"wpc:whole,omitempty"`
	Line    *ALine
}

// UnmarshalXML ...
func (w *WPCWhole) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) (err error) {
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
			case "ln":
				w.Line = new(ALine)
				err = d.DecodeElement(w.Line, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
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
