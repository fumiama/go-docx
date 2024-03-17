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

// SectPr show the properties of the document, like paper size
type SectPr struct {
	XMLName xml.Name `xml:"w:sectPr,omitempty"` // properties of the document, including paper size
	PgSz    *PgSz    `xml:"w:pgSz,omitempty"`
}

// PgSz show the paper size
type PgSz struct {
	W xml.Attr `xml:"w:w,attr"` // width of paper
	H xml.Attr `xml:"w:h,attr"` // high of paper
}

// UnmarshalXML ...
func (sect *SectPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "pgSz":
				var value PgSz
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgSz = &value
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// UnmarshalXML ...
func (pgsz *PgSz) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Name.Local == "w" {
			pgsz.W = xml.Attr{Name: xml.Name{Local: "w:w"}, Value: attr.Value}
		}
		if attr.Name.Local == "h" {
			pgsz.H = xml.Attr{Name: xml.Name{Local: "w:w"}, Value: attr.Value}
		}
	}
	// Consume the end element
	_, err = d.Token()
	return nil
}
