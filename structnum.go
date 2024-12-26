/*
   Copyright (c) 2024 l0g1n

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
)

// NumProperties show the number properties
type NumProperties struct {
	XMLName xml.Name `xml:"w:numPr,omitempty"`
	NumID   *NumID
	Ilvl    *Ilevel
}

// NumID show the number id
type NumID struct {
	XMLName xml.Name `xml:"w:numId,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Ilevel show the level
type Ilevel struct {
	XMLName xml.Name `xml:"w:ilvl,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// UnmarshalXML ...
func (n *NumProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "numId":
				var value NumID
				value.Val = getAtt(tt.Attr, "val")
				n.NumID = &value
			case "ilvl":
				var value Ilevel
				value.Val = getAtt(tt.Attr, "val")
				n.Ilvl = &value
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
