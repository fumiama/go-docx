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

//nolint:revive,stylecheck
const (
	XMLNS_W   = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R   = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
	XMLNS_WP  = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
	XMLNS_WPS = `http://schemas.microsoft.com/office/word/2010/wordprocessingShape`
	XMLNS_WPC = `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas`
	XMLNS_WPG = `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup`
	// XMLNS_MC  = `http://schemas.openxmlformats.org/markup-compatibility/2006`
	// XMLNS_WP14 = `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`

	XMLNS_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
)

func getAtt(atts []xml.Attr, name string) string {
	for _, at := range atts {
		if at.Name.Local == name {
			return at.Value
		}
	}
	return ""
}

// Body <w:body>
type Body struct {
	Items []interface{}

	file *Docx
}

// UnmarshalXML ...
func (b *Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "p":
				var value Paragraph
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.file = b.file
				b.Items = append(b.Items, &value)
			case "tbl":
				var value WTable
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.file = b.file
				b.Items = append(b.Items, &value)
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

// Document <w:document>
type Document struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	XMLW    string   `xml:"xmlns:w,attr"`             // cannot be unmarshalled in
	XMLR    string   `xml:"xmlns:r,attr,omitempty"`   // cannot be unmarshalled in
	XMLWP   string   `xml:"xmlns:wp,attr,omitempty"`  // cannot be unmarshalled in
	XMLWPS  string   `xml:"xmlns:wps,attr,omitempty"` // cannot be unmarshalled in
	XMLWPC  string   `xml:"xmlns:wpc,attr,omitempty"` // cannot be unmarshalled in
	XMLWPG  string   `xml:"xmlns:wpg,attr,omitempty"` // cannot be unmarshalled in
	// XMLMC   string   `xml:"xmlns:mc,attr,omitempty"`  // cannot be unmarshalled in
	// XMLWP14 string   `xml:"xmlns:wp14,attr,omitempty"` // cannot be unmarshalled in

	Body Body `xml:"w:body"`
}

// UnmarshalXML ...
func (doc *Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "body" {
				err = d.DecodeElement(&doc.Body, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}
	}
	return nil
}
