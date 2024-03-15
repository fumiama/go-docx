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
)

func newEmptyFile() *Docx {
	docx := &Docx{
		Document: Document{
			XMLName: xml.Name{
				Space: "w",
			},
			XMLW:   XMLNS_W,
			XMLR:   XMLNS_R,
			XMLWP:  XMLNS_WP,
			XMLWPS: XMLNS_WPS,
			XMLWPC: XMLNS_WPC,
			XMLWPG: XMLNS_WPG,
			// XMLMC:  XMLNS_MC,
			// XMLO:   XMLNS_O,
			// XMLV:   XMLNS_V,
			// XMLWP14: XMLNS_WP14,
			Body: Body{
				Items: make([]interface{}, 0, 64),
			},
		},
		docRelation: Relationships{
			Xmlns: XMLNS_REL,
			Relationship: []Relationship{
				{
					ID:     "rId1",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`,
					Target: "styles.xml",
				},
				{
					ID:     "rId2",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme`,
					Target: "theme/theme1.xml",
				},
				{
					ID:     "rId3",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable`,
					Target: "fontTable.xml",
				},
			},
		},
		media:        make([]Media, 0, 64),
		mediaNameIdx: make(map[string]int, 64),
		rID:          3,
		slowIDs:      make(map[string]uintptr, 64),
	}
	docx.Document.Body.file = docx
	return docx
}
