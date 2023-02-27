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
	"bytes"
	"encoding/xml"
)

func newEmptyA4File() *Docx {
	docx := &Docx{
		Document: Document{
			XMLName: xml.Name{
				Space: "w",
			},
			XMLW:   XMLNS_W,
			XMLR:   XMLNS_R,
			XMLWP:  XMLNS_WP,
			XMLWPS: XMLNS_WPS,
			// XMLMC:  XMLNS_MC,
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
				{
					ID:     "rId4",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings`,
					Target: "settings.xml",
				},
				{
					ID:     "rId5",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings`,
					Target: "webSettings.xml",
				},
			},
		},
		media:        make([]Media, 0, 64),
		mediaNameIdx: make(map[string]int, 64),
		rID:          5,
		slowIDs:      make(map[string]uintptr, 64),
		template:     "a4",
		tmpfslst: []string{
			"_rels/.rels",
			"docProps/app.xml",
			"docProps/core.xml",
			"word/theme/theme1.xml",
			"word/fontTable.xml",
			"word/settings.xml",
			"word/styles.xml",
			"word/webSettings.xml",
			"[Content_Types].xml",
		},
		buf: bytes.NewBuffer(make([]byte, 0, 1024*1024)),
	}
	docx.Document.Body.file = docx
	return docx
}
