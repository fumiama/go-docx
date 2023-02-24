package docxlib

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
			XMLW:  XMLNS_W,
			XMLR:  XMLNS_R,
			XMLWP: XMLNS_WP,
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
