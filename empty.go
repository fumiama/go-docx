package docxlib

import (
	"bytes"
	"encoding/xml"
)

func newEmptyFile() *Docx {
	return &Docx{
		Document: Document{
			XMLName: xml.Name{
				Space: "w",
			},
			XMLW:    XMLNS_W,
			XMLR:    XMLNS_R,
			XMLWP:   XMLNS_WP,
			XMLWP14: XMLNS_WP14,
			Body: &Body{
				XMLName: xml.Name{
					Space: "w",
				},
				Paragraphs: make([]*Paragraph, 0, 64),
			},
		},
		DocRelation: Relationships{
			Xmlns: XMLNS,
			Relationships: []*Relationship{
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
		rId: 3,
		buf: bytes.NewBuffer(make([]byte, 0, 1024*1024*4)),
	}
}
