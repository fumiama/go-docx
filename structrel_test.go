package docxlib

import (
	"crypto/md5"
	"encoding/hex"
	"io"
	"os"
	"testing"
)

func TestRelationships(t *testing.T) {
	rel := Relationships{
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
	}
	f, err := os.Create("TestRelationships.xml")
	if err != nil {
		t.Fatal(err)
	}
	h := md5.New()
	_, err = io.Copy(io.MultiWriter(f, h), marshaller{data: &rel})
	if err != nil {
		t.Fatal(err)
	}
	m := hex.EncodeToString(h.Sum(make([]byte, 0, 16)))
	if m != "c75af73ef6cc9536a193669c4a3605c3" {
		t.Fatal("real md5:", m)
	}
}
