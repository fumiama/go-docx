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
	if m != "62c753dc14365fce007fc4c7c3bd0c82" {
		t.Fatal("real md5:", m)
	}
}
