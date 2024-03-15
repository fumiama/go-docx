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

// Package docx is one of the most functional libraries to read and write .docx
// (a.k.a. Microsoft Word documents or ECMA-376 Office Open XML) files in Go.
package docx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"io/fs"
	"os"
	"sync"
)

// Docx is the structure that allow to access the internal represntation
// in memory of the doc (either read or about to be written)
type Docx struct {
	Document Document // Document is word/document.xml

	docRelation Relationships // docRelation is word/_rels/document.xml.rels

	media        []Media
	mediaNameIdx map[string]int

	rID       uintptr
	imageID   uintptr
	docID     uintptr
	slowIDs   map[string]uintptr
	slowIDsMu sync.Mutex

	template string
	tmplfs   fs.FS
	tmpfslst []string

	io.Reader
	io.WriterTo
}

// New generates a new empty docx file that we can manipulate and
// later on, save
func New() *Docx {
	return newEmptyFile()
}

// Parse generates a new docx file in memory from a reader
// You can it invoke from a file
//
//	readFile, err := os.Open(FILE_PATH)
//	if err != nil {
//		panic(err)
//	}
//	fileinfo, err := readFile.Stat()
//	if err != nil {
//		panic(err)
//	}
//	size := fileinfo.Size()
//	doc, err := docxlib.Parse(readFile, int64(size))
//
// but also you can invoke from a webform (BEWARE of trusting users data!!!)
//
//	func uploadFile(w http.ResponseWriter, r *http.Request) {
//		r.ParseMultipartForm(10 << 20)
//
//		file, handler, err := r.FormFile("file")
//		if err != nil {
//			fmt.Println("Error Retrieving the File")
//			fmt.Println(err)
//			http.Error(w, err.Error(), http.StatusBadRequest)
//			return
//		}
//		defer file.Close()
//		docxlib.Parse(file, handler.Size)
//	}
func Parse(reader io.ReaderAt, size int64) (doc *Docx, err error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}
	doc, err = unpack(zipReader)
	return
}

// LoadBodyItems will load body and media to a new Docx struct.
// You should call UseTemplate to set a template later.
func LoadBodyItems(items []interface{}, media []Media) *Docx {
	doc := &Docx{
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
			Body:   Body{Items: items},
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
		media:        media,
		mediaNameIdx: make(map[string]int, 64),
		rID:          3,
		slowIDs:      make(map[string]uintptr, 64),
	}
	doc.Document.Body.file = doc
	for i, m := range media {
		doc.mediaNameIdx[m.Name] = i
	}
	doc.slowIDs["图片"] = uintptr(len(media) + 1)
	return doc
}

// WriteTo allows to save a docx to a writer
func (f *Docx) WriteTo(writer io.Writer) (_ int64, err error) {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	return 0, f.pack(zipWriter)
}

// Read is a fake function and cannot be used
func (f *Docx) Read(_ []byte) (int, error) {
	return 0, os.ErrInvalid
}
