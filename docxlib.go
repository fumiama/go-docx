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

// Package docxlib is one of the most functional libraries to read and write .docx
// (a.k.a. Microsoft Word documents or ECMA-376 Office Open XML) files in Go.
package docxlib

import (
	"archive/zip"
	"bytes"
	"io"
	"io/fs"
)

// Docx is the structure that allow to access the internal represntation
// in memory of the doc (either read or about to be written)
type Docx struct {
	Document Document // Document is word/document.xml

	docRelation Relationships // docRelation is word/_rels/document.xml.rels

	media        []Media
	mediaNameIdx map[string]int

	rID     uintptr
	imageID uintptr

	template string
	tmplfs   fs.FS
	tmpfslst []string

	buf        *bytes.Buffer
	isbufempty bool

	io.Reader
	io.WriterTo
}

// NewA4 generates a new empty A4 docx file that we can manipulate and
// later on, save
func NewA4() *Docx {
	return newEmptyA4File()
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

// WriteTo allows to save a docx to a writer
func (f *Docx) WriteTo(writer io.Writer) (_ int64, err error) {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	return 0, f.pack(zipWriter)
}

// Read allows to save a docx to buf
func (f *Docx) Read(p []byte) (n int, err error) {
	if !f.isbufempty {
		n, err = f.buf.Read(p)
		if err == io.EOF {
			f.buf.Reset()
			f.isbufempty = true
			return
		}
	}
	zipWriter := zip.NewWriter(f.buf)
	defer zipWriter.Close()
	f.isbufempty = false
	return f.buf.Read(p)
}
