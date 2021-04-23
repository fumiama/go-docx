package docxlib

import (
	"archive/zip"
	"io"
)

type Docx struct {
	Document    Document
	DocRelation Relationships

	rId int
}

// New generates a new empty docx file that we can manipulate and
// later on, save
func New() *Docx {
	return emptyFile()
}

// Parse generates a new docx file in memory from a reader
func Parse(reader io.ReaderAt, size int64) (doc *Docx, err error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}
	doc, err = unpack(zipReader)
	return
}

// Write allows to save a docx to a writer
func (f *Docx) Write(writer io.Writer) (err error) {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	return f.pack(zipWriter)
}
