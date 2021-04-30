package docxlib

import (
	"archive/zip"
	"errors"
	"io"
)

// DocxLib is the structure that allow to access the internal represntation
// in memory of the doc (either read or about to be written)
type DocxLib struct {
	Document    Document
	DocRelation Relationships

	rId int
}

// New generates a new empty docx file that we can manipulate and
// later on, save
func New() *DocxLib {
	return emptyFile()
}

// Parse generates a new docx file in memory from a reader
func Parse(reader io.ReaderAt, size int64) (doc *DocxLib, err error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}
	doc, err = unpack(zipReader)
	return
}

// Write allows to save a docx to a writer
func (f *DocxLib) Write(writer io.Writer) (err error) {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	return f.pack(zipWriter)
}

// References gets the url for a reference
func (f *DocxLib) References(id string) (href string, err error) {
	for _, a := range f.DocRelation.Relationships {
		if a.ID == id {
			href = a.Target
			return
		}
	}
	err = errors.New("id not found")
	return
}
