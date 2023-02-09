package docxlib

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"os"
)

// This receives a zip file writer (word documents are a zip with multiple xml inside)
// and writes the relevant files. Some of them come from the empty_constants file,
// others from the actual in-memory structure
func (f *Docx) pack(zipWriter *zip.Writer) (err error) {
	fileslst := []string{
		"_rels/.rels",
		"docProps/app.xml",
		"docProps/core.xml",
		"word/theme/theme1.xml",
		"word/styles.xml",
		"[Content_Types].xml",
	}
	files := make(map[string]io.Reader, 64)

	for _, name := range fileslst {
		files[name], err = TEMP_XML_FS.Open("xml/" + name)
		if err != nil {
			return
		}
	}
	files["word/_rels/document.xml.rels"] = marshaller{data: f.DocRelation}
	files["word/document.xml"] = marshaller{data: f.Document}

	for path, r := range files {
		w, err := zipWriter.Create(path)
		if err != nil {
			return err
		}

		_, err = io.Copy(w, r)
		if err != nil {
			return err
		}
	}

	return
}

type marshaller struct {
	data interface{}
	io.Reader
	io.WriterTo
}

// Read is fake and is to trigger io.WriterTo
func (m marshaller) Read(p []byte) (n int, err error) {
	return 0, os.ErrInvalid
}

// WriteTo n is always 0 for we don't care that value
func (m marshaller) WriteTo(w io.Writer) (n int64, err error) {
	_, err = io.WriteString(w, xml.Header)
	if err != nil {
		return
	}
	err = xml.NewEncoder(w).Encode(m.data)
	return
}
