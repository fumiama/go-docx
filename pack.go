package docxlib

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
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
	files := make(map[string][]byte, 64)

	for _, name := range fileslst {
		files[name], err = TEMP_XML_FS.ReadFile("xml/" + name)
		if err != nil {
			return
		}
	}
	files["word/_rels/document.xml.rels"], err = marshal(f.DocRelation)
	if err != nil {
		return
	}
	files["word/document.xml"], err = marshal(f.Document)
	if err != nil {
		return
	}

	for path, data := range files {
		w, err := zipWriter.Create(path)
		if err != nil {
			return err
		}

		_, err = w.Write(data)
		if err != nil {
			return err
		}
	}

	return
}

func marshal(data interface{}) (out []byte, err error) {
	buf := bytes.NewBuffer(make([]byte, 0, 1024))
	buf.WriteString(xml.Header)
	err = xml.NewEncoder(buf).Encode(data)
	if err != nil {
		return
	}
	out = buf.Bytes()
	return
}
