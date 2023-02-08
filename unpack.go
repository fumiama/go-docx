package docxlib

// This contains internal functions needed to unpack (read) a zip file
import (
	"archive/zip"
	"encoding/xml"
	"io"

	"github.com/golang/glog"
)

// This receives a zip file (word documents are a zip with multiple xml inside)
// and parses the files that are relevant for us:
// 1.-Document
// 2.-Relationships
func unpack(zipReader *zip.Reader) (docx *Docx, err error) {
	docx = new(Docx)
	for _, f := range zipReader.File {
		if f.Name == "word/_rels/document.xml.rels" {
			err = processRelations(f, &docx.DocRelation)
			if err != nil {
				return
			}
		}
		if f.Name == "word/document.xml" {
			err = processDoc(f, &docx.Document)
			if err != nil {
				return
			}
		}
	}
	return
}

// Processes one of the relevant files, the one with the actual document
func processDoc(file *zip.File, doc *Document) error {
	filebytes, err := readZipFile(file)
	if err != nil {
		glog.Errorln("Error reading from internal zip file")
		return err
	}
	glog.V(0).Infoln("Doc:", string(filebytes))

	doc.XMLW = XMLNS_W
	doc.XMLR = XMLNS_R
	doc.XMLName.Space = XMLNS_W
	doc.XMLName.Local = "document"
	err = xml.Unmarshal(filebytes, doc)
	if err != nil {
		glog.Errorln("Error unmarshalling doc", string(filebytes))
		return err
	}
	glog.V(0).Infoln("Paragraph", doc.Body.Paragraphs)
	return nil
}

// Processes one of the relevant files, the one with the relationships
func processRelations(file *zip.File, rels *Relationships) error {
	filebytes, err := readZipFile(file)
	if err != nil {
		glog.Errorln("Error reading from internal zip file")
		return err
	}
	glog.V(0).Infoln("Relations:", string(filebytes))

	rels.Xmlns = XMLNS_R
	err = xml.Unmarshal(filebytes, rels)
	if err != nil {
		glog.Errorln("Error unmarshalling relationships")
		return err
	}
	return nil
}

// From a zip file structure, we return a byte array
func readZipFile(zf *zip.File) ([]byte, error) {
	f, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return io.ReadAll(f)
}
