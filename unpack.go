package docxlib

// This contains internal functions needed to unpack (read) a zip file
import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io/ioutil"
)

// This receives a zip file (word documents are a zip with multiple xml inside)
// and parses the files that are relevant for us:
// 1.-Document
// 2.-Relationships
func unpack(zipReader *zip.Reader) (docx *DocxLib, err error) {
	var doc *Document
	var relations *Relationships
	for _, f := range zipReader.File {
		if f.Name == "word/_rels/document.xml.rels" {
			relations, err = processRelations(f)
			if err != nil {
				return nil, err
			}
		}
		if f.Name == "word/document.xml" {
			doc, err = processDoc(f)
			if err != nil {
				return nil, err
			}
		}
	}
	docx = &DocxLib{
		Document:    *doc,
		DocRelation: *relations,
	}
	return docx, nil
}

// Processes one of the relevant files, the one with the actual document
func processDoc(file *zip.File) (*Document, error) {
	filebytes, err := readZipFile(file)
	if err != nil {
		fmt.Println("Error reading from internal zip file")
		return nil, err
	}
	doc := Document{
		XMLW:    XMLNS_W,
		XMLR:    XMLNS_R,
		XMLName: xml.Name{Space: XMLNS_W, Local: "document"}}
	err = xml.Unmarshal(filebytes, &doc)
	if err != nil {
		fmt.Println("Error unmarshalling doc")
		fmt.Println(string(filebytes))
		return nil, err
	}
	return &doc, nil
}

// Processes one of the relevant files, the one with the relationships
func processRelations(file *zip.File) (*Relationships, error) {
	filebytes, err := readZipFile(file)
	if err != nil {
		fmt.Println("Error reading from internal zip file")
		return nil, err
	}
	rels := Relationships{Xmlns: "none"}
	err = xml.Unmarshal(filebytes, &rels)
	if err != nil {
		fmt.Println("Error unmarshalling relationships")
		return nil, err
	}
	return &rels, nil
}

// From a zip file structure, we return a byte array
func readZipFile(zf *zip.File) ([]byte, error) {
	f, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return ioutil.ReadAll(f)
}
