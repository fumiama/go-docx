package docxlib

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io/ioutil"
)

func unpack(zipReader *zip.Reader) (docx *Docx, err error) {
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
	docx = &Docx{
		Document:    *doc,
		DocRelation: *relations,
	}
	return docx, nil
}

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
	//r := bytes.NewReader(filebytes)
	//err = decode(r)
	if err != nil {
		fmt.Println("Error unmarshalling doc")
		fmt.Println(string(filebytes))
		return nil, err
	}
	return &doc, nil
}

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

func readZipFile(zf *zip.File) ([]byte, error) {
	f, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return ioutil.ReadAll(f)
}
