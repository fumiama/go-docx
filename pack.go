package docxlib

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
)

func (f *Docx) pack(zipWriter *zip.Writer) (err error) {
	files := map[string]string{}

	files["_rels/.rels"] = TEMP_REL
	files["docProps/app.xml"] = TEMP_DOCPROPS_APP
	files["docProps/core.xml"] = TEMP_DOCPROPS_CORE
	files["word/theme/theme1.xml"] = TEMP_WORD_THEME_THEME
	files["word/styles.xml"] = TEMP_WORD_STYLE
	files["[Content_Types].xml"] = TEMP_CONTENT
	files["word/_rels/document.xml.rels"], err = marshal(f.DocRelation)
	if err != nil {
		return err
	}
	files["word/document.xml"], err = marshal(f.Document)
	if err != nil {
		return err
	}

	for path, data := range files {
		w, err := zipWriter.Create(path)
		if err != nil {
			return err
		}

		_, err = w.Write([]byte(data))
		if err != nil {
			return err
		}
	}

	return
}

func marshal(data interface{}) (out string, err error) {
	body, err := xml.Marshal(data)
	if err != nil {
		fmt.Println(err)
		return
	}

	out = xml.Header + string(body)
	return
}
