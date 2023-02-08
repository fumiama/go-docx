package docxlib

import (
	"archive/zip"
	"encoding/xml"
	"strings"

	"github.com/golang/glog"
)

// This receives a zip file writer (word documents are a zip with multiple xml inside)
// and writes the relevant files. Some of them come from the empty_constants file,
// others from the actual in-memory structure
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

		_, err = w.Write(StringToBytes(data))
		if err != nil {
			return err
		}
	}

	return
}

func marshal(data interface{}) (out string, err error) {
	sb := strings.Builder{}
	sb.WriteString(xml.Header)
	err = xml.NewEncoder(&sb).Encode(data)
	if err != nil {
		glog.Errorln("Error marshalling", err)
		return
	}
	out = sb.String()
	return
}
