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
// You can it invoke from a file
//		readFile, err := os.Open(FILE_PATH)
//		if err != nil {
//			panic(err)
//		}
//		fileinfo, err := readFile.Stat()
//		if err != nil {
//			panic(err)
//		}
//		size := fileinfo.Size()
//		doc, err := docxlib.Parse(readFile, int64(size))
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
