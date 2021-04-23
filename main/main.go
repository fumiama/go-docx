package main

import (
	"fmt"
	"os"

	"github.com/gonfva/docxlib"
)

const FILE_PATH = "/tmp/new-file.docx"

func main() {
	fmt.Printf("Preparing new document to write at %s\n", FILE_PATH)

	w := docxlib.New()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("test")

	para1.AddText("test font size").Size(22)
	para1.AddText("test color").Color("808080")
	para2 := w.AddParagraph()
	para2.AddText("test font size and color").Size(22).Color("ff0000")

	nextPara := w.AddParagraph()
	nextPara.AddLink("google", `http://google.com`)

	f, err := os.Create(FILE_PATH)
	if err != nil {
		panic(err)
	}
	defer f.Close()
	w.Write(f)
	fmt.Println("Document writen. \nNow trying to read it")
	// Now let's try to read the file
	readFile, err := os.Open(FILE_PATH)
	if err != nil {
		panic(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docxlib.Parse(readFile, int64(size))
	if err != nil {
		panic(err)
	}
	for _, para := range doc.Paragraphs() {
		for _, run := range para.Runs() {
			fmt.Printf("\tWe've found a new run with the text ->%s\n", run.Text.Text)
		}
	}
	fmt.Println("End of main")
}
