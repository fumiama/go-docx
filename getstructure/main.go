package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/golang/glog"
	"github.com/gonfva/docxlib"
)

const FILE_PATH = "/tmp/new-file.docx"

func main() {
	flag.Parse()
	//Now let's try to read the file
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
		glog.Infoln("There is a new paragraph", para)
		for _, child := range para.Children() {
			if child.Run != nil {
				fmt.Printf("\tWe've found a new run with the text ->%s\n", child.Run.Text.Text)
			}
			if child.Link != nil {
				id := child.Link.ID
				text := child.Link.Run.InstrText
				link, err := doc.References(id)
				if err != nil {
					fmt.Printf("\tWe found a link with id %s and text %s without target\n", id, text)
				} else {
					fmt.Printf("\tWe've found a new hyperlink with ref %s and the text %s\n", link, text)
				}

			}
		}
	}
	fmt.Println("End of main")
}
