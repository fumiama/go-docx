package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/fumiama/docxlib"
)

var fileLocation *string

func init() {
	fileLocation = flag.String("file", "new-file.docx", "file location")
	flag.Parse()
}

func main() {
	//Now let's try to read the file
	readFile, err := os.Open(*fileLocation)
	if err != nil {
		panic(err)
	}
	defer readFile.Close()
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docxlib.Parse(readFile, int64(size))
	if err != nil {
		panic(err)
	}
	for _, para := range doc.Document.Body.Paragraphs {
		fmt.Println("New paragraph")
		for _, child := range para.Children {
			if child.Run != nil {
				if child.Run.Text != nil {
					fmt.Printf("\tWe've found a new run with the text ->%s\n", child.Run.Text.Text)
				}
				if child.Run.Drawing != nil {
					fmt.Printf("\tWe've found a new run with the drawing ->%s\n", child.Run.Drawing.Inline.DistT) // TODO: replace to refid
				}
			}
			if child.Link != nil {
				id := child.Link.ID
				text := child.Link.Run.InstrText
				link, err := doc.Refer(id)
				if err != nil {
					fmt.Printf("\tWe found a link with id %s and text %s without target\n", id, text)
				} else {
					fmt.Printf("\tWe've found a new hyperlink with ref %s and the text %s\n", link, text)
				}

			}
		}
		fmt.Print("End of paragraph\n\n")
	}
	fmt.Println("End of main")
}
