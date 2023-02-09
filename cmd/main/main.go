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
	fmt.Printf("Preparing new document to write at %s\n", *fileLocation)

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

	f, err := os.Create(*fileLocation)
	if err != nil {
		panic(err)
	}
	err = w.Write(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
	fmt.Println("Document writen. \nNow trying to read it")
	// Now let's try to read the file
	readFile, err := os.Open(*fileLocation)
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
	for _, para := range doc.Document.Body.Paragraphs {
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
	}
	fmt.Println("End of main")
}
