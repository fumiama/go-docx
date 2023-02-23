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

	w := docxlib.NewA4()
	// add new paragraph
	para1 := w.AddParagraph().Justification("distribute")
	r, err := para1.AddAnchorDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		panic(err)
	}
	r.Drawing.Anchor.Size(r.Drawing.Anchor.Extent.CX/4, r.Drawing.Anchor.Extent.CY/4)
	r.Drawing.Anchor.BehindDoc = 1
	r.Drawing.Anchor.PositionH.PosOffset = r.Drawing.Anchor.Extent.CX
	r.Drawing.Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &docxlib.AAlphaModFix{Amount: 50000}
	// add text
	para1.AddText("test")
	para1.AddText("test font size").Size("44")
	para1.AddText("test color").Color("808080")

	para2 := w.AddParagraph().Justification("end")
	para2.AddText("test font size and color").Size("44").Color("ff0000")

	nextPara := w.AddParagraph()
	nextPara.AddLink("google", `http://google.com`)

	para3 := w.AddParagraph().Justification("center")
	// add text
	para3.AddText("一行2个 inline").Size("44")

	para4 := w.AddParagraph().Justification("center")
	r, err = para4.AddInlineDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		panic(err)
	}
	r.Drawing.Inline.Size(r.Drawing.Inline.Extent.CX*4/5, r.Drawing.Inline.Extent.CY*4/5)
	para4.AddTab().AddTab()
	r, err = para4.AddInlineDrawingFrom("testdata/fumiama2x.webp")
	if err != nil {
		panic(err)
	}
	r.Drawing.Inline.Size(r.Drawing.Inline.Extent.CX*4/5, r.Drawing.Inline.Extent.CY*4/5)

	para5 := w.AddParagraph().Justification("center")
	// add text
	para5.AddText("一行1个 横向 inline").Size("44")

	para6 := w.AddParagraph()
	_, err = para6.AddInlineDrawingFrom("testdata/fumiamayoko.png")
	if err != nil {
		panic(err)
	}

	f, err := os.Create(*fileLocation)
	if err != nil {
		panic(err)
	}
	_, err = w.WriteTo(f)
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
	doc, err := docxlib.Parse(readFile, size)
	if err != nil {
		panic(err)
	}
	for _, para := range doc.Document.Body.Paragraphs {
		fmt.Println("New paragraph")
		for _, child := range para.Children {
			switch o := child.(type) {
			case *docxlib.Run:
				if o.Text != nil {
					fmt.Printf("\tWe've found a new run with the text ->%s\n", o.Text.Text)
				}
				if o.Drawing != nil {
					if o.Drawing.Inline != nil {
						fmt.Printf("\tWe've found a new run with the inline drawing ->%s\n", o.Drawing.Inline.DocPr.Name)
					}
					if o.Drawing.Anchor != nil {
						fmt.Printf("\tWe've found a new run with the anchor drawing ->%s\n", o.Drawing.Anchor.DocPr.Name)
					}
				}
			case *docxlib.Hyperlink:
				id := o.ID
				text := o.Run.InstrText
				link, err := doc.ReferTarget(id)
				if err != nil {
					fmt.Printf("\tWe found a link with id %s and text %s without target\n", id, text)
				} else {
					fmt.Printf("\tWe've found a new hyperlink with ref %s and the text %s\n", link, text)
				}
			}
		}
		fmt.Print("End of paragraph\n\n")
	}
	f, err = os.Create("unmarshal_" + *fileLocation)
	if err != nil {
		panic(err)
	}
	_, err = doc.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
	fmt.Println("End of main")
}
