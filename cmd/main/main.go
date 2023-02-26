/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

// Package main is a function demo
package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/fumiama/go-docx"
)

func main() {
	fileLocation := flag.String("f", "new-file.docx", "file location")
	analyzeOnly := flag.Bool("a", false, "analyze file only")
	unm := flag.Bool("u", false, "lease unmarshalled file")
	flag.Parse()
	var w *docx.Docx
	if !*analyzeOnly {
		fmt.Printf("Preparing new document to write at %s\n", *fileLocation)

		w = docx.NewA4()
		// add new paragraph
		para1 := w.AddParagraph().Justification("distribute")
		r, err := para1.AddAnchorDrawingFrom("testdata/fumiama.JPG")
		if err != nil {
			panic(err)
		}
		r.Children[0].(*docx.Drawing).Anchor.Size(r.Children[0].(*docx.Drawing).Anchor.Extent.CX/4, r.Children[0].(*docx.Drawing).Anchor.Extent.CY/4)
		r.Children[0].(*docx.Drawing).Anchor.BehindDoc = 1
		r.Children[0].(*docx.Drawing).Anchor.PositionH.PosOffset = r.Children[0].(*docx.Drawing).Anchor.Extent.CX
		r.Children[0].(*docx.Drawing).Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &docx.AAlphaModFix{Amount: 50000}
		// add text
		para1.AddText("test").AddTab()
		para1.AddText("size").Size("44").AddTab()
		para1.AddText("color").Color("808080").AddTab()
		para1.AddText("shade").Shade("clear", "auto", "E7E6E6").AddTab()
		para1.AddText("bold").Bold().AddTab()
		para1.AddText("italic").Italic().AddTab()
		para1.AddText("underline").Underline("double").AddTab()
		para1.AddText("highlight").Highlight("yellow").AddTab()
		para1.AddText("font").Font("Consolas", "", "cs").AddTab()

		para2 := w.AddParagraph().Justification("end")
		para2.AddText("test all font attrs").
			Size("44").Color("ff0000").Font("Consolas", "", "cs").
			Shade("clear", "auto", "E7E6E6").
			Bold().Italic().Underline("wave").
			Highlight("yellow")

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
		r.Children[0].(*docx.Drawing).Inline.Size(r.Children[0].(*docx.Drawing).Inline.Extent.CX*4/5, r.Children[0].(*docx.Drawing).Inline.Extent.CY*4/5)
		para4.AddTab().AddTab()
		r, err = para4.AddInlineDrawingFrom("testdata/fumiama2x.webp")
		if err != nil {
			panic(err)
		}
		r.Children[0].(*docx.Drawing).Inline.Size(r.Children[0].(*docx.Drawing).Inline.Extent.CX*4/5, r.Children[0].(*docx.Drawing).Inline.Extent.CY*4/5)

		para5 := w.AddParagraph().Justification("center")
		// add text
		para5.AddText("一行1个 横向 inline").Size("44")

		para6 := w.AddParagraph()
		_, err = para6.AddInlineDrawingFrom("testdata/fumiamayoko.png")
		if err != nil {
			panic(err)
		}

		w.AddParagraph()

		tbl1 := w.AddTable(9, 9)
		for x, r := range tbl1.TableRows {
			red := (x + 1) * 28
			for y, c := range r.TableCells {
				green := ((y + 1) / 3) * 85
				blue := (y%3 + 1) * 85
				v := fmt.Sprintf("%02X%02X%02X", red, green, blue)
				c.Shade("clear", "auto", v).AddParagraph().AddText(v).Size("18")
			}
		}

		w.AddParagraph()

		tbl2 := w.AddTableTwips([]int64{2333, 2333, 2333}, []int64{2333, 2333}).Justification("center")
		for x, r := range tbl2.TableRows {
			r.Justification("center")
			for y, c := range r.TableCells {
				c.TableCellProperties.VAlign = &docx.WVerticalAlignment{Val: "center"}
				c.AddParagraph().Justification("center").AddText(fmt.Sprintf("(%d, %d)", x, y))
			}
		}
		tbl2.TableRows[0].TableCells[0].Shade("clear", "auto", "E7E6E6")

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
	}

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
	doc, err := docx.Parse(readFile, size)
	if err != nil {
		panic(err)
	}
	if *unm {
		i := strings.LastIndex(*fileLocation, "/")
		name := (*fileLocation)[:i+1] + "unmarshal_" + (*fileLocation)[i+1:]
		f, err := os.Create(name)
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
	}
	fmt.Println("Plain text:")
	for _, it := range doc.Document.Body.Items {
		switch para := it.(type) {
		case docx.Paragraph:
			fmt.Println(para.String())
		case docx.WTable:
			fmt.Println("------------------------------")
			for x, r := range para.TableRows {
				fmt.Printf("[%d] ", x)
				for y, c := range r.TableCells {
					if len(c.Paragraphs) > 0 && len(c.Paragraphs[0].Children) > 0 {
						fmt.Printf("<%d> %v\t", y, &c.Paragraphs[0])
					} else {
						fmt.Print("\t")
					}
				}
				fmt.Print("\n")
			}
			fmt.Println("------------------------------")
		}
	}
	fmt.Println("End of main")
}
