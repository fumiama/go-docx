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
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/fumiama/go-docx"
)

func main() {
	fileLocation := flag.String("f", "new-file.docx", "file location")
	analyzeOnly := flag.Bool("a", false, "analyze file only")
	clean := flag.Bool("c", false, "clean mode (keep text and picture only)")
	unm := flag.Bool("u", false, "lease unmarshalled file")
	splitre := flag.String("s", "", "split file into many docxs by matching regex")
	droppp := flag.Bool("p", false, "drop all paragraph properties")
	dupnum := flag.Uint("d", 0, "copy times of the file into dup_filename")
	flag.Parse()
	var w *docx.Docx
	if !*analyzeOnly {
		fmt.Printf("Preparing new document to write at %s\n", *fileLocation)

		w = docx.New().WithDefaultTheme().WithA4Page()
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
		para1.AddText("font").Font("Consolas", "", "", "cs").AddTab()

		para2 := w.AddParagraph().Justification("end")
		para2.AddText("test all font attrs").
			Size("44").Color("ff0000").
			Font("Consolas", "", "", "cs").
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

		w.AddParagraph().AddPageBreaks()
		para5 := w.AddParagraph().Justification("center")
		// add text
		para5.AddText("一行1个 横向 inline").Size("44")

		para6 := w.AddParagraph()
		_, err = para6.AddInlineDrawingFrom("testdata/fumiamayoko.png")
		if err != nil {
			panic(err)
		}

		w.AddParagraph()

		tbl1 := w.AddTable(9, 9, 0, nil)
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

		tbl2 := w.AddTableTwips([]int64{2333, 2333, 2333}, []int64{2333, 2333}, 0, nil).Justification("center")
		for x, r := range tbl2.TableRows {
			r.Justification("center")
			for y, c := range r.TableCells {
				c.TableCellProperties.VAlign = &docx.WVerticalAlignment{Val: "center"}
				c.AddParagraph().Justification("center").AddText(fmt.Sprintf("(%d, %d)", x, y))
			}
		}
		tbl2.TableRows[0].TableCells[0].Shade("clear", "auto", "E7E6E6")

		p := w.AddParagraph().Justification("center")
		p.AddText("测试 AutoShape w:ln").Size("44")
		_ = p.AddAnchorShape(808355, 238760, "AutoShape", "auto", "straightConnector1",
			&docx.ALine{
				W:         9525,
				SolidFill: &docx.ASolidFill{SrgbClr: &docx.ASrgbClr{Val: "000000"}},
				Round:     &struct{}{},
				HeadEnd:   &docx.AHeadEnd{},
				TailEnd:   &docx.ATailEnd{},
			},
		)
		_ = p.AddInlineShape(808355, 238760, "AutoShape", "auto", "straightConnector1",
			&docx.ALine{
				W:         9525,
				SolidFill: &docx.ASolidFill{SrgbClr: &docx.ASrgbClr{Val: "000000"}},
				Round:     &struct{}{},
				HeadEnd:   &docx.AHeadEnd{},
				TailEnd:   &docx.ATailEnd{},
			},
		)

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
		fmt.Println("Document written. \nNow trying to read it")
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
	if *clean {
		doc.Document.Body.DropDrawingOf("NilPicture")
	}
	if *droppp {
		for _, it := range doc.Document.Body.Items {
			switch o := it.(type) {
			case *docx.Paragraph: // printable
				o.Properties = nil
			case *docx.Table: // printable
				for _, tr := range o.TableRows {
					for _, tc := range tr.TableCells {
						for _, p := range tc.Paragraphs {
							p.Properties = nil
						}
					}
				}
			}
		}
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
		switch o := it.(type) {
		case *docx.Paragraph: // printable
			fmt.Println(o.String())
		case *docx.Table: // printable
			fmt.Println(o.String())
		}
	}
	if *splitre != "" {
		a := strings.LastIndex(*fileLocation, "/")
		b := strings.LastIndex(*fileLocation, ".")
		for i, splitteddoc := range doc.SplitByParagraph(docx.SplitDocxByPlainTextRegex(regexp.MustCompile(*splitre))) {
			name := (*fileLocation)[:a+1] + "unmarshal_" + (*fileLocation)[a+1:b] + "_split" + strconv.Itoa(i) + (*fileLocation)[b:]
			f, err := os.Create(name)
			if err != nil {
				panic(err)
			}
			_, err = splitteddoc.WriteTo(f)
			if err != nil {
				panic(err)
			}
			err = f.Close()
			if err != nil {
				panic(err)
			}
		}
	}
	if *dupnum > 1 {
		a := strings.LastIndex(*fileLocation, "/")
		name := "dup_" + (*fileLocation)
		if a > 0 {
			name = (*fileLocation)[:a+1] + "dup_" + (*fileLocation)[a:]
		}
		f, err := os.Create(name)
		if err != nil {
			panic(err)
		}
		newFile := docx.New().WithDefaultTheme().WithA4Page()
		for i := 0; i < int(*dupnum); i++ {
			newFile.AppendFile(doc)
		}
		_, err = io.Copy(f, newFile)
		if err != nil {
			panic(err)
		}
	}
	fmt.Println("End of main")
}
