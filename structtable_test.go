package docxlib

import (
	"encoding/xml"
	"hash/crc64"
	"io"
	"os"
	"testing"
)

func TestTableStructure(t *testing.T) {
	w := NewA4()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("table")
	tab1 := w.AddTable(4, 3)
	para2 := tab1.TableRows[3].TableCells[2].AddParagraph()
	r, err := para2.AddAnchorDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		t.Fatal(err)
	}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &AAlphaModFix{Amount: 50000}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr.Locks = &APicLocks{NoChangeAspect: 1}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.SpPr.Xfrm.Rot = 50000
	para3 := tab1.TableRows[0].TableCells[0].AddParagraph()
	para3.AddText("first cell")

	f, err := os.Create("TestMarshalTableStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	_, err = marshaller{data: &w.Document}.WriteTo(f)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	w = NewA4()
	err = xml.NewDecoder(f).Decode(&w.Document)
	if err != nil {
		t.Fatal(err)
	}
	f1, err := os.Create("TestUnmarshalTableStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f1.Close()
	_, err = marshaller{data: &w.Document}.WriteTo(f1)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f1.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	h := crc64.New(crc64.MakeTable(crc64.ECMA))
	_, err = io.Copy(h, f)
	if err != nil {
		t.Fatal(err)
	}
	crc1 := h.Sum64()
	h.Reset()
	_, err = io.Copy(h, f1)
	if err != nil {
		t.Fatal(err)
	}
	crc2 := h.Sum64()
	if crc1 != crc2 {
		t.Fail()
	}
}
