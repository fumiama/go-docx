package docxlib

import (
	"encoding/xml"
	"hash/crc64"
	"io"
	"os"
	"testing"
)

func TestDrawingStructure(t *testing.T) {
	w := NewA4()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("直接粘贴 inline").AddTab()
	r, err := para1.AddAnchorDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		t.Fatal(err)
	}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &AAlphaModFix{Amount: 50000}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr.Locks = &APicLocks{NoChangeAspect: 1}
	r.Drawing.Anchor.Graphic.GraphicData.Pic.SpPr.Xfrm.Rot = 50000
	para2 := w.AddParagraph().Justification("center")
	para2.AddInlineDrawingFrom("testdata/fumiama.JPG")
	para2.AddTab().AddTab().AppendTab().AppendTab()
	para2.AddInlineDrawingFrom("testdata/fumiama2x.webp")

	para3 := w.AddParagraph()
	para3.AddInlineDrawingFrom("testdata/fumiamayoko.png")

	f, err := os.Create("TestMarshalInlineDrawingStructure.xml")
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
	f1, err := os.Create("TestUnmarshalInlineDrawingStructure.xml")
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
