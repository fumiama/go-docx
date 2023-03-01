package docx

import (
	"encoding/xml"
	"hash/crc64"
	"io"
	"os"
	"testing"
)

func TestShapeStructure(t *testing.T) {
	w := NewA4()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("test shape")
	para1.AddShape(808355, 238760, "AutoShape", "auto", "straightConnector1", []interface{}{
		&ALine{
			W:         9525,
			SolidFill: &ASolidFill{SrgbClr: &ASrgbClr{Val: "000000"}},
			Round:     &struct{}{},
			HeadEnd:   &AHeadEnd{},
			TailEnd:   &ATailEnd{},
		},
	})

	f, err := os.Create("TestMarshalShapeStructure.xml")
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
	f1, err := os.Create("TestUnmarshalShapeStructure.xml")
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
