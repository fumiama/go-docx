package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

type WordprocessingCanvas struct {
	XMLName    xml.Name `xml:"wpc:wpc,omitempty"`
	Background *WPCBackground
	Whole      *WPCWhole

	Items []interface{}
}

// UnmarshalXML ...
func (c *WordprocessingCanvas) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "bg":
				c.Background = new(WPCBackground)
				err = d.DecodeElement(c.Background, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "whole":
				c.Whole = new(WPCWhole)
				err = d.DecodeElement(c.Whole, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "wsp":
				var value WPSWordprocessingShape
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Items = append(c.Items, &value)
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

type WPCBackground struct {
	XMLName xml.Name  `xml:"wpc:bg,omitempty"`
	NoFill  *struct{} `xml:"a:noFill,omitempty"`
}

// UnmarshalXML ...
func (b *WPCBackground) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "noFill":
				b.NoFill = &struct{}{}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

type WPCWhole struct {
	XMLName xml.Name `xml:"wpc:whole,omitempty"`
	Line    *ALine
}

// UnmarshalXML ...
func (w *WPCWhole) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "ln":
				w.Line = new(ALine)
				err = d.DecodeElement(w.Line, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}
