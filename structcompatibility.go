//go:build ignore

package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

// AlternateContent ...
type AlternateContent struct {
	XMLName xml.Name `xml:"mc:AlternateContent,omitempty"`

	Choice   *MCChoice
	Fallback struct{} `xml:"mc:Fallback"`

	file *Docx
}

// UnmarshalXML ...
func (a *AlternateContent) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "Choice":
				var value MCChoice
				value.file = a.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				a.Choice = &value
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

// MCChoice ...
type MCChoice struct {
	XMLName  xml.Name `xml:"mc:Choice,omitempty"`
	Requires string   `xml:",attr,omitempty"`

	Elems []interface{}

	file *Docx
}

// UnmarshalXML ...
func (c *MCChoice) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "Requires":
			c.Requires = attr.Value
		default:
			// ignore other attributes
		}
	}
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
			case "drawing":
				var value Drawing
				value.file = c.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Elems = append(c.Elems, &value)
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
