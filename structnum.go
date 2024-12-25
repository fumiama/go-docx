package docx

import (
	"encoding/xml"
	"io"
)

// NumProperties show the number properties
type NumProperties struct {
	XMLName xml.Name `xml:"w:numPr,omitempty"`
	NumID   *NumID
	Ilvl    *Ilevel
}

// NumID show the number id
type NumID struct {
	XMLName xml.Name `xml:"w:numId,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// Ilevel show the level
type Ilevel struct {
	XMLName xml.Name `xml:"w:ilvl,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// UnmarshalXML ...
func (n *NumProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "numId":
				var value NumID
				value.Val = getAtt(tt.Attr, "val")
				n.NumID = &value
			case "ilvl":
				var value Ilevel
				value.Val = getAtt(tt.Attr, "val")
				n.Ilvl = &value
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
