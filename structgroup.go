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

package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

// WordprocessingGroup represents a group of drawing objects or pictures
type WordprocessingGroup struct {
	XMLName              xml.Name `xml:"wpg:wgp,omitempty"`
	CNvGrpSpPr           *WPGcNvGrpSpPr
	GroupShapeProperties *ShapeProperties `xml:"wpg:grpSpPr,omitempty"`
	Elems                []interface{}

	file *Docx
}

// UnmarshalXML ...
func (w *WordprocessingGroup) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "cNvGrpSpPr":
				var value WPGcNvGrpSpPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.CNvGrpSpPr = &value
			case "grpSpPr":
				var value ShapeProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.GroupShapeProperties = &value
			case "pic":
				var value Picture
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
			case "wsp":
				var value WordprocessingShape
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
			case "wpc":
				var value WordprocessingCanvas
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
			case "grpSp":
				var value WPGGroupShape
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
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

// WPGcNvGrpSpPr represents the non-visual properties of a group shape.
type WPGcNvGrpSpPr struct {
	XMLName xml.Name `xml:"wpg:cNvGrpSpPr,omitempty"`
	Locks   *AGroupShapeLocks
}

// UnmarshalXML ...
func (w *WPGcNvGrpSpPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "grpSpLocks":
				w.Locks = new(AGroupShapeLocks)
				err = d.Skip() // skip innerxml
				if err != nil {
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

// AGroupShapeLocks represents the locks applied to a group shape.
type AGroupShapeLocks struct {
	XMLName xml.Name `xml:"a:grpSpLocks,omitempty"`
}

// WPGGroupShape ...
type WPGGroupShape struct {
	XMLName              xml.Name             `xml:"wpg:grpSp,omitempty"`
	CNvPr                *NonVisualProperties `xml:"wpg:cNvPr,omitempty"`
	CNvGrpSpPr           *WPGcNvGrpSpPr
	GroupShapeProperties *ShapeProperties `xml:"wpg:grpSpPr,omitempty"`
	Elems                []interface{}

	file *Docx
}

// UnmarshalXML ...
func (w *WPGGroupShape) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "cNvPr":
				var value NonVisualProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.CNvPr = &value
			case "cNvGrpSpPr":
				var value WPGcNvGrpSpPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.CNvGrpSpPr = &value
			case "grpSpPr":
				var value ShapeProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.GroupShapeProperties = &value
			case "pic":
				var value Picture
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
			case "wsp":
				var value WordprocessingShape
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
			case "wpc":
				var value WordprocessingCanvas
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Elems = append(w.Elems, &value)
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
