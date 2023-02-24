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

package docxlib

import (
	"encoding/xml"
	"io"
	"strconv"
	"strings"
	"sync"
)

// WTable represents a table within a Word document.
type WTable struct {
	XMLName         xml.Name `xml:"w:tbl,omitempty"`
	TableProperties *WTableProperties
	TableGrid       *WTableGrid
	TableRows       []*WTableRow

	file *Docx
}

// UnmarshalXML implements the xml.Unmarshaler interface.
func (t *WTable) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := token.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "tr":
				var value WTableRow
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.file = t.file
				t.TableRows = append(t.TableRows, &value)
			case "tblPr":
				t.TableProperties = new(WTableProperties)
				err = d.DecodeElement(t.TableProperties, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tblGrid":
				t.TableGrid = new(WTableGrid)
				err = d.DecodeElement(t.TableGrid, &tt)
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

// WTableProperties is an element that represents the properties of a table in Word document.
type WTableProperties struct {
	XMLName       xml.Name `xml:"w:tblPr,omitempty"`
	Position      *WTablePositioningProperties
	Style         *WTableStyle
	Width         *WTableWidth
	Justification *Justification `xml:"w:jc,omitempty"`
	TableBorders  *WTableBorders `xml:"w:tblBorders"`
	Look          *WTableLook
}

// UnmarshalXML implements the xml.Unmarshaler interface.
func (t *WTableProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := token.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "tblpPr":
				t.Position = new(WTablePositioningProperties)
				err = d.DecodeElement(t.Position, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tblStyle":
				t.Style = new(WTableStyle)
				err = d.DecodeElement(t.Style, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tblW":
				t.Width = new(WTableWidth)
				err = d.DecodeElement(t.Width, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "jc":
				th := new(Justification)
				for _, attr := range tt.Attr {
					if attr.Name.Local == "val" {
						th.Val = attr.Value
						break
					}
				}
				t.Justification = th
				err = d.Skip()
				if err != nil {
					return err
				}
			case "tblLook":
				t.Look = new(WTableLook)
				err = d.DecodeElement(t.Look, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tblBorders":
				t.TableBorders = new(WTableBorders)
				err = d.DecodeElement(t.TableBorders, &tt)
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

// WTablePositioningProperties is an element that contains the properties
// for positioning a table within a document page, including its horizontal
// and vertical anchors, distance from text, and coordinates.
type WTablePositioningProperties struct {
	XMLName       xml.Name `xml:"w:tblpPr,omitempty"`
	LeftFromText  int      `xml:"w:leftFromText,attr"`
	RightFromText int      `xml:"w:rightFromText,attr"`
	VertAnchor    string   `xml:"w:vertAnchor,attr"`
	HorzAnchor    string   `xml:"w:horzAnchor,attr"`
	TblpX         int      `xml:"w:tblpX,attr"`
	TblpY         int      `xml:"w:tblpY,attr"`
}

// UnmarshalXML ...
func (tp *WTablePositioningProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "leftFromText":
			tp.LeftFromText, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "rightFromText":
			tp.RightFromText, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "vertAnchor":
			tp.VertAnchor = attr.Value
		case "horzAnchor":
			tp.HorzAnchor = attr.Value
		case "tblpX":
			tp.TblpX, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "tblpY":
			tp.TblpY, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		}
	}

	// Consume the end element
	_, err = d.Token()
	return err
}

// WTableStyle represents the style of a table in a Word document.
type WTableStyle struct {
	XMLName xml.Name `xml:"w:tblStyle,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

// UnmarshalXML ...
func (t *WTableStyle) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			t.Val = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// WTableWidth represents the width of a table in a Word document.
type WTableWidth struct {
	XMLName xml.Name `xml:"w:tblW,omitempty"`
	W       int64    `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// UnmarshalXML ...
func (t *WTableWidth) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "w":
			t.W, err = strconv.ParseInt(attr.Value, 10, 64)
			if err != nil {
				return
			}
		case "type":
			t.Type = attr.Value
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// WTableLook represents the look of a table in a Word document.
type WTableLook struct {
	XMLName  xml.Name `xml:"w:tblLook,omitempty"`
	Val      string   `xml:"w:val,attr"`
	FirstRow int      `xml:"w:firstRow,attr"`
	LastRow  int      `xml:"w:lastRow,attr"`
	FirstCol int      `xml:"w:firstColumn,attr"`
	LastCol  int      `xml:"w:lastColumn,attr"`
	NoHBand  int      `xml:"w:noHBand,attr"`
	NoVBand  int      `xml:"w:noVBand,attr"`
}

// UnmarshalXML ...
func (t *WTableLook) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			t.Val = attr.Value
		case "firstRow":
			t.FirstRow = int(attr.Value[0] - '0')
		case "lastRow":
			t.LastRow = int(attr.Value[0] - '0')
		case "firstColumn":
			t.FirstCol = int(attr.Value[0] - '0')
		case "lastColumn":
			t.LastCol = int(attr.Value[0] - '0')
		case "noHBand":
			t.NoHBand = int(attr.Value[0] - '0')
		case "noVBand":
			t.NoVBand = int(attr.Value[0] - '0')
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err := d.Token()
	if err != nil {
		return err
	}
	return nil
}

// WTableGrid is a structure that represents the table grid of a Word document.
type WTableGrid struct {
	XMLName  xml.Name    `xml:"w:tblGrid,omitempty"`
	GridCols []*WGridCol `xml:"w:gridCol,omitempty"`
}

// UnmarshalXML ...
func (t *WTableGrid) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if el, ok := tok.(xml.StartElement); ok {
			switch el.Name.Local {
			case "gridCol":
				var gc WGridCol
				err := d.DecodeElement(&gc, &el)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				t.GridCols = append(t.GridCols, &gc)
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

// WGridCol is a structure that represents a table grid column of a Word document.
type WGridCol struct {
	XMLName xml.Name `xml:"w:gridCol,omitempty"`
	W       int64    `xml:"w:w,attr"`
}

// UnmarshalXML ...
func (g *WGridCol) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "w":
			g.W, err = strconv.ParseInt(attr.Value, 10, 64)
			if err != nil {
				return
			}
		default:
			// ignore other attributes
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// WTableRow represents a row within a table.
type WTableRow struct {
	XMLName            xml.Name `xml:"w:tr,omitempty"`
	RsidR              string   `xml:"w:rsidR,attr"`
	RsidTr             string   `xml:"w:rsidTr,attr"`
	TableRowProperties *WTableRowProperties
	TableCells         []*WTableCell

	file *Docx
}

// UnmarshalXML ...
func (w *WTableRow) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			w.RsidR = attr.Value
		case "rsidTr":
			w.RsidTr = attr.Value
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
			case "trPr":
				w.TableRowProperties = new(WTableRowProperties)
				err = d.DecodeElement(w.TableRowProperties, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			case "tc":
				var value WTableCell
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				value.file = w.file
				w.TableCells = append(w.TableCells, &value)
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

// WTableRowProperties represents the properties of a row within a table.
type WTableRowProperties struct {
	XMLName        xml.Name `xml:"w:trPr,omitempty"`
	TableRowHeight *WTableRowHeight
	Justification  *Justification `xml:"w:jc,omitempty"`
}

// UnmarshalXML ...
func (t *WTableRowProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := tok.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "trHeight":
				th := new(WTableRowHeight)
				for _, attr := range tt.Attr {
					if attr.Name.Local == "val" {
						th.Val, err = strconv.ParseInt(attr.Value, 10, 64)
						if err != nil {
							return err
						}
						break
					}
				}
				t.TableRowHeight = th
				err = d.Skip()
				if err != nil {
					return err
				}
			case "jc":
				th := new(Justification)
				for _, attr := range tt.Attr {
					if attr.Name.Local == "val" {
						th.Val = attr.Value
						break
					}
				}
				t.Justification = th
				err = d.Skip()
				if err != nil {
					return err
				}
			default:
				err = d.Skip()
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// WTableRowHeight represents the height of a row within a table.
type WTableRowHeight struct {
	XMLName xml.Name `xml:"w:trHeight,omitempty"`
	Val     int64    `xml:"w:val,attr"`
}

// WTableCell represents a cell within a table.
type WTableCell struct {
	mu sync.Mutex

	XMLName             xml.Name `xml:"w:tc,omitempty"`
	TableCellProperties *WTableCellProperties
	Paragraphs          []Paragraph `xml:"w:p,omitempty"`

	file *Docx
}

// UnmarshalXML ...
func (c *WTableCell) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "p":
				var value Paragraph
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				if len(value.Children) > 0 {
					value.file = c.file
					c.mu.Lock()
					c.Paragraphs = append(c.Paragraphs, value)
					c.mu.Unlock()
				}
			case "tcPr":
				var value WTableCellProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.TableCellProperties = &value
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

// WTableCellProperties represents the properties of a table cell.
type WTableCellProperties struct {
	XMLName        xml.Name `xml:"w:tcPr,omitempty"`
	TableCellWidth *WTableCellWidth
	GridSpan       *WGridSpan
	VAlign         *WVerticalAlignment
	TableBorders   *WTableBorders `xml:"w:tcBorders"`
}

// UnmarshalXML ...
func (r *WTableCellProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "tcW":
				r.TableCellWidth = new(WTableCellWidth)
				v := getAtt(tt.Attr, "w")
				if v == "" {
					continue
				}
				r.TableCellWidth.W, err = strconv.ParseInt(v, 10, 64)
				if err != nil {
					return err
				}
				r.TableCellWidth.Type = getAtt(tt.Attr, "type")
			case "gridSpan":
				r.GridSpan = new(WGridSpan)
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				r.GridSpan.Val, err = strconv.Atoi(v)
				if err != nil {
					return err
				}
			case "vAlign":
				r.VAlign = new(WVerticalAlignment)
				r.VAlign.Val = getAtt(tt.Attr, "val")
			case "tcBorders":
				r.TableBorders = new(WTableBorders)
				err = d.DecodeElement(r.TableBorders, &tt)
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

// WTableCellWidth represents the width of a table cell.
//
// 在w:tcW元素中，type属性可以有以下几种取值：
//
//	"auto"：表示表格列宽度由文本或表格布局决定。
//	"dxa"：表示表格列宽度使用磅为单位。
//
// 不同的取值对应着不同的宽度计量单位和宽度定义方式。
type WTableCellWidth struct {
	XMLName xml.Name `xml:"w:tcW,omitempty"`
	W       int64    `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// WTableBorders is a structure representing the borders of a Word table.
type WTableBorders struct {
	Top     *WTableBorder `xml:"w:top,omitempty"`
	Left    *WTableBorder `xml:"w:left,omitempty"`
	Bottom  *WTableBorder `xml:"w:bottom,omitempty"`
	Right   *WTableBorder `xml:"w:right,omitempty"`
	InsideH *WTableBorder `xml:"w:insideH,omitempty"`
	InsideV *WTableBorder `xml:"w:insideV,omitempty"`
}

// UnmarshalXML ...
func (w *WTableBorders) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "top":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Top = value
			case "left":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Left = value
			case "bottom":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Bottom = value
			case "right":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.Right = value
			case "insideH":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.InsideH = value
			case "insideV":
				value := new(WTableBorder)
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				w.InsideV = value
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

// WTableBorder is a structure representing a single border of a Word table.
type WTableBorder struct {
	Val   string `xml:"w:val,attr"`
	Size  int    `xml:"w:sz,attr"`
	Space int    `xml:"w:space,attr"`
	Color string `xml:"w:color,attr"`
}

// UnmarshalXML ...
func (t *WTableBorder) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			t.Val = attr.Value
		case "sz":
			sz, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			t.Size = sz
		case "space":
			space, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			t.Space = space
		case "color":
			t.Color = attr.Value
		}
	}
	// Consume the end element
	_, err := d.Token()
	if err != nil {
		return err
	}
	return nil
}

// WGridSpan represents the number of grid columns this cell should span.
type WGridSpan struct {
	XMLName xml.Name `xml:"w:gridSpan,omitempty"`
	Val     int      `xml:"w:val,attr"`
}

// WVerticalAlignment represents the vertical alignment of the content of a cell.
type WVerticalAlignment struct {
	XMLName xml.Name `xml:"w:vAlign,omitempty"`
	Val     string   `xml:"w:val,attr"`
}
