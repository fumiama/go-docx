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

// Table represents a table within a Word document.
type Table struct {
	XMLName         xml.Name `xml:"w:tbl,omitempty"`
	TableProperties *WTableProperties
	TableGrid       *WTableGrid
	TableRows       []*WTableRow

	file *Docx
}

func (t *Table) String() string {
	if len(t.TableRows) == 0 || len(t.TableRows[0].TableCells) == 0 {
		return ""
	}
	sb := strings.Builder{}
	sb.WriteString("| ")
	for i := 0; i < len(t.TableRows[0].TableCells); i++ {
		sb.WriteString(" :----: |")
	}
	for _, r := range t.TableRows {
		sb.WriteString("\n|")
		for _, c := range r.TableCells {
			if len(c.Paragraphs) > 0 && len(c.Paragraphs[0].Children) > 0 {
				sb.WriteByte(' ')
				sb.WriteString(c.Paragraphs[0].String())
			} else {
				sb.WriteString("       ")
			}
			sb.WriteString(" |")
		}
	}
	return sb.String()
}

// UnmarshalXML implements the xml.Unmarshaler interface.
func (t *Table) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				value.file = t.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
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
func (t *WTableProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
	LeftFromText  int      `xml:"w:leftFromText,attr,omitempty"`
	RightFromText int      `xml:"w:rightFromText,attr,omitempty"`
	VertAnchor    string   `xml:"w:vertAnchor,attr,omitempty"`
	HorzAnchor    string   `xml:"w:horzAnchor,attr,omitempty"`
	TblpXSpec     string   `xml:"w:tblpXSpec,attr,omitempty"`
	TblpYSpec     string   `xml:"w:tblpYSpec,attr,omitempty"`
	TblpX         int      `xml:"w:tblpX,attr,omitempty"`
	TblpY         int      `xml:"w:tblpY,attr,omitempty"`
}

// UnmarshalXML ...
func (tp *WTablePositioningProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "leftFromText":
			tp.LeftFromText, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		case "rightFromText":
			tp.RightFromText, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		case "vertAnchor":
			tp.VertAnchor = attr.Value
		case "horzAnchor":
			tp.HorzAnchor = attr.Value
		case "tblpXSpec":
			tp.TblpXSpec = attr.Value
		case "tblpYSpec":
			tp.TblpYSpec = attr.Value
		case "tblpX":
			tp.TblpX, err = GetInt(attr.Value)
			if err != nil {
				return err
			}
		case "tblpY":
			tp.TblpY, err = GetInt(attr.Value)
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
			t.W, err = GetInt64(attr.Value)
			if err != nil {
				return err
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
	return err
}

// WTableGrid is a structure that represents the table grid of a Word document.
type WTableGrid struct {
	XMLName  xml.Name    `xml:"w:tblGrid,omitempty"`
	GridCols []*WGridCol `xml:"w:gridCol,omitempty"`
}

// UnmarshalXML ...
func (t *WTableGrid) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			g.W, err = GetInt64(attr.Value)
			if err != nil {
				return err
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
	XMLName xml.Name `xml:"w:tr,omitempty"`
	// RsidR              string   `xml:"w:rsidR,attr,omitempty"`
	// RsidRPr            string   `xml:"w:rsidRPr,attr,omitempty"`
	// RsidTr             string   `xml:"w:rsidTr,attr,omitempty"`
	TableRowProperties *WTableRowProperties
	TableCells         []*WTableCell

	file *Docx
}

// UnmarshalXML ...
func (w *WTableRow) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	/*for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			w.RsidR = attr.Value
		case "rsidRPr":
			w.RsidRPr = attr.Value
		case "rsidTr":
			w.RsidTr = attr.Value
		default:
			// ignore other attributes
		}
	}*/

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
				value.file = w.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
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
	Justification  *Justification
}

// UnmarshalXML ...
func (t *WTableRowProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
					switch attr.Name.Local {
					case "val":
						th.Val, err = GetInt64(attr.Value)
						if err != nil {
							return err
						}
					case "hRule":
						th.Rule = attr.Value
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
	Rule    string   `xml:"w:hRule,attr,omitempty"`
	Val     int64    `xml:"w:val,attr"`
}

// WTableCell represents a cell within a table.
type WTableCell struct {
	XMLName             xml.Name `xml:"w:tc,omitempty"`
	TableCellProperties *WTableCellProperties
	Paragraphs          []*Paragraph `xml:"w:p,omitempty"`
	Tables              []*Table     `xml:"w:tbl,omitempty"`

	file *Docx
}

// UnmarshalXML ...
func (c *WTableCell) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				value.file = c.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Paragraphs = append(c.Paragraphs, &value)
			case "tcPr":
				var value WTableCellProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.TableCellProperties = &value
			case "tbl":
				var table Table
				table.file = c.file
				if err = d.DecodeElement(&table, &tt); err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				c.Tables = append(c.Tables, &table)
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
	VMerge         *WvMerge
	GridSpan       *WGridSpan
	TableBorders   *WTableBorders `xml:"w:tcBorders"`
	Shade          *Shade
	VAlign         *WVerticalAlignment
}

// UnmarshalXML ...
func (r *WTableCellProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				r.TableCellWidth.W, err = GetInt64(v)
				if err != nil {
					return err
				}
				r.TableCellWidth.Type = getAtt(tt.Attr, "type")
			case "vMerge":
				r.VMerge = &WvMerge{Val: getAtt(tt.Attr, "val")}
			case "gridSpan":
				r.GridSpan = new(WGridSpan)
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				r.GridSpan.Val, err = GetInt(v)
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
			case "shd":
				var value Shade
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Shade = &value
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

// WvMerge element is used to specify whether a table cell
// should be vertically merged with the cell(s) above or below it.
// When a cell is merged, its content is merged as well.
//
// The <w:vMerge> element has a single attribute called val which
// specifies the merge behavior. Its possible values are:
//
//	continue: This value indicates that the current cell is part
//	of a vertically merged group of cells, but it is not the first cell
//	in that group. It means that the current cell should not have its
//	own content and should inherit the content of the first cell in the merged group.
//
//	restart: This value indicates that the current cell is the first cell in a
//	new vertically merged group of cells. It means that the current cell should
//	have its own content and should be used as the topmost cell in the merged group.
//
// Note that the <w:vMerge> element is only used in table cells that are part of
// a vertically merged group. For cells that are not part of a merged group,
// this element should be omitted.
type WvMerge struct {
	XMLName xml.Name `xml:"w:vMerge,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
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
func (w *WTableBorders) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
	Val   string `xml:"w:val,attr,omitempty"`
	Size  int    `xml:"w:sz,attr,omitempty"`
	Space int    `xml:"w:space,attr,omitempty"`
	Color string `xml:"w:color,attr,omitempty"`
}

// UnmarshalXML ...
func (t *WTableBorder) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "val":
			t.Val = attr.Value
		case "sz":
			if attr.Value == "" {
				continue
			}
			sz, err := GetInt(attr.Value)
			if err != nil {
				return err
			}
			t.Size = sz
		case "space":
			if attr.Value == "" {
				continue
			}
			space, err := GetInt(attr.Value)
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
	return err
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
