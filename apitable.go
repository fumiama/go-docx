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

// AddTable add a new table to body by col*row
//
// unit: twips (1/20 point)
func (f *Docx) AddTable(row int, col int) *Table {
	trs := make([]*WTableRow, row)
	for i := 0; i < row; i++ {
		cells := make([]*WTableCell, col)
		for i := range cells {
			cells[i] = &WTableCell{
				TableCellProperties: &WTableCellProperties{
					TableCellWidth: &WTableCellWidth{Type: "auto"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			TableRowProperties: &WTableRowProperties{},
			TableCells:         cells,
		}
	}
	tbl := &Table{
		TableProperties: &WTableProperties{
			Width: &WTableWidth{Type: "auto"},
			TableBorders: &WTableBorders{
				Top:     &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Left:    &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Bottom:  &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Right:   &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				InsideH: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				InsideV: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
			},
			Look: &WTableLook{
				Val: "0000",
			},
		},
		TableGrid: &WTableGrid{},
		TableRows: trs,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

// AddTableTwips add a new table to body by height and width
//
// unit: twips (1/20 point)
func (f *Docx) AddTableTwips(rowHeights []int64, colWidths []int64) *Table {
	grids := make([]*WGridCol, len(colWidths))
	trs := make([]*WTableRow, len(rowHeights))
	for i, w := range colWidths {
		if w > 0 {
			grids[i] = &WGridCol{
				W: w,
			}
		}
	}
	for i, h := range rowHeights {
		cells := make([]*WTableCell, len(colWidths))
		for i, w := range colWidths {
			cells[i] = &WTableCell{
				TableCellProperties: &WTableCellProperties{
					TableCellWidth: &WTableCellWidth{W: w, Type: "dxa"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			TableRowProperties: &WTableRowProperties{},
			TableCells:         cells,
		}
		if h > 0 {
			trs[i].TableRowProperties.TableRowHeight = &WTableRowHeight{
				Val: h,
			}
		}
	}
	tbl := &Table{
		TableProperties: &WTableProperties{
			Width: &WTableWidth{Type: "auto"},
			TableBorders: &WTableBorders{
				Top:     &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Left:    &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Bottom:  &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				Right:   &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				InsideH: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
				InsideV: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: "000000"},
			},
			Look: &WTableLook{
				Val: "0000",
			},
		},
		TableGrid: &WTableGrid{
			GridCols: grids,
		},
		TableRows: trs,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

// Justification allows to set table's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (t *Table) Justification(val string) *Table {
	if t.TableProperties.Justification == nil {
		t.TableProperties.Justification = &Justification{Val: val}
		return t
	}
	t.TableProperties.Justification.Val = val
	return t
}

// Justification allows to set table's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (w *WTableRow) Justification(val string) *WTableRow {
	if w.TableRowProperties.Justification == nil {
		w.TableRowProperties.Justification = &Justification{Val: val}
		return w
	}
	w.TableRowProperties.Justification.Val = val
	return w
}

// Shade allows to set cell's shade
func (c *WTableCell) Shade(val, color, fill string) *WTableCell {
	c.TableCellProperties.Shade = &Shade{
		Val:   val,
		Color: color,
		Fill:  fill,
	}
	return c
}
