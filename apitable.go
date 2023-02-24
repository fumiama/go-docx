package docxlib

import "unsafe"

// AddTable add a new table to body by col*row
//
// unit: twips (1/20 point)
func (f *Docx) AddTable(row int, col int) *WTable {
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
	f.Document.Body.mu.Lock()
	defer f.Document.Body.mu.Unlock()
	f.Document.Body.Items = append(f.Document.Body.Items, WTable{
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
	})

	t := f.Document.Body.Items[len(f.Document.Body.Items)-1]

	return *(**WTable)(unsafe.Add(unsafe.Pointer(&t), unsafe.Sizeof(uintptr(0))))
}

// AddTableTwips add a new table to body by height and width
//
// unit: twips (1/20 point)
func (f *Docx) AddTableTwips(rowHeights []int64, colWidths []int64) *WTable {
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
	f.Document.Body.mu.Lock()
	defer f.Document.Body.mu.Unlock()
	f.Document.Body.Items = append(f.Document.Body.Items, WTable{
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
	})

	t := f.Document.Body.Items[len(f.Document.Body.Items)-1]

	return *(**WTable)(unsafe.Add(unsafe.Pointer(&t), unsafe.Sizeof(uintptr(0))))
}
