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
	"strconv"
	"sync/atomic"
)

// AddInlineShape adds wsp named drawing to paragraph
func (p *Paragraph) AddInlineShape(w, h int64, name, bwMode, prst string, ln *ALine) *Run {
	idn := int(atomic.AddUintptr(&p.file.docID, 1))
	id := strconv.Itoa(int(p.file.IncreaseID(name)))
	d := &Drawing{
		Inline: &WPInline{
			Extent: &WPExtent{
				CX: w,
				CY: h,
			},
			EffectExtent: &WPEffectExtent{},
			DocPr: &WPDocPr{
				ID:   idn,
				Name: name + " " + id,
			},
			CNvGraphicFramePr: &WPCNvGraphicFramePr{
				Locks: AGraphicFrameLocks{XMLA: XMLNS_DRAWINGML_MAIN},
			},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_WPS,
					Shape: &WordprocessingShape{
						CNvCnPr: &WPSCNvCnPr{
							ConnShapeLocks: &struct{}{},
						},
						SpPr: &ShapeProperties{
							BWMode: bwMode,

							Xfrm: AXfrm{
								Ext: AExt{
									CX: w,
									CY: h,
								},
							},
							PrstGeom: APrstGeom{
								Prst: prst,
							},
							NoFill: &struct{}{},
							Line:   ln,
						},
						BodyPr: &WPSBodyPr{},
					},
				},
			},
		},
	}
	c := make([]interface{}, 1, 64)
	c[0] = d
	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}
	p.Children = append(p.Children, run)
	return run
}

// AddAnchorShape adds wsp named drawing to paragraph
func (p *Paragraph) AddAnchorShape(w, h int64, name, bwMode, prst string, ln *ALine) *Run {
	idn := int(atomic.AddUintptr(&p.file.docID, 1))
	id := strconv.Itoa(int(p.file.IncreaseID(name)))
	d := &Drawing{
		Anchor: &WPAnchor{
			LayoutInCell: 1,
			AllowOverlap: 1,

			SimplePosXY: &WPSimplePos{},
			PositionH: &WPPositionH{
				RelativeFrom: "column",
			},
			PositionV: &WPPositionV{
				RelativeFrom: "paragraph",
			},

			Extent: &WPExtent{
				CX: w,
				CY: h,
			},
			EffectExtent: &WPEffectExtent{},
			WrapNone:     &struct{}{},
			DocPr: &WPDocPr{
				ID:   idn,
				Name: name + " " + id,
			},
			CNvGraphicFramePr: &WPCNvGraphicFramePr{
				Locks: AGraphicFrameLocks{XMLA: XMLNS_DRAWINGML_MAIN},
			},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_WPS,
					Shape: &WordprocessingShape{
						CNvCnPr: &WPSCNvCnPr{
							ConnShapeLocks: &struct{}{},
						},
						SpPr: &ShapeProperties{
							BWMode: bwMode,

							Xfrm: AXfrm{
								Ext: AExt{
									CX: w,
									CY: h,
								},
							},
							PrstGeom: APrstGeom{
								Prst: prst,
							},
							NoFill: &struct{}{},
							Line:   ln,
						},
						BodyPr: &WPSBodyPr{},
					},
				},
			},
		},
	}
	c := make([]interface{}, 1, 64)
	c[0] = d
	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}
	p.Children = append(p.Children, run)
	return run
}
