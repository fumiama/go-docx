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
	"bytes"
	"os"
	"strconv"
	"sync/atomic"

	"github.com/fumiama/imgsz"
)

// AddInlineDrawing adds inline drawing to paragraph
func (p *Paragraph) AddInlineDrawing(pic []byte) (*Run, error) {
	sz, format, err := imgsz.DecodeSize(bytes.NewReader(pic))
	if err != nil {
		return nil, err
	}
	idn := int(atomic.AddUintptr(&p.file.docID, 1))
	id := int(p.file.IncreaseID("图片"))
	ids := strconv.Itoa(id)
	rid := p.file.addImage(format, pic)
	w, h := int64(sz.Width), int64(sz.Height)
	if float64(w)/float64(h) > 1.2 {
		h = A4_EMU_MAX_WIDTH * h / w
		w = A4_EMU_MAX_WIDTH
	} else {
		h = A4_EMU_MAX_WIDTH * h / w / 2
		w = A4_EMU_MAX_WIDTH / 2
	}
	d := &Drawing{
		Inline: &WPInline{
			// AnchorID: fmt.Sprintf("%08X", rand.Uint32()),
			// EditID:   fmt.Sprintf("%08X", rand.Uint32()),

			Extent: &WPExtent{
				CX: w,
				CY: h,
			},
			EffectExtent: &WPEffectExtent{},
			DocPr: &WPDocPr{
				ID:   idn,
				Name: "图片 " + ids,
			},
			CNvGraphicFramePr: &WPCNvGraphicFramePr{
				Locks: AGraphicFrameLocks{
					XMLA:           XMLNS_DRAWINGML_MAIN,
					NoChangeAspect: 1,
				},
			},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_PICTURE,
					Pic: &Picture{
						XMLPIC: XMLNS_DRAWINGML_PICTURE,
						NonVisualPicProperties: &PICNonVisualPicProperties{
							NonVisualDrawingProperties: NonVisualProperties{
								ID:   id,
								Name: "图片 " + ids,
							},
						},
						BlipFill: &PICBlipFill{
							Blip: ABlip{
								Embed:  rid,
								Cstate: "print",
							},
						},
						SpPr: &PICSpPr{
							Xfrm: AXfrm{
								Ext: AExt{
									CX: w,
									CY: h,
								},
							},
							PrstGeom: &APrstGeom{
								Prst: "rect",
							},
						},
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
	return run, nil
}

// AddInlineDrawingFrom adds drawing from file to paragraph
func (p *Paragraph) AddInlineDrawingFrom(file string) (*Run, error) {
	data, err := os.ReadFile(file)
	if err != nil {
		return nil, err
	}
	return p.AddInlineDrawing(data)
}

// Size of the inline drawing by EMU
func (r *WPInline) Size(w, h int64) {
	if r.Extent != nil {
		r.Extent.CX = w
		r.Extent.CY = h
	}
	if r.Graphic != nil && r.Graphic.GraphicData != nil && r.Graphic.GraphicData.Pic != nil && r.Graphic.GraphicData.Pic.SpPr != nil {
		r.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = w
		r.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CY = h
	}
}

// AddAnchorDrawing adds inline drawing to paragraph
func (p *Paragraph) AddAnchorDrawing(pic []byte) (*Run, error) {
	sz, format, err := imgsz.DecodeSize(bytes.NewReader(pic))
	if err != nil {
		return nil, err
	}
	idn := int(atomic.AddUintptr(&p.file.docID, 1))
	id := int(p.file.IncreaseID("图片"))
	ids := strconv.Itoa(id)
	rid := p.file.addImage(format, pic)
	w, h := int64(sz.Width), int64(sz.Height)
	if float64(w)/float64(h) > 1.2 {
		h = A4_EMU_MAX_WIDTH * h / w
		w = A4_EMU_MAX_WIDTH
	} else {
		h = A4_EMU_MAX_WIDTH * h / w / 2
		w = A4_EMU_MAX_WIDTH / 2
	}
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
				Name: "图片 " + ids,
			},
			CNvGraphicFramePr: &WPCNvGraphicFramePr{
				Locks: AGraphicFrameLocks{
					XMLA:           XMLNS_DRAWINGML_MAIN,
					NoChangeAspect: 1,
				},
			},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_PICTURE,
					Pic: &Picture{
						XMLPIC: XMLNS_DRAWINGML_PICTURE,
						NonVisualPicProperties: &PICNonVisualPicProperties{
							NonVisualDrawingProperties: NonVisualProperties{
								ID:   id,
								Name: "图片 " + ids,
							},
						},
						BlipFill: &PICBlipFill{
							Blip: ABlip{
								Embed:  rid,
								Cstate: "print",
							},
						},
						SpPr: &PICSpPr{
							Xfrm: AXfrm{
								Ext: AExt{
									CX: w,
									CY: h,
								},
							},
							PrstGeom: &APrstGeom{
								Prst: "rect",
							},
						},
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
	return run, nil
}

// AddAnchorDrawingFrom adds drawing from file to paragraph
func (p *Paragraph) AddAnchorDrawingFrom(file string) (*Run, error) {
	data, err := os.ReadFile(file)
	if err != nil {
		return nil, err
	}
	return p.AddAnchorDrawing(data)
}

// Size of the anchor drawing by EMU
func (r *WPAnchor) Size(w, h int64) {
	if r.Extent != nil {
		r.Extent.CX = w
		r.Extent.CY = h
	}
	if r.Graphic != nil && r.Graphic.GraphicData != nil && r.Graphic.GraphicData.Pic != nil && r.Graphic.GraphicData.Pic.SpPr != nil {
		r.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = w
		r.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CY = h
	}
}
