package docx

import (
	"strconv"
	"sync/atomic"
)

// AddShape adds wsp named drawing to paragraph
func (p *Paragraph) AddShape(w, h int64, name, bwMode, prst string, elems []interface{}) (*Run, error) {
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
			CNvGraphicFramePr: &WPCNvGraphicFramePr{},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_WPS,
					Shape: &WPSWordprocessingShape{
						CNvCnPr: &WPSCNvCnPr{
							ConnShapeLocks: &struct{}{},
						},
						SpPr: &WPSSpPr{
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
							Elems:  elems,
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
	return run, nil
}
