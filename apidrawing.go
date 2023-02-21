package docxlib

import (
	"bytes"
	"fmt"
	"math/rand"
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
	id := strconv.Itoa(int(atomic.AddUintptr(&p.file.imageId, 1)))
	rId := p.file.addImage(Media{Name: "image" + id + "." + format, Data: pic})
	w, h := sz.Width, sz.Height
	if float64(w)/float64(h) > 1.2 {
		h = A4_EMU_MAX_WIDTH * h / w
		w = A4_EMU_MAX_WIDTH
	} else {
		h = A4_EMU_MAX_WIDTH * h / w / 2
		w = A4_EMU_MAX_WIDTH / 2
	}
	d := &Drawing{
		Inline: &WPInline{
			AnchorID: fmt.Sprintf("%08X", rand.Uint32()),
			EditID:   fmt.Sprintf("%08X", rand.Uint32()),

			Extent: &WPExtent{
				CX: w,
				CY: h,
			},
			EffectExtent: &WPEffectExtent{},
			DocPr: &WPDocPr{
				ID:   id,
				Name: "图片 " + id,
			},
			CNvGraphicFramePr: &WPCNvGraphicFramePr{
				Locks: &AGraphicFrameLocks{
					NoChangeAspect: 1,
				},
			},
			Graphic: &AGraphic{
				XMLA: XMLNS_DRAWINGML_MAIN,
				GraphicData: &AGraphicData{
					URI: XMLNS_PICTURE,
					Pic: &PICPic{
						XMLPIC: XMLNS_DRAWINGML_PICTURE,
						NonVisualPicProperties: &PICNonVisualPicProperties{
							NonVisualDrawingProperties: PICNonVisualDrawingProperties{
								ID: id,
							},
						},
						BlipFill: &PICBlipFill{
							Blip: ABlip{
								Embed:  rId,
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
							PrstGeom: APrstGeom{
								Prst: "rect",
							},
						},
					},
				},
			},
		},
	}
	run := &Run{
		Drawing:       d,
		RunProperties: &RunProperties{},
	}
	p.Children = append(p.Children, ParagraphChild{Run: run})
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
func (in *WPInline) Size(w, h int) {
	if in.Extent != nil {
		in.Extent.CX = w
		in.Extent.CY = h
	}
	if in.Graphic != nil && in.Graphic.GraphicData != nil && in.Graphic.GraphicData.Pic != nil && in.Graphic.GraphicData.Pic.SpPr != nil {
		in.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = w
		in.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CY = h
	}
}
