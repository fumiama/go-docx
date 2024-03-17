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
	"reflect"
	"regexp"
	"strings"
)

//nolint:revive,stylecheck
const (
	XMLNS_W   = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_R   = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
	XMLNS_WP  = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
	XMLNS_WPS = `http://schemas.microsoft.com/office/word/2010/wordprocessingShape`
	XMLNS_WPC = `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas`
	XMLNS_WPG = `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup`
	XMLNS_MC  = `http://schemas.openxmlformats.org/markup-compatibility/2006`
	// XMLNS_WP14 = `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`

	XMLNS_O = `urn:schemas-microsoft-com:office:office`
	XMLNS_V = `urn:schemas-microsoft-com:vml`

	XMLNS_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
)

func getAtt(atts []xml.Attr, name string) string {
	for _, at := range atts {
		if at.Name.Local == name {
			return at.Value
		}
	}
	return ""
}

// Body <w:body>
type Body struct {
	Items []interface{}

	file *Docx
}

// UnmarshalXML ...
func (b *Body) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				value.file = b.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			case "tbl":
				var value Table
				value.file = b.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			case "sectPr":
				var value SectPr
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// KeepElements keep named elems amd removes others
//
// names: *docx.Paragraph *docx.Table
func (b *Body) KeepElements(name ...string) {
	items := make([]interface{}, 0, len(b.Items))
	namemap := make(map[string]struct{}, len(name)*2)
	for _, n := range name {
		namemap[n] = struct{}{}
	}
	for _, item := range b.Items {
		_, ok := namemap[reflect.ValueOf(item).Type().String()]
		if ok {
			items = append(items, item)
		}
	}
	b.Items = items
}

// DropDrawingOf drops all matched drawing in body
// name: Canvas, Shape, Group, ShapeAndCanvas, ShapeAndCanvasAndGroup, NilPicture
func (b *Body) DropDrawingOf(name string) {
	for _, item := range b.Items {
		switch o := item.(type) {
		case *Paragraph:
			f := reflect.ValueOf(o).MethodByName("Drop" + name)
			if !f.IsValid() {
				continue
			}
			_ = f.Call(nil)
		case *Table:
			for _, tr := range o.TableRows {
				for _, tc := range tr.TableCells {
					for _, p := range tc.Paragraphs {
						f := reflect.ValueOf(p).MethodByName("Drop" + name)
						if !f.IsValid() {
							continue
						}
						_ = f.Call(nil)
					}
				}
			}
		}
	}
}

// Document <w:document>
type Document struct {
	XMLName xml.Name `xml:"w:document"`
	XMLW    string   `xml:"xmlns:w,attr"`             // cannot be unmarshalled in
	XMLR    string   `xml:"xmlns:r,attr,omitempty"`   // cannot be unmarshalled in
	XMLWP   string   `xml:"xmlns:wp,attr,omitempty"`  // cannot be unmarshalled in
	XMLWPS  string   `xml:"xmlns:wps,attr,omitempty"` // cannot be unmarshalled in
	XMLWPC  string   `xml:"xmlns:wpc,attr,omitempty"` // cannot be unmarshalled in
	XMLWPG  string   `xml:"xmlns:wpg,attr,omitempty"` // cannot be unmarshalled in
	// XMLMC   string   `xml:"xmlns:mc,attr,omitempty"`  // cannot be unmarshalled in
	// XMLWP14 string   `xml:"xmlns:wp14,attr,omitempty"` // cannot be unmarshalled in

	// XMLO string `xml:"xmlns:o,attr,omitempty"` // cannot be unmarshalled in
	// XMLV string `xml:"xmlns:v,attr,omitempty"` // cannot be unmarshalled in

	// MCIgnorable string `xml:"mc:Ignorable,attr,omitempty"`

	Body Body `xml:"w:body"`
}

// UnmarshalXML ...
func (doc *Document) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "body" {
				err = d.DecodeElement(&doc.Body, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// ParagraphSplitRule check whether the paragraph is a separator or not
type ParagraphSplitRule func(*Paragraph) bool

// SplitDocxByPlainTextRegex matches p.String()
func SplitDocxByPlainTextRegex(re *regexp.Regexp) ParagraphSplitRule {
	return func(p *Paragraph) bool {
		return re.MatchString(p.String())
	}
}

// SplitByParagraph splits a doc to many docs by using a matched paragraph
// as the separator.
//
// The separator will be placed to the first doc item
func (f *Docx) SplitByParagraph(separator ParagraphSplitRule) (docs []*Docx) {
	items := f.Document.Body.Items
newdoclop:
	for len(items) > 0 {
		ndoc := new(Docx)

		// migrate base data
		ndoc.mediaNameIdx = make(map[string]int, 64)
		ndoc.slowIDs = make(map[string]uintptr, 64)
		ndoc.template = f.template
		ndoc.tmplfs = f.tmplfs
		ndoc.tmpfslst = f.tmpfslst

		ndoc.Document.XMLW = XMLNS_W
		ndoc.Document.XMLR = XMLNS_R
		ndoc.Document.XMLWP = XMLNS_WP
		// ndoc.Document.XMLMC = XMLNS_MC
		// ndoc.Document.XMLO = XMLNS_O
		// ndoc.Document.XMLV = XMLNS_V
		ndoc.Document.XMLWPS = XMLNS_WPS
		ndoc.Document.XMLWPC = XMLNS_WPC
		ndoc.Document.XMLWPG = XMLNS_WPG
		// ndoc.Document.XMLWP14 = XMLNS_WP14
		ndoc.Document.XMLName.Space = XMLNS_W
		ndoc.Document.XMLName.Local = "document"
		ndoc.Document.Body.file = ndoc

		ndoc.docRelation = Relationships{
			Xmlns: XMLNS_REL,
			Relationship: []Relationship{
				{
					ID:     "rId1",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`,
					Target: "styles.xml",
				},
				{
					ID:     "rId2",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme`,
					Target: "theme/theme1.xml",
				},
				{
					ID:     "rId3",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable`,
					Target: "fontTable.xml",
				},
			},
		}

		ndoc.rID = 3

		for i, item := range items {
			switch o := item.(type) {
			case *Paragraph:
				if separator(o) && len(ndoc.Document.Body.Items) > 0 {
					items = items[i:]
					docs = append(docs, ndoc)
					continue newdoclop
				}
				np := o.copymedia(ndoc)
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, &np)
			case *Table:
				nt := o.copymedia(ndoc)
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, &nt)
			default:
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, o)
			}
		}

		if len(ndoc.Document.Body.Items) > 0 {
			docs = append(docs, ndoc)
		}
		break
	}
	return
}

func (r *Run) copymedia(to *Docx) *Run {
	nr := *r
	nr.Children = make([]interface{}, 0, len(r.Children))
	nr.file = to
	for _, rc := range r.Children {
		if d, ok := rc.(*Drawing); ok {
			nr.Children = append(nr.Children, d.copymedia(to))
			continue
		}
		nr.Children = append(nr.Children, rc)
	}
	return &nr
}

func (p *Paragraph) copymedia(to *Docx) (np Paragraph) {
	np = *p
	np.Children = make([]interface{}, 0, len(p.Children))
	np.file = to
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			np.Children = append(np.Children, r.copymedia(to))
			continue
		}
		if h, ok := pc.(*Hyperlink); ok {
			tgt, err := p.file.ReferTarget(h.ID)
			if err != nil {
				continue
			}
			rid := to.addLinkRelation(tgt)
			np.Children = append(np.Children, &Hyperlink{
				ID:  rid,
				Run: *h.Run.copymedia(to),
			})
			continue
		}
		np.Children = append(np.Children, pc)
	}
	return
}

func (t *Table) copymedia(to *Docx) (nt Table) {
	nt = *t
	nt.TableRows = make([]*WTableRow, 0, len(t.TableRows))
	nt.file = to
	for _, tr := range t.TableRows {
		ntr := *tr
		ntr.TableCells = make([]*WTableCell, 0, len(tr.TableCells))
		ntr.file = to
		for _, tc := range tr.TableCells {
			ntc := *tc
			ntc.Paragraphs = make([]*Paragraph, 0, len(tc.Paragraphs))
			ntc.file = to
			for _, p := range tc.Paragraphs {
				np := p.copymedia(to)
				ntc.Paragraphs = append(ntc.Paragraphs, &np)
			}
			ntr.TableCells = append(ntr.TableCells, &ntc)
		}
		nt.TableRows = append(nt.TableRows, &ntr)
	}
	return
}

// AppendFile appends all contents in af to f
func (f *Docx) AppendFile(af *Docx) {
	for _, item := range af.Document.Body.Items {
		switch o := item.(type) {
		case *Paragraph:
			np := o.copymedia(f)
			f.Document.Body.Items = append(f.Document.Body.Items, &np)
		case *Table:
			nt := o.copymedia(f)
			f.Document.Body.Items = append(f.Document.Body.Items, &nt)
		default:
			f.Document.Body.Items = append(f.Document.Body.Items, o)
		}
	}
}
