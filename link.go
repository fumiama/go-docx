package docxlib

import (
	"strconv"
	"sync/atomic"
)

// when adding an hyperlink we need to store a reference in the relationship field
//
//	this func is not thread-safe
func (f *Docx) addLinkRelation(link string) string {
	rel := &Relationship{
		ID:         "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rId, 1))),
		Type:       REL_HYPERLINK,
		Target:     link,
		TargetMode: REL_TARGETMODE,
	}

	f.DocRelation.Relationships = append(f.DocRelation.Relationships, rel)

	return rel.ID
}

// when adding an image we need to store a reference in the relationship field
//
//	this func is not thread-safe
func (f *Docx) addImageRelation(m Media) string {
	rel := &Relationship{
		ID:     "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rId, 1))),
		Type:   REL_IMAGE,
		Target: "media/" + m.Name,
	}

	f.DocRelation.Relationships = append(f.DocRelation.Relationships, rel)

	return rel.ID
}

// ReferHref gets the url for a reference
func (f *Docx) ReferHref(id string) (href string, err error) {
	f.DocRelation.mu.RLock()
	defer f.DocRelation.mu.RUnlock()
	for _, a := range f.DocRelation.Relationships {
		if a.ID == id {
			href = a.Target
			return
		}
	}
	err = ErrRefIDNotFound
	return
}
