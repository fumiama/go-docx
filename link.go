package docxlib

import (
	"errors"
	"strconv"
	"sync/atomic"
)

var (
	// ErrRefIDNotFound cannot find such reference
	ErrRefIDNotFound = errors.New("ref id not found")
	// ErrRefTargetNotFound cannot find such target
	ErrRefTargetNotFound = errors.New("ref target not found")
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

// ReferTarget gets the target for a reference
func (f *Docx) ReferTarget(id string) (string, error) {
	f.DocRelation.mu.RLock()
	defer f.DocRelation.mu.RUnlock()
	for _, a := range f.DocRelation.Relationships {
		if a.ID == id {
			return a.Target, nil
		}
	}
	return "", ErrRefIDNotFound
}

// ReferID gets the rId from target
func (f *Docx) ReferID(target string) (string, error) {
	f.DocRelation.mu.RLock()
	defer f.DocRelation.mu.RUnlock()
	for _, a := range f.DocRelation.Relationships {
		if a.Target == target {
			return a.ID, nil
		}
	}
	return "", ErrRefIDNotFound
}
