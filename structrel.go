package docxlib

import (
	"sync"
)

//nolint:revive,stylecheck
const (
	XMLNS_REL     = `http://schemas.openxmlformats.org/package/2006/relationships`
	REL_HYPERLINK = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`
	REL_IMAGE     = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

	REL_TARGETMODE = "External"
)

// Relationships ...
type Relationships struct {
	mu           sync.RWMutex
	Xmlns        string `xml:"xmlns,attr"`
	Relationship []Relationship
}

// Relationship ...
type Relationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}
