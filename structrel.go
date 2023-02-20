package docxlib

import (
	"encoding/xml"
	"sync"
)

const (
	XMLNS_REL     = `http://schemas.openxmlformats.org/package/2006/relationships`
	REL_HYPERLINK = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`
	REL_IMAGE     = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

	REL_TARGETMODE = "External"
)

type Relationships struct {
	mu            sync.RWMutex
	XMLName       xml.Name        `xml:"Relationships"`
	Xmlns         string          `xml:"xmlns,attr"`
	Relationships []*Relationship `xml:"Relationship"`
}

type Relationship struct {
	XMLName    xml.Name `xml:"Relationship"`
	ID         string   `xml:"Id,attr"`
	Type       string   `xml:"Type,attr"`
	Target     string   `xml:"Target,attr"`
	TargetMode string   `xml:"TargetMode,attr,omitempty"`
}
