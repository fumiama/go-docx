package docxlib

import "embed"

var (
	//go:embed xml
	//go:embed xml/_rels/*
	TEMP_XML_FS embed.FS
)
