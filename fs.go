package docxlib

import "embed"

var (
	// TemplateXMLFS stores template docx files
	//go:embed xml
	//go:embed xml/a4/_rels/*
	TemplateXMLFS embed.FS
)
