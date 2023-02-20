package docxlib

// addImage add image to docx and return its rId
func (f *Docx) addImage(m Media) string {
	f.addMedia(m)
	return f.addImageRelation(m)
}
