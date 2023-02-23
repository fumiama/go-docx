package docxlib

//nolint:revive,stylecheck
const MEDIA_FOLDER = `word/media/`

// Media is in word/media
type Media struct {
	Name string // Name is for word/media/Name
	Data []byte // Data is data of this media
}

// String is the full path of the media
func (m *Media) String() string {
	return MEDIA_FOLDER + m.Name
}

// Media get media struct pointer (or nil on notfound) by name
func (f *Docx) Media(name string) *Media {
	i, ok := f.mediaNameIdx[name]
	if !ok {
		return nil
	}
	return &f.media[i]
}

// addMedia append the media to docx's media list
func (f *Docx) addMedia(m Media) {
	f.mediaNameIdx[m.Name] = len(f.media)
	f.media = append(f.media, m)
}
