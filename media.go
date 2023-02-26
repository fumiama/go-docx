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
