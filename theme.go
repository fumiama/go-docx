/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2024 Fumiama Minamoto (源文雨)

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
	"io/fs"
)

// UseTemplate will replace template files
func (f *Docx) UseTemplate(template string, tmpfslst []string, tmplfs fs.FS) *Docx {
	f.template = template
	f.tmplfs = tmplfs
	f.tmpfslst = tmpfslst
	return f
}

// WithDefaultTheme use default theme embeded
func (f *Docx) WithDefaultTheme() *Docx {
	return f.UseTemplate("default", DefaultTemplateFilesList, TemplateXMLFS)
}

// WithA3Theme use A3 theme embeded
func (f *Docx) WithA3Theme() *Docx {
	f.Document.Body.SectPr.PgSz.W = xml.Attr{Name: xml.Name{Local: "w:w"}, Value: "16840"}
	f.Document.Body.SectPr.PgSz.H = xml.Attr{Name: xml.Name{Local: "w:h"}, Value: "23820"}
	return f.UseTemplate("default", DefaultTemplateFilesList, TemplateXMLFS)
}
