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

package docxlib

import "encoding/xml"

// AddTab adds tab to para
func (p *Paragraph) AddTab() *Run {
	run := &Run{
		RunProperties: &RunProperties{},
		FrontTab: []struct {
			XMLName xml.Name "xml:\"w:tab,omitempty\""
		}{{}},
	}
	p.Children = append(p.Children, run)
	return run
}

// AddText adds text to paragraph
func (p *Paragraph) AddText(text string) *Run {
	if text == "\t" {
		return p.AddTab()
	}

	t := &Text{
		Text: text,
	}

	run := &Run{
		Text:          t,
		RunProperties: &RunProperties{},
	}

	p.Children = append(p.Children, run)

	return run
}
