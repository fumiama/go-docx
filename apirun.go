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

// Color allows to set run color
func (r *Run) Color(color string) *Run {
	r.RunProperties.Color = &Color{
		Val: color,
	}

	return r
}

// Size allows to set run size
func (r *Run) Size(size string) *Run {
	r.RunProperties.Size = &Size{
		Val: size,
	}

	return r
}

// AddTab add a tab in front of the run
func (r *Run) AddTab() *Run {
	r.FrontTab = append(r.FrontTab, struct {
		XMLName xml.Name "xml:\"w:tab,omitempty\""
	}{})
	return r
}

// AppendTab add a tab after the run
func (r *Run) AppendTab() *Run {
	r.RearTab = append(r.RearTab, struct {
		XMLName xml.Name "xml:\"w:tab,omitempty\""
	}{})
	return r
}
