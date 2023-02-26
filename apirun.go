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

// Shade allows to set run shade
func (r *Run) Shade(val, color, fill string) *Run {
	r.RunProperties.Shade = &Shade{
		Val:   val,
		Color: color,
		Fill:  fill,
	}
	return r
}

// Bold ...
func (r *Run) Bold() *Run {
	r.RunProperties.Bold = &Bold{}
	return r
}

// Italic ...
func (r *Run) Italic() *Run {
	r.RunProperties.Italic = &Italic{}
	return r
}

// Underline has several possible values including
//
//	none: Specifies that no underline should be applied.
//	single: Specifies a single underline.
//	words: Specifies that only words within the text should be underlined.
//	double: Specifies a double underline.
//	thick: Specifies a thick underline.
//	dotted: Specifies a dotted underline.
//	dash: Specifies a dash underline.
//	dotDash: Specifies an alternating dot-dash underline.
//	dotDotDash: Specifies an alternating dot-dot-dash underline.
//	wave: Specifies a wavy underline.
//	dashLong: Specifies a long dash underline.
//	wavyDouble: Specifies a double wavy underline.
func (r *Run) Underline(val string) *Run {
	r.RunProperties.Underline = &Underline{Val: val}
	return r
}

// Highlight ...
func (r *Run) Highlight(val string) *Run {
	r.RunProperties.Highlight = &Highlight{Val: val}
	return r
}

// AddTab add a tab in front of the run
func (r *Run) AddTab() *Run {
	r.Children = append(r.Children, &Tab{})
	return r
}

// Font sets the font of the run
func (r *Run) Font(ascii, hansi, hint string) *Run {
	r.RunProperties.Fonts = &RunFonts{
		ASCII: ascii,
		HAnsi: hansi,
		Hint:  hint,
	}
	return r
}
