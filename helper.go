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

import (
	"fmt"
	"strconv"
	"unsafe"
)

// BytesToString 没有内存开销的转换
func BytesToString(b []byte) string {
	return *(*string)(unsafe.Pointer(&b))
}

// StringToBytes 没有内存开销的转换
func StringToBytes(s string) (b []byte) {
	bh := (*slice)(unsafe.Pointer(&b))
	sh := (*slice)(unsafe.Pointer(&s))
	bh.data = sh.data
	bh.len = sh.len
	bh.cap = sh.len
	return b
}

// GetInt64 from string
func GetInt64(s string) (int64, error) {
	v, err := strconv.ParseInt(s, 10, 64)
	if err == nil {
		return v, nil
	}
	v2, err := strconv.ParseFloat(s, 64)
	if err == nil {
		return int64(v2), nil
	}
	_, err = fmt.Sscanf(s, "%d", &v)
	return v, err
}

// GetInt from string
func GetInt(s string) (int, error) {
	v, err := strconv.Atoi(s)
	if err == nil {
		return v, nil
	}
	v2, err := strconv.ParseFloat(s, 64)
	if err == nil {
		return int(v2), nil
	}
	_, err = fmt.Sscanf(s, "%d", &v)
	return v, err
}
