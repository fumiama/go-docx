package docxlib

import (
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
