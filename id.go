package docx

func (f *Docx) IncreaseID(name string) (n uintptr) {
	f.slowIDsMu.Lock()
	n, _ = f.slowIDs[name] //nolint: go-staticcheck
	n++
	f.slowIDs[name] = n
	f.slowIDsMu.Unlock()
	return
}
