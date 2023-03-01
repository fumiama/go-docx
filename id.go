package docx

func (f *Docx) IncreaseID(name string) (n uintptr) {
	f.slowIDsMu.Lock()
	n = f.slowIDs[name]
	n++
	f.slowIDs[name] = n
	f.slowIDsMu.Unlock()
	return
}
