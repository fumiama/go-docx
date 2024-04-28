# Docx library

One of the most functional libraries to read and write .docx (a.k.a. Microsoft Word documents or ECMA-376 Office Open XML) files in Go.

This is a variant optimized and expanded by fumiama. The original repo is [gonfva/docxlib](https://github.com/gonfva/docxlib).

## Introduction

> As part of my work for [Basement Crowd](https://www.basementcrowd.com) and [FromCounsel](https://www.fromcounsel.com), we were in need of a basic library to manipulate (both read and write) Microsoft Word documents.
> 
> The difference with other projects is the following:
> - [UniOffice](https://github.com/unidoc/unioffice) is probably the most complete but it is also commercial (you need to pay). It also very complete, but too much for my needs.
> - [gingfrederik/docx](https://github.com/gingfrederik/docx) only allows to write.
> 
> There are also a couple of other projects [kingzbauer/docx](https://github.com/kingzbauer/docx) and [nguyenthenguyen/docx](https://github.com/nguyenthenguyen/docx)
> 
> [gingfrederik/docx](https://github.com/gingfrederik/docx) was a heavy influence (the original structures and the main method come from that project).
> 
> However, those original structures didn't handle reading and extending them was particularly difficult due to Go xml parser being a bit limited including a [6 year old bug](https://github.com/golang/go/issues/9519).
> 
> Additionally, my requirements go beyond the original structure and a hard fork seemed more sensible.
> 
> The plan is to evolve the library, so the API is likely to change according to my company's needs. But please do feel free to send patches, reports and PRs (or fork).
> 
> In the mean time, shared as an example in case somebody finds it useful.

The Introduction above is copied from the original repo. I had evolved that repo again to fit my needs. Here are the supported functions now.

- [x] Parse and save document
- [x] Edit text (color, size, alignment, link, ...)
- [x] Edit picture
- [x] Edit table
- [x] Edit shape
- [x] Edit canvas
- [x] Edit group

## Quick Start
```bash
go run cmd/main/main.go -u
```
And you will see two files generated under `pwd` with the same contents as below.

<table>
	<tr>
		<td align="center"><img src="https://user-images.githubusercontent.com/41315874/223348099-4a6099d2-0fec-4e13-92a7-152c00bc6f6b.png"></td>
		<td align="center"><img src="https://user-images.githubusercontent.com/41315874/223349486-e78ac0f1-c879-4888-9110-ea4db2590241.png"></td>
	</tr>
	<tr>
		<td align="center">p1</td>
		<td align="center">p2</td>
	</tr>
</table>

## Use Package in your Project
```bash
go get -d github.com/fumiama/go-docx@latest
```
### Generate Document
```go
package main

import (
	"os"
	"strings"

	"github.com/fumiama/go-docx"
)

func main() {
	w := docx.New().WithDefaultTheme()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("test").AddTab()
	para1.AddText("size").Size("44").AddTab()
	f, err := os.Create("generated.docx")
	// save to file
	if err != nil {
		panic(err)
	}
	_, err = w.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
}
```
### Parse Document
```go
package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/fumiama/go-docx"
)

func main() {
	readFile, err := os.Open("file2parse.docx")
	if err != nil {
		panic(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)
	if err != nil {
		panic(err)
	}
	fmt.Println("Plain text:")
	for _, it := range doc.Document.Body.Items {
		switch it.(type) {
		case *docx.Paragraph, *docx.Table: // printable
			fmt.Println(it)
		}
	}
}
```

## License

AGPL-3.0. See [LICENSE](LICENSE)
