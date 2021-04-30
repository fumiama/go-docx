# Docx library

Yet another library to read and write .docx (Microsoft Word) files in Go.

## Introduction

As part of my work for [Basement Crowd](https://www.basementcrowd.com) y [FromCounsel](https://www.fromcounsel.com), we were in need of a basic library to manipulate (both read and write) Microsoft Word documents.

The difference with other projects is the following:

- [UniOffice](https://github.com/unidoc/unioffice) is probably the most complete but it is also commercial (you need to pay). It also very complete, but too much for my needs.

- [gingfrederik/docx](https://github.com/gingfrederik/docx) only allows to write.

There are also a couple of other projects [kingzbauer/docx](https://github.com/kingzbauer/docx) and [nguyenthenguyen/docx](https://github.com/nguyenthenguyen/docx)

[gingfrederik/docx](https://github.com/gingfrederik/docx) was a heavy influence (the original structures and the main method come from that project).

However, the structures didn't handle reading and extending them was particularly difficult due to Go xml parser being limited and [6 year old bug](https://github.com/golang/go/issues/9519).

Additionally, my requirements go beyond the original structure and a hard fork seemed more sensible.

The plan is to evolve the library, so the API is likely to change according to my company's needs. But please do feel free to send patches, reports and PRs or fork.

In the mean time, shared as an example.

## Getting Started

### Install

Go modules supported

```sh
go get github.com/gonfva/docxlib
```

### Usage

See [main](main/main.go) for an example

```
$ go build -o docxlib ./main
$ ./docxlib
Preparing new document to write at /tmp/new-file.docx
Document writen.
Now trying to read it
	We've found a new run with the text ->test
	We've found a new run with the text ->test font size
	We've found a new run with the text ->test color
	We've found a new run with the text ->test font size and color
	We've found a new hyperlink with ref http://google.com and the text google
End of main
```

### Build

```
$ go build ./...
```

## License

MIT. See [LICENSE](LICENSE)
