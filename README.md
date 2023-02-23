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

- [x] parse and save document
- [x] edit simple text (color, size, alignment, ...)
- [x] edit picture

## License

MIT. See [LICENSE](LICENSE)
