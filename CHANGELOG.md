# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
uses a CalVer version scheme — `YY.M.PATCH`, matching the other language editions of the library.

Nothing has been published to nuget.org yet, so everything below is unreleased.

## [Unreleased]

This is the first entry. The library existed before it, but it wrote files that were wrong in ways
its own tests could not see, and the work below is what made the produced `.pptx` a package other
software can read. **If you built against an earlier commit, read the Changed section first — calls
that used to succeed now raise.**

### Changed

- **`Save` refuses the formats it cannot write instead of silently writing a `.pptx`.** `SaveFormat`
  declares 21 values; three of them — `Pptx`, `Ppsx` and `Potx` — are written, and the other 18 now
  raise `NotSupportedException` with a message naming the three that work. Previously every value
  produced a presentation package under whatever file name you gave it, so asking for PDF, ODP or a
  macro-enabled file wrote a mislabelled presentation and reported success.
- **`Ppsx` and `Potx` produce genuinely different packages.** Each of the three written formats now
  declares its own main-part content type — slideshow, template and presentation respectively —
  rather than all three writing the presentation content type.
- **Using a `Presentation` after `Dispose` raises `ObjectDisposedException`.** It previously wrote a
  zero-byte file over the target path.
- **`Save` raises when there is no package to write** rather than returning as though it had
  succeeded.
- **A rejected format no longer truncates the target file.** `Save(string, SaveFormat)` validates the
  format before it creates the file, so a file already at that path survives a refused save.
- **The slide-subset overloads of `Save` honour their `int[] slides` argument.** They previously
  ignored it and wrote every slide.
- **`Sections.AddSection` rejects a slide that belongs to another presentation** instead of accepting
  it and producing a section that lists nothing.
- **A section whose start slide precedes the previous section's start is clamped**, not refused, so
  that a reorder in progress cannot throw. A clamped section has no slides.

### Fixed

- **Table style references are written as `<a:tableStyleId>`**, the element ECMA-376 defines, instead
  of `<a:tblStyleId>`, and before the extension list rather than after it so the child order is
  valid. PowerPoint discarded table styling written the old way.
- **Embedded images resolve.** Adding a picture now writes both the relationship in the slide's
  `.rels` and the content type for the media part, so `<a:blip r:embed="…">` points at something that
  exists.
- **Every slide and notes-slide part is declared in `[Content_Types].xml`.** A part with no content
  type makes the package unreadable to a strict consumer.
- **The one-line effect enablers write complete elements.** `EnableOuterShadowEffect` and its seven
  siblings previously produced elements missing attributes the schema requires.
- **Removing a slide removes its part from the package**, not just its entry in the slide list.
  Media the slide used is deliberately left in place; see the note in `PackageSlides`.
- **Sections reach the file.** They were held in memory and never written into
  `ppt/presentation.xml`.
- **`<p:sldIdLst>` keeps its required position** after every master list in `ppt/presentation.xml`.
- **A notes slide is written with the notes master it requires**, and the notes master gets its own
  theme part rather than borrowing the slide master's.
- **Document property parts are declared and related.** `docProps/core.xml`, `docProps/app.xml` and
  `docProps/custom.xml` are written, content-typed and related from the package root;
  `docProps/app.xml` is regenerated on each save and `dcterms:modified` is stamped.
- **Comments no longer carry a `parentCmId` attribute**, which is not in the schema and which no
  consumer renders a thread from. `IComment.ParentComment` answers from memory; the file carries the
  flat comment list it really has.
- **A paragraph added through `ITextFrame.Paragraphs` is written into the slide.** It was previously
  visible in the object model and absent from the file.

### Added

- **A conformance test suite** — 80 tests that write a file through the public API, open it as a ZIP
  archive and assert on the XML inside, never asking the library to read its own output back. It
  runs as a required part of the build. See
  `tests/Aspose.Slides.Foss.ConformanceTests/README.md`.
- **A part-by-part round-trip test** against a 46-part deck PowerPoint wrote, which pins that opening
  and saving a foreign presentation keeps every part.
- **XML documentation for the whole public surface**, generated into the package so editors show it,
  with missing-documentation warnings left as build errors so it stays complete.
- **NuGet packaging metadata** — package id, description, tags, licence expression, project and
  repository URLs, the readme and the licence file inside the package, a symbol `.snupkg`, Source
  Link and a deterministic build.
- **Continuous integration** — build and test on Linux, Windows and macOS; pack the library, install
  the produced package into a throwaway project and run it; publish the package, the symbols, the
  test results and two CycloneDX SBOMs as run artefacts.
- **Community documentation** — this changelog, `CONTRIBUTING.md`, `SECURITY.md`,
  `CODE_OF_CONDUCT.md`, issue forms and a pull request template, and a rewritten `README.md` that
  states what the library cannot do as plainly as what it can.

### Known limitations

Not defects, and not scheduled: charts, SmartArt, OLE objects, video and audio, group shapes,
animations, slide transitions, hyperlinks, slide backgrounds, themes, an API for the slide size,
threaded comments, encryption, macros, digital signatures, and any form of rendering or conversion.
The README lists these in full, with the API member that is missing for each.

[Unreleased]: https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/commits/main
