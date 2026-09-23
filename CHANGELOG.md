# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
uses a CalVer version scheme — `YY.M.PATCH`, matching the other language editions of the library.

## [Unreleased]

### Fixed

- **Removing the first slide of a section no longer moves the section before it.** With sections
  Intro (A) and Body (B, C), removing B wrote Intro with no slides and Body with A and C. The section
  now starts from its next slide, which gives Intro (A) and Body (C) — what PowerPoint saves for the
  same deletion. A section whose only slide is removed stays in the list, empty and in its place.
- **Editing the slides of a deck opened from a file keeps its sections in step.** Adding, inserting,
  cloning or removing a slide renumbers the slide ids that sections refer to, and on an opened deck
  the section list was written back unchanged unless `Sections` itself had been edited: new slides
  were in no section, and sections named slides that no longer existed. The section list is now
  rewritten after any change to the slides, from the slides each section holds. An empty section read
  from a file keeps its place instead of emptying every section after it.
- **`ISlide.Remove()` removes the slide from the file.** It took the slide out of the object model
  only, so `Slides.Count` went down while the saved deck still contained the slide. It now does what
  `Slides.Remove(slide)` does.
- **A subset save lists only the slides it writes in its sections.** `Save(…, int[] slides, …)` left
  the dropped slides' ids in the section list. They are now removed; a section left with no slides is
  kept, empty, as PowerPoint keeps it.
- **A subset save's `docProps/app.xml` describes the file written.** It reported the slide count,
  notes, hidden slides and slide titles of the whole deck — `<Slides>3</Slides>` in a one-slide file.
- **Removing a slide removes its comments from the file.** When the comments had been read or added
  in the same session, the removed slide's comments part — comment text included — and a
  relationships part for the removed slide were written back on save, while the object model
  reported no such comments.
- **An author with no comments is written to a part the package declares.** `ppt/commentAuthors.xml`
  was written without a content type or a relationship unless a comment was written too — which is
  also what happened once the only commented slide was removed. Reading `CommentAuthors` on a deck
  that has none no longer adds the part at all.

### Changed

- **`ISection.StartedFromSlide` follows removals.** When the slide a section starts from is removed,
  it becomes the next slide of that section, or `null` if the section has no slides left. It
  previously kept pointing at the removed slide.
- **A section with no start slide sits in its place in the section order.** Such a section — empty
  after removals, read empty from a file, or appended with `AppendEmptySection` at an index past the
  last slide — now ends where the next section begins, so the sections after it keep their slides.
  It was previously placed after the last slide, which left every later section empty.
- **Reading sections no longer causes them to be rewritten.** Opening a deck and reading `Sections`
  used to mark the section list as changed, so it was regenerated on save even when nothing had been
  edited; it is now kept as the file had it until a section or the slide list changes.

## [26.9.0] - 2026-09-10

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
- **`Color` compares by component with `==` as well as with `Equals`.** `Color` is a class, and
  without an equality operator `Color.FromArgb(255, 255, 0, 0) == Color.Red` was `false` while
  `.Equals` on the same pair was `true`. `==` and `!=` now answer the same as `Equals`.

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
  test results and two CycloneDX SBOMs as run artefacts. The run also reads those test results back
  and fails if a suite produced none, or ran fewer tests than the count recorded in
  `CONTRIBUTING.md`, so a suite dropped from the solution or skipped cannot leave a green run behind
  it.
- **Community documentation** — this changelog, `CONTRIBUTING.md`, `SECURITY.md`,
  `CODE_OF_CONDUCT.md`, issue forms and a pull request template, and a rewritten `README.md` that
  states what the library cannot do as plainly as what it can.

### Known limitations

Not defects, and not scheduled: charts, SmartArt, OLE objects, video and audio, group shapes,
animations, slide transitions, hyperlinks, slide backgrounds, themes, an API for the slide size,
adding or cloning a layout, threaded comments, encryption, macros, digital signatures, and any form
of rendering or conversion.
The README lists these in full, with the API member that is missing for each.

[26.9.0]: https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/releases/tag/v26.9.0
