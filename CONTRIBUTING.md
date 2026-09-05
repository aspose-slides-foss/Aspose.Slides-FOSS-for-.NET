# Contributing

Thank you for considering a contribution. This document is specific to the .NET edition of
Aspose.Slides FOSS; the Java, C++ and Python editions live in their own repositories and have their
own conventions.

## The one rule that is particular to this project

**A fix to a writer ships with a test that asserts on the produced `.pptx` package — not on what the
library reads back.**

A library agrees with itself for free. If a writer emits `<a:tblStyleId>` where ECMA-376 says
`<a:tableStyleId>`, and the matching reader looks for `<a:tblStyleId>`, then every property reads
back exactly what was set, every round-trip test passes, and the file is still one that PowerPoint
quietly strips on load. That defect was real and it lived in this repository. The same shape of
mistake hides a relationship id that is never written into the `.rels`, a part with no content type,
and an effect element missing a required attribute: the object model is intact in memory, so a test
built on the object model sees nothing at all.

So: object-model tests are welcome and we have many, but they cannot close a writer bug. Put the
assertion in `tests/Aspose.Slides.Foss.ConformanceTests`, where it opens the file as a ZIP archive
and reads the XML.

## Prerequisites

The .NET 10.0 SDK, and nothing else — it builds both target frameworks. Running the tests needs
the .NET 8 runtime as well, because the suites execute on both. The library has no package dependencies and no native ones; the
test projects pull xunit, `Microsoft.NET.Test.Sdk`, `coverlet.collector`, `FluentAssertions` (unit
and integration suites) and `DocumentFormat.OpenXml` (conformance suite) from nuget.org.

**One of those is not under an open-source licence.** `FluentAssertions` 8.8.0 ships under the Xceed
Community License and sets `requireLicenseAcceptance`: free for open-source and non-commercial use,
paid for commercial use. It is a test-only dependency — it is not referenced by
`src/Aspose.Slides.Foss` and nothing in the published package touches it — but restoring this
repository does fetch it, so if you are working here inside a commercial setting, that is the term
to read. The conformance suite deliberately does not use it; see that project's `README.md`.

Package versions are managed centrally in `Directory.Packages.props`, so a `PackageReference` in a
project file carries no `Version` attribute — add the version there instead.

## Build

```bash
git clone https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET.git
cd Aspose.Slides-FOSS-for-.NET
dotnet build --configuration Release
```

Two things about this build worth knowing before your first pull request:

- **Warnings are errors.** `TreatWarningsAsErrors` is set for every project in
  `Directory.Build.props`. A build that prints nothing is a build that is clean, and CI additionally
  passes `-warnaserror` so that MSBuild's and NuGet's own warnings fail as well.
- **Public members must be documented.** `GenerateDocumentationFile` is on and CS1591 is *not*
  suppressed, so a new public type, method, property or field without an XML doc comment fails the
  build. The generated XML ships inside the NuGet package, so the comment you write is what a
  consumer sees in their editor.

## Test

```bash
dotnet test Aspose.Slides.Foss.sln --configuration Release
```

Three suites run, and all three must be green:

| Project | Tests | What it covers |
|---|---|---|
| `tests/Aspose.Slides.Foss.Tests` | 3,908 | unit tests of the object model |
| `tests/Aspose.Slides.Foss.IntegrationTests` | 194 | end-to-end use of the public API |
| `tests/Aspose.Slides.Foss.ConformanceTests` | 160 | assertions against the produced `.pptx` package |

Run one suite on its own with, for example:

```bash
dotnet test tests/Aspose.Slides.Foss.ConformanceTests --configuration Release
```

Test the whole solution rather than naming projects one at a time. Naming them means a suite added
later is silently never run.

The three counts above are not decoration: CI reads them back out of the `.trx` files after the test
run and fails if a suite produced no results at all, if anything did not pass, or if fewer tests
passed than the number in the table. They are floors — adding tests passes, so raise a number here
only when you want the new count guaranteed. This is what stops a suite dropped from
`Aspose.Slides.Foss.sln`, or skipped, from leaving a green run behind it.

### The conformance suite

`tests/Aspose.Slides.Foss.ConformanceTests/README.md` is the full description; the short version is
that `Harness/` gives you five helpers:

| Type | What it gives you |
|---|---|
| `PptxPackage` | opens a produced file as a `ZipArchive`: part names, part text, part XML, the relationships declared for a part, the content type a part resolves to |
| `PackageAssert` | the rules a consumer applies — every `r:id`/`r:embed`/`r:link` resolves in that part's `.rels`, every part is content-typed, no `Override` names a missing part, no relationship targets a missing part, an element exists at an XPath with the attributes you name, children appear in schema-required order |
| `SchemaValidation` | runs `OpenXmlValidator(FileFormatVersions.Office2019)` over the package and asserts zero errors |
| `TestWorkspace` | a temp directory deleted with the test class, plus a 1×1 PNG |
| `TestFixtures` | the decks to open, including a 46-part deck PowerPoint wrote that carries parts this library never writes — the only way to notice that a save dropped one |

The validator and the package assertions catch different things and neither subsumes the other. The
validator knows the schema: invented element names, missing required attributes, wrong child order.
It does not check that relationships resolve or that parts are content-typed, because those live in
the packaging layer below the schema. Use both when the defect is one a schema knows about.

`DocumentFormat.OpenXml` is referenced by the conformance project **only**. It must never become a
dependency of `src/Aspose.Slides.Foss`: this library writes OOXML by hand, and validating that output
with the same library that produced it would prove nothing.

## What a good pull request looks like

1. **It fixes one thing.** A pull request that repairs a writer and also renames three files is two
   pull requests.
2. **It has a failing test first.** Write the test against the unfixed code, watch it fail, and put
   that failure message in the pull request description. A test that has never failed has not been
   shown to test anything.
3. **The test names the user-visible failure**, in the words a user would use:
   `APictureFrameResolvesTheImageItPointsAt`, not `TestPictureFrame3`. The name is the bug report.
4. **It writes its file through the public API only.** If a case needs an internal member to be set
   up, it is testing the internals, not the artefact.
5. **It says what it changed about the file.** "Writes `<a:tableStyleId>` instead of
   `<a:tblStyleId>`" is reviewable. "Fixed table styles" is not.
6. **It updates `CHANGELOG.md`** under `## [Unreleased]`, in the language a caller would use, if the
   change is one a caller can observe.
7. **It documents new public members**, because otherwise it does not build.
8. **It does not add a dependency to `src/Aspose.Slides.Foss`.** "No package dependencies, no native
   dependencies" is a property this library advertises, and it is checked by the fact that the
   library's project file contains no `PackageReference` at all. If you believe a dependency is
   genuinely necessary, open an issue and make the case before writing the code.

### Behaviour that is deliberate, not a bug

Before filing a fix for one of these, please open an issue instead — they are decisions with
reasoning behind them, and the reasoning is in the source:

- `Save` raises `NotSupportedException` for the eighteen `SaveFormat` values it does not write. It
  will not write a presentation package under a name that claims to be PDF, ODP or a macro-enabled
  file.
- The modern `ppt/threadedComments/` part is not written. `IComment.ParentComment` answers from
  memory and the file carries the flat, classic comment list, which is what it really has.
- Removing a slide does not delete media that slide used.
- A section whose start slide comes before the previous section's start is clamped rather than
  refused.

## Commits and style

- Follow the layout and naming of the code you are changing; there is no separate style guide and no
  analyzer configuration beyond warnings-as-errors.
- Write the commit subject as a sentence saying what the change does to the software — the existing
  history is the model.
- Sign nothing off; there is no CLA and no DCO check.

## Continuous integration

Every push and pull request builds and tests on `ubuntu-latest`, `windows-latest` and
`macos-latest`, then packs the library and installs the resulting package into a throwaway project to
prove it is usable. Test results, the `.nupkg`, the symbol `.snupkg` and two CycloneDX SBOMs are
attached to the run as artefacts. If CI is red, the pull request is not ready, including when the
failure is on a platform you did not use.

## Releasing

A release is a tag push, and it stops for a human on the way: `.github/workflows/nuget-release.yml`
runs on a `v*` tag, and its publish job waits on a protected environment until a reviewer approves
it. No API key is stored in this repository — the workflow proves its identity to nuget.org and is
issued a key that lives for an hour.

**The full procedure is in [PUBLISHING.md](PUBLISHING.md)**: how to cut a release, what each guard
checks and why, how to verify a publication by fetching it rather than by reading a green tick, and
what to do when a release fails halfway. Read that rather than this.

Two things are worth repeating here, because they are what a contributor gets wrong:

- **Bumping `<Version>` in `Directory.Build.props` is one step and not the whole of it.** Several
  files state in the present tense that no package has been published. PUBLISHING.md carries the
  checklist; publishing without working through it puts "There is no NuGet package yet" on the page
  for the package that exists.
- **The page nuget.org renders is `docs/nuget/README.md`, not this repository's `README.md`.** They
  are different documents for different readers, and only the first one ships.

## Licence

By contributing you agree that your contribution is licensed under the
[MIT License](LICENSE), the same terms as the rest of the project.
