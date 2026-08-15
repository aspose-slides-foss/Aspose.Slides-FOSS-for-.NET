# Conformance tests

These tests answer one question: **is the `.pptx` this library writes a file that other software
can read?**

They do it the only way that question can be answered — by writing a file through the public API,
opening it as a ZIP archive, and asserting on the XML inside. Nothing in this project asks the
library to read its own output back.

## Why not just read it back?

Because a library agrees with itself for free.

If a writer puts `<a:tblStyleId>` where ECMA-376 says `<a:tableStyleId>`, and the matching reader
looks for `<a:tblStyleId>`, then every round-trip test passes, every property reads back exactly what
was set, and the file is still one PowerPoint quietly strips on load. The same holds for a
relationship id that is never written to the `.rels`, a part with no content type, and an effect
element with a required attribute missing: the object model is intact in memory, so a test built on
the object model sees nothing at all.

The rest of this repository's tests are written against the object model, and they are worth having —
they check that the API behaves. These check the artefact. Both are needed, and only one of them can
tell you whether a user's file opens.

**The rule, with no exceptions: assert on the produced package.** A conformance test that calls back
into `Aspose.Slides.Foss` to inspect its own output is not a conformance test.

## What the harness checks

`Harness/` holds four pieces, all of them ordinary helpers you call from a `[Fact]`:

| Type | What it gives you |
|---|---|
| `PptxPackage` | Opens a produced file as a `ZipArchive`. Part names, part text, part XML, the relationships declared for a part, and the content type a part resolves to. |
| `PackageAssert` | The rules a consumer applies: every `r:id`/`r:embed`/`r:link` resolves in that part's `.rels`; every part resolves the content type ECMA-376 gives it; no `Override` names a missing part; no relationship targets a missing part; an XPath-addressed element exists with the attributes you name; children appear in schema-required order. |
| `SchemaValidation` | Runs `OpenXmlValidator(FileFormatVersions.Office2019)` over the package and asserts zero errors. A package a strict reader refuses to open at all is reported as one error carrying the reason. |
| `TestWorkspace` | A temp directory for produced files, deleted with the test class, plus a 1×1 PNG for image cases. |

The two layers catch different things and neither subsumes the other. The validator finds invented
names, missing required attributes and wrong child order — it knows the schema. It does *not* check
that relationships resolve or that parts are content-typed, because those live in the packaging layer
below the schema. The package assertions cover exactly that gap.

## Dependencies

`DocumentFormat.OpenXml` (MIT, Microsoft) is referenced **by this test project only**. It must never
become a dependency of `src/Aspose.Slides.Foss`: this library writes OOXML by hand, and validating
that output with the same code that produced it would prove nothing.

This project does not use `FluentAssertions`; plain xUnit assertions with an explicit failure message
are enough, and keep the conformance suite free of a dependency whose licence terms downstream forks
have to think about.

## Reading a failure

Failure messages are written to be actionable on their own — they print what was written, not just
that an assertion failed:

```
1 relationship reference(s) do not resolve:
ppt/slides/slide1.xml: <blip r:embed="rId2"> is not declared in ppt/slides/_rels/slide1.xml.rels (declared: rId1)
```

If you cannot tell from the message which part of the package is wrong, the message needs improving —
that counts as a defect in the harness.

## Adding a case

1. **Name the test after the user-visible failure**, in the words a user would use:
   `APictureFrameResolvesTheImageItPointsAt`, not `TestPictureFrame3`. The name is the bug report.
2. **Write a file through the public API only.** If a case needs an internal to be set up, it is
   testing the internals, not the artefact.
3. **Open it with `PptxPackage.Open` and assert.** Reach for the `PackageAssert` rule first; add a
   new rule there when the same check would apply to any package, and keep case-specific XPath in the
   test.
4. **Validate as well as assert**, whenever the defect is one a schema knows about. The two layers
   fail differently and the pair localises the problem faster than either alone.
5. **Run the test against the unfixed code before you fix anything**, and keep the failure message in
   the commit or PR description. A conformance test that has never failed has not been shown to test
   anything.

Files are grouped by the surface they exercise — `PictureConformanceTests`, `EffectConformanceTests`,
and so on. `PackageRegressionTests` is the exception: everything in it passes today, and it is there
so that a repair to one writer cannot silently break a neighbouring one.

## Running them

```
dotnet test tests/Aspose.Slides.Foss.ConformanceTests
```

Some of these tests fail today. That is what they are for: each one names a defect in a writer, and a
failing conformance test is a defect report with the evidence attached.
