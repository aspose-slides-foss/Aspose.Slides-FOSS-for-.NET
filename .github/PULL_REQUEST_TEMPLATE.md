<!--
Thank you for the pull request. CONTRIBUTING.md has the detail; this template is the short form.
Delete any section that genuinely does not apply, rather than leaving it blank.
-->

## What this changes

<!-- One or two sentences. If it changes what ends up in the .pptx, name the part and the element:
     "writes <a:tableStyleId> instead of <a:tblStyleId> in ppt/slides/slide1.xml". -->

Closes #

## Why

<!-- The user-visible problem. What did someone see happen, and what did they expect? -->

## How it was verified

<!-- Paste the failing test output from BEFORE the fix, and the passing run after it. A test that
     has never failed has not been shown to test anything. -->

```
```

## Checklist

- [ ] `dotnet build --configuration Release` is clean. Warnings are errors here, so this is
      pass or fail, not a judgement call.
- [ ] `dotnet test Aspose.Slides.Foss.sln --configuration Release` passes — all three suites.
- [ ] **If this changes what is written to the file**, there is a test in
      `tests/Aspose.Slides.Foss.ConformanceTests` that opens the produced `.pptx` and asserts on
      its XML. A test that reads the value back through this library does not count.
- [ ] The test is named after the user-visible failure, not after the method it calls.
- [ ] New public members carry XML documentation. (The build enforces this; the check is here so
      the comment gets written for a reader rather than for the compiler.)
- [ ] `CHANGELOG.md` is updated under `## [Unreleased]` if a caller can observe this change.
- [ ] No new `PackageReference` in `src/Aspose.Slides.Foss`. The library has no dependencies and
      that is a property it advertises.
- [ ] No generated files, build output or test artefacts are included in the diff.

## Anything a reviewer should look at closely

<!-- A decision you were unsure about, a case you did not cover, a behaviour you changed on
     purpose. Say it here rather than letting it be found. -->
