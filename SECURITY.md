# Security policy

## Supported versions

No version of this library has been published to nuget.org yet. Until a release exists, the only
code that receives security fixes is the current `main` branch of this repository.

| Version | Supported |
|---|---|
| `main` | yes |
| Anything built from an older commit | no — rebuild from `main` |

When the first release is published this table will list the released versions instead, and this
sentence will be replaced by a support window.

## Reporting a vulnerability

**Do not open a public issue for a security problem, and do not attach a proof-of-concept file to
one.**

Use GitHub's private vulnerability reporting on this repository:
[**Report a vulnerability**](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/security/advisories/new).
It opens a private advisory that only the maintainers can read, and it lets you attach files and
discuss a fix before anything becomes public.

If that page is not available to you, open a public issue containing **only** the sentence "I would
like to report a security issue privately" and no details at all, and wait to be contacted.

Please include, as far as you can:

- the commit or version you tested,
- the .NET runtime and operating system,
- a minimal `.pptx` or code sample that triggers it,
- what happens, and what you expected instead,
- the impact you believe it has.

You will get an acknowledgement of the report. We cannot promise a fix deadline for a project with no
paid support contract behind it, but you will be told what is happening and when a fix lands, and you
will be credited in the advisory unless you ask not to be.

## What is in scope

This library parses untrusted input: a `.pptx` is a ZIP archive full of XML, and both layers are
attacker-controlled when the file came from outside. Reports about the handling of a malicious or
malformed presentation are in scope, including:

- a crafted archive or XML that causes a crash, an unbounded allocation or an infinite loop,
- a part name or relationship target that escapes the package and reaches the file system,
- XML processing that reaches the network or the local disk,
- anything that lets a presentation influence the process beyond the object model it is parsed into.

## What is out of scope

- Missing capabilities. `Save` raising `NotSupportedException` for a format it does not write, and
  the absent features listed in the README, are documented behaviour, not vulnerabilities.
- Vulnerabilities in the .NET runtime itself — report those to
  [Microsoft](https://msrc.microsoft.com/).
- Vulnerabilities in the commercial Aspose.Slides product, which is different software. Report those
  through [Aspose support](https://forum.aspose.com/c/slides/11).
- Findings from an automated scanner with no demonstrated impact on this library.
