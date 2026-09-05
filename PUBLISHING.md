# Publishing

How a release of `Aspose.Slides.FOSS` reaches nuget.org, what stops it, and what to do when
something goes wrong halfway.

There is **no API key in this repository**, and there is not meant to be. The publish job proves its
identity to nuget.org with a GitHub OIDC token and receives a key valid for one hour. Nothing
long-lived exists to leak, rotate, or forget to revoke.

## Cutting a release

1. **Make the release commit.** Set `<Version>` in `Directory.Build.props` — it is the only place a
   version is written — and fix everything the release makes untrue. The checklist is below.
2. **Push it to `main` and let CI go green.**
3. **Dry-run the release workflow.** Actions → *Release to NuGet* → *Run workflow*. It runs every
   guard, the whole test suite, the pack, the package assertions and the consumer check, and stops
   short of publishing. It reports the tag a release would need.
4. **Push the tag.** `git tag v26.9.0 && git push origin v26.9.0`. The tag must be `v` followed by
   exactly the version in `Directory.Build.props`; the workflow refuses anything else.
5. **Approve the deployment.** The publish job waits on the `nuget.org` environment until a reviewer
   approves it. Until then nothing has been sent.
6. **Watch it land.** The workflow polls nuget.org until the version is actually downloadable, then
   creates the GitHub release from the changelog section for that version.

### The release-commit checklist

Editing `<Version>` is one step of cutting a release and not the whole of it. These files state, in
the present tense, things that stop being true the moment the package exists:

| File | What changes |
|---|---|
| `Directory.Build.props` | `<Version>`. |
| `CHANGELOG.md` | Delete the "nothing has been published" line. Rename `## [Unreleased]` to `## [<version>] - <the actual publication date>`, and add the matching link-reference at the foot. **The release notes are that section**, so it is the release page's text. |
| `README.md` | The **Installation** section. `dotnet add package Aspose.Slides.FOSS` resolves once a release exists, and the sentence saying it does not is then false. |
| `docs/nuget/README.md` | The package page. Usually needs nothing at release time, but it is the file consumers read first — check it says what this version does. |
| `SECURITY.md` | **Supported versions**: the table and the sentence above it both rest on nothing having been released. |

Do not date the changelog heading in advance. The release commit lands before the tag, and the tag
is pushed by a person; a date written ahead of time asserts a publication that has not happened.

## What the workflow checks, in order

Everything before the push is designed to fail on a branch, in a dry run, or in the guard job —
somewhere cheap — rather than after a version number has been spent.

| # | Guard | Why it exists |
|---|---|---|
| 1 | The id and version come from MSBuild, not from a file name | `-getProperty` on an undefined property prints an empty line and exits 0, so both are explicitly checked for emptiness. An empty version builds `Aspose.Slides.FOSS..nupkg` and fails far from the cause. |
| 2 | The tag equals `v<version>`, and the version is a plain three-part release | A version is published once and can never be replaced. A tag that disagrees with the package is a mislabelled release nobody can correct afterwards. |
| 3 | Nothing key-shaped is committed in the tracked tree | This repository holds no publishing credential by design. The day this fires, something has changed. |
| 4 | nuget.org is asked whether the version already exists | Three outcomes, and the third is the point: an unexpected status **stops** the release. Treating "I could not tell" as "not published" is how a version gets pushed over one that exists. |
| 5 | The full CI suite runs at the released commit, as a reusable workflow | The artefact that is pushed is the one those jobs built and inspected, not a rebuild nothing tested. |
| 6 | The downloaded artefact is named `<id>.<version>.nupkg`, and `check_package.py` passes on it | The package checker is handed a path and knows nothing of the tag; the cross-check belongs to the release. |
| 7 | After the push, nuget.org must actually serve the version | A successful push is not a published package. |

## Verifying by hand

A green tick is not evidence. These are:

```bash
ID=aspose.slides.foss
VER=26.9.0

# It exists and is downloadable.
curl -s "https://api.nuget.org/v3-flatcontainer/$ID/index.json" | grep "$VER"
curl -s -o /dev/null -w '%{http_code}\n' \
  "https://api.nuget.org/v3-flatcontainer/$ID/$VER/$ID.$VER.nupkg"

# The metadata that shipped is the metadata CI asserted. Package metadata is immutable after
# upload, so this is a check, not a chance to fix anything.
curl -s "https://api.nuget.org/v3-flatcontainer/$ID/$VER/$ID.nuspec"

# It is listed, and it is owned by who it should be. Owners and the verified-prefix flag live ONLY
# on the search resource — they are not in the registration index, and a check that looks for them
# there finds nothing and says so quietly.
curl -s "https://azuresearch-usnc.nuget.org/query?q=packageid:Aspose.Slides.FOSS&prerelease=true"

# Being published and being findable are different states. The Java edition has been on Maven
# Central since July and its search index still returns nothing for it. Search has two replicas;
# check both, because one answering does not mean the other has.
curl -s "https://azuresearch-ussc.nuget.org/query?q=Aspose.Slides.FOSS"
```

And on a machine that has never seen the package:

```bash
dotnet new console -o /tmp/check && cd /tmp/check
dotnet add package Aspose.Slides.FOSS --version 26.9.0
```

Finally, open the package page and look at it. The icon, the README, the badges and the banner
either render or they do not, and nuget.org reports a readme it could not render **only to the
package owners** — no HTTP status anywhere says so.

## Re-running a release that failed halfway

**Re-running the same tag is safe, and is the right move.** Guard 4 asks nuget.org whether the
version already exists; if it does, the push is skipped and only the verification and the GitHub
release run. Do not invent a new version number to get around a failed run — that spends a version
to work around a problem that may not be in the package at all.

Where the failure was matters:

- **Before the push** — any guard, CI, the package assertions. Nothing was sent. Fix and re-push the
  tag (`git tag -f` then a force push of the tag, or delete and recreate it).
- **The wait step timed out.** This does **not** mean the push failed. nuget.org validates and
  indexes after accepting a package, and the workflow gives that about 28 minutes. Look at
  `https://www.nuget.org/packages/Aspose.Slides.FOSS/<version>`: validation failures appear there
  and are emailed to the package owners. If it is simply slow, re-run the tag once it has landed and
  the run will skip the push and finish the verification.
- **The GitHub release step failed.** The package is published; only the release page is missing.
  Re-run the tag — the step is idempotent, it leaves an existing release alone.

### What is not yet known, and is not guessed here

Three behaviours this document would like to state are **undocumented by nuget.org**, and no answer
is written here rather than an inferred one:

1. What the flat container returns for a version that has been pushed but is still validating.
2. What a second push of the same version returns during that window. The protocol reference ties
   `409` to "a package with the provided ID and version already exists" and says nothing about the
   validation window.
3. Whether a version that **fails** validation can ever be re-pushed, or whether the version number
   is spent.

The way to answer them is a rehearsal against `https://int.nugettest.org`, the test instance, which
is what it is for. **That rehearsal has not been run.** When it is, its observations belong in this
section, with the date they were observed — the answers are properties of a service that can change,
not facts of this repository. Note that the test instance needs its own trusted-publishing policy and
its own token-service URL and audience; the login action's defaults are nuget.org's, and the test
host's equivalents are not documented either.

## Credentials, and who holds what

| Thing | Where it lives | Who can change it |
|---|---|---|
| The trusted-publishing policy | nuget.org, under the `Aspose` organisation | a member of that organisation |
| Its binding | this repository (by GitHub id, not by name), the workflow file name `nuget-release.yml`, and the environment `nuget.org` | changing any of the three breaks the exchange |
| The `nuget.org` environment | this repository's settings: required reviewer, deployable only from a `v*` tag | a repository admin |
| `NUGET_USER` | an environment secret holding the nuget.org **profile name** — not an email address, which fails the exchange | a repository admin |
| The published package's owners | `Aspose`, and `asposefoss` added after the first publish — once per package id, forever | an administrator on each side |

Renaming this file, or the environment, or moving the workflow, breaks publishing in a way that
reports itself as a token error rather than as a rename. If any of the three has to change, change
the policy on nuget.org first.

The policy is tied to the membership of the person who created it. If they leave the `Aspose`
organisation it becomes inactive until they are added back, and nothing warns anybody — it surfaces
as a failed release.

## Publication history

| Version | Date | Notes |
|---|---|---|
| — | — | Nothing has been published yet. |
