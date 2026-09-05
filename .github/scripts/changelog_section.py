#!/usr/bin/env python3
"""Print one version's section of CHANGELOG.md, for use as release notes.

    python .github/scripts/changelog_section.py 26.9.0

Prints the body of the `## [26.9.0]` section — everything after its heading, up to but not including
the next `## ` heading — with the changelog's link-reference definitions appended so that any
`[Unreleased]`-style reference inside the body still resolves. Exits non-zero, with a message naming
what it looked for, if the section is not there.

WHY THIS EXISTS RATHER THAN `gh release create --generate-notes`. Generated notes are built from the
merge history: pull-request titles, contributor handles, a "New Contributors" block. On this
repository that produces a release page that says nothing about the release — the user-facing
content is entirely in this changelog — and that quotes internal wording out of pull-request bodies
onto a public page. The changelog is written for the reader of the release; the merge history is not.

Failing loudly matters more than it looks. A release whose notes came out empty is not something
anyone notices until someone reads the page, by which time the tag exists and the package is
published. So an empty or missing section is an error, not a blank file.
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CHANGELOG = os.path.join(os.path.dirname(os.path.dirname(HERE)), "CHANGELOG.md")

# `## [26.9.0] - 2026-09-26`, `## [26.9.0]`, `## 26.9.0` — the date part is optional and free-form,
# because a release date is written by a person on the day and this should not be the thing that
# rejects it.
def heading_for(version):
    return re.compile(r"^##\s+\[?" + re.escape(version) + r"\]?\s*(?:[-–—]\s*\S.*)?$")


ANY_H2 = re.compile(r"^##\s+")
# A markdown link-reference definition: `[26.9.0]: https://...`
LINK_DEF = re.compile(r"^\[[^\]]+\]:\s+\S+")


def extract(text, version):
    lines = text.splitlines()
    wanted = heading_for(version)

    start = None
    for i, line in enumerate(lines):
        if wanted.match(line):
            start = i + 1
            break
    if start is None:
        return None, None

    end = len(lines)
    for i in range(start, len(lines)):
        if ANY_H2.match(lines[i]):
            end = i
            break

    # Link-reference definitions live at the foot of the file rather than inside any section, but the
    # LAST section runs to end of file and so swallows them. They are stripped out of the body here
    # and re-attached below only if something in the body actually needs them — otherwise the final
    # release of a changelog gets the block twice, once as content and once as a footer.
    body = [line for line in lines[start:end] if not LINK_DEF.match(line)]
    refs = [line for line in lines if LINK_DEF.match(line)]
    return body, refs


def main(argv):
    if len(argv) != 2:
        sys.stderr.write("usage: changelog_section.py <version>\n")
        return 2
    version = argv[1]

    if not os.path.isfile(CHANGELOG):
        sys.stderr.write("no CHANGELOG.md at %s\n" % CHANGELOG)
        return 2
    with io.open(CHANGELOG, encoding="utf-8") as fh:
        text = fh.read()

    body, refs = extract(text, version)
    if body is None:
        sys.stderr.write(
            "CHANGELOG.md has no section for %s.\n"
            "Looked for a line matching: ## [%s]  (optionally followed by a date).\n"
            "Cutting a release renames the Unreleased heading to the version being released; that\n"
            "edit belongs in the release commit, before the tag.\n" % (version, version))
        return 1

    while body and not body[0].strip():
        body.pop(0)
    while body and not body[-1].strip():
        body.pop()

    if not body:
        sys.stderr.write("CHANGELOG.md has a section for %s and it is empty.\n" % version)
        return 1

    out = "\n".join(body)
    if refs and re.search(r"(?<!!)\[[^\]]+\](?!\()", out):
        out += "\n\n" + "\n".join(refs)
    sys.stdout.write(out + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
