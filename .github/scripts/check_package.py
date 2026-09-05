#!/usr/bin/env python3
"""Assert on the produced .nupkg, and lint the README it carries.

Run it against a packed artefact:

    python .github/scripts/check_package.py artifacts/package/Aspose.Slides.FOSS.26.9.0.nupkg

It exits non-zero and prints every failure. It takes no arguments beyond the package path and reads
nothing from the build that produced it -- the point is to examine the file that would be uploaded,
not to ask the build whether it did the right thing.

WHY THIS EXISTS. The workflow used to list the package contents with `unzip -l` and print them into
the log for a human to compare against the previous run. That is a useful thing to have and it is
not a check: a package missing its icon, missing its readme, or carrying only one of its two target
frameworks would have gone out green. Every assertion here is one that would have been silent.

Two of them are worth naming, because they are the failures that do not look like failures:

  * `lib/<tfm>/` for every framework in the manifest. A `TargetFramework` (singular) inherited from
    Directory.Build.props silently overrides `TargetFrameworks` (plural) in the project file. The
    build succeeds, the tests pass, and the package ships one framework.

  * The README is linted for what nuget.org silently drops rather than rejects: an image from a host
    outside its allow-list, a relative link, a Mermaid fence. None of those is an error at push time.
    The package page just renders wrong, and the warning is visible only to the package owners.
"""
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

# nuget.org renders an image only if its host is on this list. Everything else is dropped with a
# warning that only the package owner ever sees. Taken from the nuget.org documentation, "Package
# readme on nuget.org". `products.aspose.org` is deliberately NOT here -- it is not on the list, and
# a banner served from it does not render. That is why the banner is mirrored into this repository
# and served from raw.githubusercontent.com.
IMAGE_HOSTS = {
    "api.codacy.com", "api.codeclimate.com", "api.dependabot.com", "api.reuse.software",
    "api.travis-ci.com", "app.codacy.com", "app.deepsource.com", "avatars.githubusercontent.com",
    "badgen.net", "badges.gitter.im", "camo.githubusercontent.com", "caniuse.bitsofco.de",
    "cdn.jsdelivr.net", "cdn.syncfusion.com", "ci.appveyor.com", "circleci.com", "cloudback.it",
    "codecov.io", "codefactor.io", "coveralls.io", "dev.azure.com", "devpod.sh", "flat.badgen.net",
    "github.com", "gitlab.com", "i.imgur.com", "img.shields.io", "infragistics.com",
    "isitmaintained.com", "media.githubusercontent.com", "opencollective.com", "raw.github.com",
    "raw.githubusercontent.com", "snyk.io", "sonarcloud.io", "travis-ci.com", "travis-ci.org",
    "user-images.githubusercontent.com",
}

NS = "{http://schemas.microsoft.com/packaging/2012/06/nuspec.xsd}"


def lint_readme(text, package_id, problems):
    """Everything nuget.org drops without telling anybody but the owner."""
    if "```mermaid" in text:
        problems.append("readme: a ```mermaid fence. nuget.org parses CommonMark through Markdig, "
                        "which has no Mermaid, so the diagram renders as a wall of code.")

    for alt, target in re.findall(r"!\[([^\]]*)\]\(([^)\s]+)", text):
        if not target.startswith("http"):
            problems.append("readme: image %r has a relative path. nuget.org does not resolve "
                            "relative images." % target)
            continue
        host = target.split("/")[2].lower()
        if host not in IMAGE_HOSTS:
            problems.append("readme: image host %r is not on the nuget.org allow-list, so %r will "
                            "not render." % (host, target))

    for label, target in re.findall(r"(?<!!)\[([^\]]*)\]\(([^)\s]+)", text):
        if not target.startswith("http"):
            problems.append("readme: link %r -> %r is relative. The package is not a repository; "
                            "relative links resolve to nothing on the package page."
                            % (label, target))

    installs = re.findall(r"dotnet add package\s+(\S+)", text)
    if not installs:
        problems.append("readme: no `dotnet add package` line. The page has to say how to install "
                        "the thing it is the page for.")
    for got in installs:
        if got != package_id:
            problems.append("readme: install command says %r, the package id is %r. One of the two "
                            "is wrong and the reader will copy the wrong one." % (got, package_id))


def main(argv):
    if len(argv) != 2:
        sys.stderr.write("usage: check_package.py <path to .nupkg>\n")
        return 2
    path = argv[1]
    problems = []

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        specs = [n for n in names if n.endswith(".nuspec")]
        if len(specs) != 1:
            sys.stderr.write("expected exactly one .nuspec, found %d\n" % len(specs))
            return 2
        meta = ET.fromstring(z.read(specs[0])).find(NS + "metadata")

        def field(tag):
            el = meta.find(NS + tag)
            return None if el is None else (el.text or "")

        package_id = field("id")

        # The .nuspec is generated, so these assert that the project file said what it meant to.
        if not package_id:
            problems.append("nuspec: no <id>")
        if specs[0] != package_id + ".nuspec":
            problems.append("nuspec: file is %r but <id> is %r; the casing must match byte for byte, "
                            "because it is the casing nuget.org displays forever." % (specs[0], package_id))
        for tag in ("version", "title", "authors", "description", "projectUrl", "icon", "readme",
                    "copyright", "tags", "releaseNotes"):
            if not field(tag):
                problems.append("nuspec: <%s> is missing or empty" % tag)

        lic = meta.find(NS + "license")
        if lic is None or lic.get("type") != "expression" or (lic.text or "") != "MIT":
            problems.append("nuspec: <license> must be the MIT expression")
        if (field("requireLicenseAcceptance") or "false").lower() != "false":
            problems.append("nuspec: requireLicenseAcceptance must be false for an MIT package")

        # Every framework the manifest claims must actually carry an assembly. This is the
        # assertion that catches a silently single-targeted package.
        deps = meta.find(NS + "dependencies")
        groups = [] if deps is None else [g.get("targetFramework") for g in deps.findall(NS + "group")]
        if not groups:
            problems.append("nuspec: no dependency groups, so no target framework is declared")
        for tfm in groups:
            if not any(n.startswith("lib/%s/" % tfm) for n in names):
                problems.append("package: <dependencies> declares %r but there is no lib/%s/ "
                                "in the package" % (tfm, tfm))
        for name in names:
            if name.startswith("lib/"):
                tfm = name.split("/")[1]
                if tfm not in groups:
                    problems.append("package: lib/%s/ is present but %r is not declared in "
                                    "<dependencies>" % (tfm, tfm))

        for required in (field("icon"), field("readme"), "LICENSE"):
            if required and required not in names:
                problems.append("package: %r is named in the manifest but is not in the package"
                                % required)

        if sum(1 for n in names if n == "README.md") > 1:
            problems.append("package: more than one README.md at the package root")

        for tfm in groups:
            dll = [n for n in names if n.startswith("lib/%s/" % tfm) and n.endswith(".dll")]
            xml = [n for n in names if n.startswith("lib/%s/" % tfm) and n.endswith(".xml")]
            if not dll:
                problems.append("package: lib/%s/ carries no assembly" % tfm)
            if not xml:
                problems.append("package: lib/%s/ carries no XML documentation" % tfm)

        readme_name = field("readme")
        if readme_name and readme_name in names:
            lint_readme(z.read(readme_name).decode("utf-8"), package_id, problems)

    print(json.dumps({
        "package": path,
        "id": package_id,
        "entries": len(names),
        "frameworks": sorted(groups),
    }, indent=2))

    for problem in problems:
        print("FAIL " + problem)
    print("%d assertion(s) failed" % len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
