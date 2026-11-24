#!/usr/bin/env python3
"""Clean and verify local image and Office files. See README.md for scope."""
import argparse
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timezone
import fcntl
import io
import json
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import zipfile
import zlib

if sys.version_info < (3, 9):
    sys.exit("Python 3.9 or newer is required")

try:
    from lxml import etree
except ImportError:
    sys.exit("Missing lxml for python3. Install with sudo apt install python3-lxml")


class Failure(Exception):
    pass


IMAGE = {".jpg": "JPEG", ".jpeg": "JPEG", ".png": "PNG", ".webp": "WEBP",
         ".tif": "TIFF", ".tiff": "TIFF"}
MIME_EXT = {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp",
            "image/tiff": ".tif"}
OFFICE = {".docx": "word/document.xml", ".pptx": "ppt/presentation.xml"}
CORE = set("creator lastModifiedBy title subject description keywords category "
           "contentStatus revision created modified lastPrinted".split())
APP = set("Application Company Manager AppVersion Template TotalTime HyperlinkBase "
          "HLinks TitlesOfParts HeadingPairs PresentationFormat".split())
PROPS = {"docProps/core.xml": CORE, "docProps/app.xml": APP, "docProps/custom.xml": None}
# Unicode format controls, fillers and variation selectors. No normalization.
POINTS = {0xFEFF, 0x00AD, 0x034F, 0x061C, 0x180E, 0x115F, 0x1160,
          0x17B4, 0x17B5, 0x3164, 0xFFA0}
for first, last in [(0x200B, 0x200F), (0x202A, 0x202E), (0x2060, 0x2064),
                    (0x2066, 0x2069), (0x180B, 0x180D), (0x180F, 0x180F),
                    (0xFE00, 0xFE0F), (0xE0100, 0xE01EF)]:
    POINTS.update(range(first, last + 1))
DELETE = dict.fromkeys(POINTS)
INVISIBLE = re.compile("[" + "".join(map(chr, sorted(POINTS))) + "]")
TIFF_PRIVATE = set("Artist Copyright Software ImageDescription Make Model HostComputer "
                   "DocumentName ModifyDate PageName CameraOwnerName CameraSerialNumber "
                   "XPTitle XPComment XPAuthor XPKeywords XPSubject".split())
MIB = 1024 * 1024


@dataclass
class Report:
    fields: Counter = field(default_factory=Counter)
    chars: int = 0
    images: int = 0


def command(args, tiff_all=False):
    result = subprocess.run(args, capture_output=True, encoding="utf-8", errors="replace",
                            env={**os.environ, "LC_ALL": "C"}, timeout=300)
    messages = result.stderr.splitlines()
    if tiff_all:
        # TIFF requires its primary image directory to decode the pixels.
        messages = [s for s in messages if not s.startswith(
            "Warning: [minor] Can't delete IFD0 from TIFF - ")]
    if result.returncode or messages:
        detail = "\n".join(messages) or result.stdout or f"exit {result.returncode}"
        raise Failure(f"{args[0]}: {detail.strip()[:1200]}")
    return result.stdout


def exif(*args, tiff_all=False):
    # Disable user-supplied ExifTool config. File operands are absolute private paths.
    return command(["exiftool", "-config", "", *args], tiff_all)


def dependencies(magic):
    if not shutil.which("exiftool"):
        raise Failure("Missing exiftool. Install with sudo apt install libimage-exiftool-perl")
    version = exif("-ver").strip()
    match = re.fullmatch(r"(\d+)\.(\d+)(?:\.\d+)?", version)
    if not match or tuple(map(int, match.groups())) < (12, 75):
        raise Failure(f"ExifTool {version}: install 12.75+ for raw JUMBF verification")
    if magic and not shutil.which("magicdispel"):
        raise Failure("--magicdispel requires magicdispel on PATH. See README.md")


def tags(path):
    row, = json.loads(exif("-j", "-G0:1", "-a", "-u", "-n", "-all", "-JUMBF",
                           "-JUMBF:all", "-Warning", "-Error", str(path)))
    for key in row:
        if key.rsplit(":", 1)[-1] in {"Warning", "Error"}:
            raise Failure(f"ExifTool diagnostic: {row[key]}")
    actual = row.get("File:FileType", "").upper()
    actual = {"EXTENDED WEBP": "WEBP", "APNG": "PNG"}.get(actual, actual)
    if actual != IMAGE[path.suffix]:
        raise Failure(f"image content does not match expected {IMAGE[path.suffix]} type")
    return row


def private_tiff_tags(row):
    # Keep the concrete directory (IFD0, IFD1, etc.) for multipage TIFF.
    return {k.removeprefix("EXIF:") for k in row if k.startswith("EXIF:")
            and k.rsplit(":", 1)[-1] in TIFF_PRIVATE}


def check_image(path):
    row = tags(path)
    remaining = [k for k in row if "jumbf" in k.lower() or "c2pa" in k.lower()
                 or k.startswith(("XMP:", "IPTC:", "EXIF:GPS:"))]
    if IMAGE[path.suffix] == "TIFF":
        remaining.extend(private_tiff_tags(row))
    elif any(k.startswith("EXIF:") for k in row):
        remaining.append("EXIF")
    if remaining:
        raise Failure("remaining image tags: " + ", ".join(sorted(remaining)))
    return row


def clean_image(path, magic, report):
    before = tags(path)
    if magic:
        command(["magicdispel", str(path)])
        copies = list(path.parent.glob(path.stem + "_clean.*"))
        if len(copies) != 1 or IMAGE.get(copies[0].suffix.lower()) != IMAGE[path.suffix]:
            raise Failure("MagicDispel did not produce one same-format clean copy")
        os.replace(copies[0], path)
    is_tiff = IMAGE[path.suffix] == "TIFF"
    exif("-all=", "-overwrite_original", str(path), tiff_all=is_tiff)
    if is_tiff:
        extra = sorted(private_tiff_tags(before))
        if extra:
            exif(*(f"-{key}=" for key in extra), "-overwrite_original", str(path))
    exif("-jumbf:all=", "-overwrite_original", str(path))
    after = check_image(path)
    if not before.get("Composite:ImageSize") or before["Composite:ImageSize"] != after.get("Composite:ImageSize"):
        raise Failure("missing or changed image dimensions")
    if is_tiff:
        page_sizes = {k: v for k, v in before.items()
                      if re.fullmatch(r"EXIF:IFD[0-9]+:Image(?:Width|Height)", k)}
        if any(after.get(k) != v for k, v in page_sizes.items()):
            raise Failure("TIFF page count or dimensions changed")
    for key in before.keys() - after.keys():
        if ":" in key and not key.startswith(("File:System:", "ExifTool:", "Composite:")):
            report.fields[key] += 1
    report.images += 1


def parse_xml(blob):
    if len(blob) > 32 * MIB:
        raise Failure("XML part exceeds 32 MiB")
    parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True,
                             remove_blank_text=False, strip_cdata=False, collect_ids=False)
    tree = etree.parse(io.BytesIO(blob), parser)
    if tree.docinfo.doctype:
        raise Failure("XML with a DTD is unsupported")
    return tree


def local(node):
    return etree.QName(node).localname


def property_nodes(root, name):
    fields = PROPS[name]
    for child in root:
        if isinstance(child.tag, str) and (fields is None or local(child) in fields):
            yield child


def clean_xml(blob, name, report):
    if name not in PROPS and not name.startswith(("word/", "ppt/")):
        return blob
    tree = parse_xml(blob)
    root = tree.getroot()
    changed = False
    if name in PROPS:
        for child in list(property_nodes(root, name)):
            key = "custom.property" if PROPS[name] is None else f"{Path(name).stem}.{local(child)}"
            report.fields[key] += 1
            text = etree.tostring(child, encoding="unicode", with_tail=False)
            report.chars += len(INVISIBLE.findall(text))
            # Preserve the text following a removed element.
            prev = child.getprevious()
            if child.tail:
                if prev is None:
                    root.text = (root.text or "") + child.tail
                else:
                    prev.tail = (prev.tail or "") + child.tail
            root.remove(child)
            changed = True
    for node in root.iter():
        # Comments and processing instructions have a tail but are not text nodes.
        for attr in (("text", "tail") if isinstance(node.tag, str) else ("tail",)):
            value = getattr(node, attr)
            if value and INVISIBLE.search(value):
                cleaned = value.translate(DELETE)
                report.chars += len(value) - len(cleaned)
                setattr(node, attr, cleaned)
                changed = True
    if not changed:
        return blob
    return etree.tostring(tree, encoding="UTF-8", xml_declaration=True,
                          standalone=tree.docinfo.standalone)


def index_package(z, ext):
    entries = z.infolist()
    names = set(z.namelist())
    if len(entries) != len(names) or len(entries) > 20000:
        raise Failure("duplicate ZIP names or over 20,000 ZIP members")
    if sum(i.file_size for i in entries) > 1024 * MIB:
        raise Failure("expanded package exceeds 1 GiB")
    for i in entries:
        name = i.filename
        if (name != i.orig_filename or "\\" in name or i.flag_bits & 1 or
                any(part in ("", ".", "..") for part in name.rstrip("/").split("/")) or
                i.file_size > 128 * MIB):
            raise Failure(f"unsafe, encrypted or oversized ZIP member {name!r}")
        if i.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            raise Failure(f"unsupported Office ZIP compression in {name!r}")
        if name.lower().startswith("_xmlsignatures/"):
            raise Failure("signed Office package would lose its signature")
    main = OFFICE[ext]
    missing = {"[Content_Types].xml", "_rels/.rels", main} - names
    if missing:
        raise Failure(f"missing required parts: {sorted(missing)}")
    manifest = parse_xml(z.read("[Content_Types].xml")).getroot()
    defaults, overrides = {}, {}
    for child in manifest:
        if isinstance(child.tag, str):
            kind = child.get("ContentType", "").lower()
            if "digital-signature" in kind:
                raise Failure("signed Office package would lose its signature")
            if local(child) == "Default":
                defaults[child.get("Extension", "").lower()] = kind
            elif local(child) == "Override":
                overrides[child.get("PartName", "").lstrip("/")] = kind
    family, document = (("wordprocessingml", "document") if ext == ".docx"
                        else ("presentationml", "presentation"))
    main_type = f"application/vnd.openxmlformats-officedocument.{family}.{document}.main+xml"
    if overrides.get(main, defaults.get("xml")) != main_type:
        raise Failure("main part has the wrong Office content type")
    main_root = parse_xml(z.read(main)).getroot()
    namespaces = {f"http://schemas.openxmlformats.org/{family}/2006/main",
                  f"http://purl.oclc.org/ooxml/{family}/main"}
    if local(main_root) != document or etree.QName(main_root).namespace not in namespaces:
        raise Failure("main part has the wrong Office root element")
    rels = parse_xml(z.read("_rels/.rels")).getroot()
    if not any(r.get("Type", "").endswith("/officeDocument") and
               r.get("Target", "").lstrip("/") == main and r.get("TargetMode") != "External"
               for r in rels):
        raise Failure("package relationship does not point to the expected main part")
    return {n: overrides.get(n, defaults.get(Path(n).suffix[1:].lower(), ""))
            for n in names}


def image_ext(name, mime):
    suffix = Path(name).suffix.lower()
    if name.startswith(("word/media/", "ppt/media/")) or name.startswith("docProps/thumbnail."):
        if suffix in IMAGE:
            return suffix
        if mime in MIME_EXT:
            return MIME_EXT[mime]
        if mime.startswith(("audio/", "video/")):
            return None
        raise Failure(f"unsupported embedded media {name!r} ({mime or 'unknown type'})")
    return None


def xml_part(name, mime):
    return name.lower().endswith((".xml", ".rels", ".vml")) or mime.endswith(("+xml", "/xml"))


def verify_office(path, expected, ext, scratch):
    problems = []
    with zipfile.ZipFile(path) as z:
        types = index_package(z, ext)
        if types != expected:
            raise Failure("ZIP part names or content types changed")
        bad = z.testzip()
        if bad:
            raise Failure(f"ZIP CRC failure in {bad!r}")
        for i in z.infolist():
            if i.is_dir():
                continue
            name = i.filename
            mime = types[name]
            suffix = image_ext(name, mime)
            if suffix:
                image = scratch / ("probe" + suffix)
                image.write_bytes(z.read(i))
                try:
                    check_image(image)
                except Failure as exc:
                    raise Failure(f"{name!r}: {exc}") from exc
                image.unlink()
            elif xml_part(name, mime):
                tree = parse_xml(z.read(i))
                if name in PROPS:
                    left = [local(n) for n in property_nodes(tree.getroot(), name)]
                    if left:
                        problems.append(f"{name!r}: remaining fields {left}")
                # Equivalent to a Unicode grep after XML decoding. Includes attributes,
                # comments, namespace declarations and numeric character references.
                content = etree.tostring(tree, encoding="unicode")
                counts = Counter(map(ord, INVISIBLE.findall(content)))
                if counts:
                    summary = ", ".join(f"U+{cp:04X} × {n}" for cp, n in sorted(counts.items()))
                    problems.append(f"{name!r}: {summary}")
    if problems:
        raise Failure("XML verification failed: " + " | ".join(problems))


def clean_office(source, dest, ext, scratch, magic, report):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(dest, "w") as zout:
        types = index_package(zin, ext)
        for i in zin.infolist():
            name = i.filename
            data = zin.read(i)
            if not i.is_dir():
                mime = types[name]
                suffix = image_ext(name, mime)
                if suffix:
                    image = scratch / ("picture" + suffix)
                    image.write_bytes(data)
                    try:
                        clean_image(image, magic, report)
                    except Failure as exc:
                        raise Failure(f"{name!r}: {exc}") from exc
                    data = image.read_bytes()
                    image.unlink()
                elif xml_part(name, mime) and not name.lower().endswith(".rels"):
                    data = clean_xml(data, name, report)
            # Fresh entry headers discard original ZIP comments, extra fields and dates.
            entry = zipfile.ZipInfo(name)
            entry.compress_type = i.compress_type
            zout.writestr(entry, data)
    verify_office(dest, types, ext, scratch)


def check_parents(path):
    for parent in (path, *path.parents):
        if parent.is_symlink():
            raise Failure(f"symlink directory in destination: {parent!s}")


def files(root, excluded):
    if not root.is_dir():
        if root.suffix.lower() not in IMAGE and root.suffix.lower() not in OFFICE:
            raise Failure("unsupported input extension")
        return [root]

    def walk_error(error):
        raise error

    result = []
    for parent, dirs, names in os.walk(root, onerror=walk_error):
        base = Path(parent)
        dirs[:] = sorted(d for d in dirs if base / d != excluded
                         and not re.fullmatch(r"\.metadata-[a-z0-9_]{8}\.tmp", d)
                         and not (base / d).is_symlink())
        for name in sorted(names):
            p = base / name
            if p.suffix.lower() in IMAGE or p.suffix.lower() in OFFICE:
                if not p.is_symlink() and p.is_file():
                    result.append(p)
    return result


def identity(st):
    return st.st_dev, st.st_ino, st.st_size, st.st_mtime_ns, st.st_ctime_ns


def unchanged(source, original):
    if identity(source.lstat()) != original:
        raise Failure("source changed during processing")


def sync_file(path):
    with path.open("rb") as f:
        os.fsync(f.fileno())


def sync_dir(path):
    fd = os.open(path, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(fd)
    finally:
        os.close(fd)


def backup(stream, source, scratch):
    # Publish a complete backup with no-clobber semantics before replacement.
    stream.seek(0)
    copy = scratch / "backup"
    with copy.open("xb") as f:
        shutil.copyfileobj(stream, f)
        f.flush()
        os.fsync(f.fileno())
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    dest = source.with_name(f"{source.name}.{stamp}.bak")
    os.link(copy, dest)
    sync_dir(dest.parent)
    return dest


def process(source, dest, args):
    ext = source.suffix.lower()
    report = Report()
    check_parents(dest.parent)
    if not args.in_place and os.path.lexists(dest):
        raise Failure(f"output exists: {dest!s}")
    if args.dry_run:
        print(f"PLAN {ext[1:].upper()} {str(source)!r} -> {str(dest)!r}"
              + (" (backup first)" if args.in_place else ""))
        return
    dest.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    fd = os.open(source, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK)
    with os.fdopen(fd, "rb") as stream, tempfile.TemporaryDirectory(
            prefix=".metadata-", suffix=".tmp", dir=dest.parent) as tmp:
        # Coordinate concurrent invocations on the same inode.
        fcntl.flock(stream, (fcntl.LOCK_EX if args.in_place else fcntl.LOCK_SH) | fcntl.LOCK_NB)
        st = os.fstat(stream.fileno())
        if not stat.S_ISREG(st.st_mode):
            raise Failure("input must be a regular file")
        original = identity(st)
        scratch = Path(tmp)
        staged = scratch / ("output" + ext)
        if ext in OFFICE:
            clean_office(stream, staged, ext, scratch, args.magicdispel, report)
        else:
            with staged.open("xb") as f:
                shutil.copyfileobj(stream, f)
            clean_image(staged, args.magicdispel, report)
        sync_file(staged)
        unchanged(source, original)
        if args.in_place:
            saved = backup(stream, source, scratch)
            print(f"BACKUP {str(saved)!r}", flush=True)
            unchanged(source, original)
            os.replace(staged, dest)
        else:
            os.link(staged, dest)  # EEXIST even if another process won the race.
        sync_dir(dest.parent)
    fields = ", ".join(f"{k} ×{v}" if v > 1 else k for k, v in sorted(report.fields.items()))
    verified = "pass" if report.images else "n/a (no target images)"
    print(f"OK {ext[1:].upper()} {str(source)!r} -> {str(dest)!r} | "
          f"fields cleared: {fields or 'none'} | invisible removed: {report.chars} | "
          f"images cleaned: {report.images} | C2PA/JUMBF: {verified}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="file or directory to scan recursively")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--output", type=Path, help="output directory (default ./cleaned_ai_metadata)")
    group.add_argument("--in-place", action="store_true", help="replace verified files after backup")
    parser.add_argument("--dry-run", action="store_true", help="list planned paths without writing")
    parser.add_argument("--magicdispel", action="store_true", help="also rebuild supported images")
    args = parser.parse_args()
    os.umask(0o077)
    dependencies(args.magicdispel)
    if args.input.is_symlink():
        raise Failure("input is a symlink")
    root = args.input.resolve(strict=True)
    if not (root.is_file() or root.is_dir()):
        raise Failure("input must be a regular file or directory")
    out = None
    if not args.in_place:
        out = (args.output or Path("cleaned_ai_metadata")).absolute()
        check_parents(out)
        out = out.resolve()
        if root == out or (root.is_dir() and out in root.parents):
            raise Failure("output must be distinct from the input and its ancestor directories")
        if out.exists() and not out.is_dir():
            raise Failure("output must be a directory")
    selected = files(root, out)
    if not selected:
        raise Failure("no supported files found")
    failed = False
    for source in selected:
        dest = source if args.in_place else out / (source.relative_to(root) if root.is_dir()
                                                   else source.name)
        try:
            process(source, dest, args)
        except (Failure, OSError, ValueError, RuntimeError, zipfile.BadZipFile,
                etree.LxmlError, zlib.error, subprocess.SubprocessError) as exc:
            print(f"FAIL {source.suffix[1:].upper()} {str(source)!r}: {exc}", file=sys.stderr)
            failed = True
    return int(failed)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (Failure, OSError, ValueError, subprocess.SubprocessError) as exc:
        sys.exit(str(exc))
    except KeyboardInterrupt:
        sys.exit(130)
