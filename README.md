# Clean Office and image metadata on Linux Mint

Keep both scripts together. The wrapper recursively processes DOCX, PPTX, JPEG, PNG, WebP and TIFF locally. By default it writes verified copies to `./cleaned_ai_metadata/`, preserving filenames and subdirectories.

## Install

```bash
sudo apt update
sudo apt install python3 python3-lxml python3-pil libimage-exiftool-perl unzip zip
chmod +x clean_ai_metadata.sh
exiftool -ver
```

Use Python 3.9+ and ExifTool 12.75+. The helper uses Python for ZIP operations and lxml to preserve Office namespace prefixes. If your Mint repository supplies an older ExifTool, install a current release and check the version on your PATH. [ExifTool installation](https://exiftool.org/install.html)

## Run

```bash
./clean_ai_metadata.sh "./My documents" --dry-run
./clean_ai_metadata.sh "./My documents"
./clean_ai_metadata.sh "./draft slides.pptx" --output "./review copies"
./clean_ai_metadata.sh "./photo.tiff" --in-place
```

Dry-run lists planned paths. Existing outputs cause a per-file error. Choose a fresh output directory for a new run. A nested output directory is excluded from the scan. Directory scans skip symlinks. Use a local filesystem with hard-link support and close files being edited.

In-place mode verifies a staged copy, saves the original bytes beside the source as `filename.UTC-timestamp.bak`, then atomically replaces the source. To restore it, copy the printed backup over the cleaned file. Outputs and backups are private files owned by the invoking user. In-place replacement creates a new inode and can change permissions, ACLs, and filesystem timestamps.

**Optional MagicDispel pass**

```bash
sudo apt install pipx
pipx install magicdispel
export PATH="$HOME/.local/bin:$PATH"
./clean_ai_metadata.sh "./photos" --output "./rebuilt" --magicdispel
```

MagicDispel rebuilds a same-format copy using the image data and checks the result. ExifTool then runs both deletion passes. Neither mode requests lossy re-encoding. [MagicDispel usage and format details](https://pypi.org/project/magicdispel/)

## Verification

Each success line lists cleared fields, Unicode removals, images cleaned, and the C2PA/JUMBF result. For Office files the helper reopens the completed ZIP, tests every CRC, compares the part inventory and content types, checks the main document relationship, checks the property fields, and reads each target image back from the ZIP. Pillow compares decoded pixels and transparency for every frame before and after cleaning. For JPEGs with a JFIF header the cleaner rebuilds a minimal header with its original density, omitting thumbnails. Some Office viewers display a black rectangle when this header disappears. Image checks query both the raw JUMBF block and its decoded tags. An image-free package reports C2PA verification as inapplicable. A failure gives a nonzero exit status. Other verified files in the batch are retained.

To spot-check a result, run the following from the directory containing the cleaned files. The two JUMBF queries should produce no output. Use the actual member name from `unzip -Z1` for the embedded image.

```bash
exiftool -config '' -a -u -G1 -s -JUMBF -JUMBF:all "./photo.tiff"
unzip -t "./draft slides.pptx"
unzip -Z1 "./draft slides.pptx"
unzip -p "./draft slides.pptx" 'ppt/media/image1.png' |
  exiftool -config '' -a -u -G1 -s -JUMBF -JUMBF:all -
unzip -p "./draft slides.pptx" docProps/core.xml
```

## Scope and limitations

- The helper clears the listed core properties, identifying application fields, and all custom properties. It retains the custom-property part and its relationships. LibreOffice may write fresh author or application metadata on a later save.
- The deletion set is U+00AD, 034F, 061C, 115F, 1160, 17B4–17B5, 180B–180F, 200B–200F, 202A–202E, 2060–2064, 2066–2069, 3164, FE00–FE0F, FEFF, FFA0, and E0100–E01EF. Joiners, bidi controls, fillers and variation selectors can affect spelling, reading order, emoji or glyph appearance. Review documents that use them. The helper edits XML character data under `word/` and `ppt/` and in the property parts. Verification scans decoded package XML, including numeric references. A control remaining in an attribute, relationship or comment fails verification with its part name and count.
- Supported images in Office media directories and standard thumbnails receive the same cleaning as standalone images. An unsupported image such as SVG or EMF fails the package. Audio, video, embedded documents, visible watermarks, tracked changes, comment authors and external resources are outside the cleaning scope.
- ExifTool removes writable metadata. TIFF needs structural EXIF fields, and proprietary data can survive. Removing profiles, orientation and display metadata can change how unchanged pixel data appears. The C2PA check detects manifests ExifTool recognizes. [ExifTool deletion coverage](https://exiftool.org/) · [JUMBF support](https://exiftool.org/TagNames/Jpeg2000.html)
- Signed or encrypted Office packages are refused. ZIP limits are 20,000 members, 128 MiB per member, 32 MiB per XML part, and 1 GiB total expanded. Package checks cover ZIP and XML integrity. Open a cleaned copy in LibreOffice to check its layout.
