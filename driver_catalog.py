#!/usr/bin/env python3
"""
Windows Driver Catalog Tool
Search, download, and extract .cab driver files from the Microsoft Update Catalog.
"""

import argparse
import datetime
import hashlib
import json
import os
import platform
import re
import shutil
import struct
import subprocess
import sys
import textwrap
import zlib
from urllib.parse import quote, urlencode

import requests
from bs4 import BeautifulSoup

CATALOG_URL = "https://www.catalog.update.microsoft.com"
SEARCH_URL = f"{CATALOG_URL}/Search.aspx"
DOWNLOAD_URL = f"{CATALOG_URL}/DownloadDialog.aspx"

DOWNLOAD_PATTERN = re.compile(
    r"\[(\d*)\]\.url\s*=\s*[\"'](http[s]?://[^'\"]+\.cab)[\"']"
)
PRODUCT_SPLIT = re.compile(r",(?=[^\s])")

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
    }
)


class CatalogEntry:
    """Represents a single entry from the Windows Update Catalog."""

    def __init__(self, row):
        cells = row.find_all("td")
        self.title = cells[1].get_text().strip()
        self.products = [
            p.strip()
            for p in re.split(PRODUCT_SPLIT, cells[2].get_text().strip())
            if p.strip()
        ]
        self.classification = cells[3].get_text().strip()
        try:
            self.last_updated = datetime.datetime.strptime(
                cells[4].get_text().strip(), "%m/%d/%Y"
            )
        except ValueError:
            self.last_updated = None
        self.version = cells[5].get_text().strip()
        size_spans = cells[6].find_all("span")
        self.size_str = size_spans[0].get_text().strip() if size_spans else "Unknown"
        input_el = cells[7].find("input")
        self.update_id = input_el.attrs["id"] if input_el else None

    def get_download_urls(self):
        """Fetch actual .cab download URLs for this entry."""
        if not self.update_id:
            return []

        update_ids = json.dumps(
            {"size": 0, "updateID": self.update_id, "uidInfo": self.update_id}
        )
        post_data = {"updateIDs": f"[{update_ids}]"}

        try:
            resp = SESSION.post(
                DOWNLOAD_URL,
                data=post_data,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                timeout=30,
            )
            resp.raise_for_status()
        except requests.RequestException as e:
            print(f"  [!] Error fetching download links: {e}")
            return []

        matches = DOWNLOAD_PATTERN.findall(resp.text)
        return [url for _, url in matches]

    def __str__(self):
        date_str = (
            self.last_updated.strftime("%Y-%m-%d") if self.last_updated else "N/A"
        )
        return f"{self.title}  [{self.size_str}]  ({date_str})"


def search_catalog(query, max_pages=1):
    """Search the Microsoft Update Catalog and return a list of CatalogEntry objects."""
    entries = []
    search_safe = quote(query)
    url = f"{SEARCH_URL}?q={search_safe}"

    post_data = None
    page = 0

    while page < max_pages:
        try:
            if post_data:
                resp = SESSION.post(url, data=post_data, timeout=30)
            else:
                resp = SESSION.get(url, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as e:
            print(f"[!] Error searching catalog: {e}")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        table = soup.find(id="ctl00_catalogBody_updateMatches")
        if not table:
            print("[!] No results table found. The catalog may have changed its layout.")
            break

        rows = table.find_all("tr")
        if len(rows) <= 1:
            if page == 0:
                print("[i] No results found.")
            break

        for row in rows[1:]:
            try:
                entries.append(CatalogEntry(row))
            except (IndexError, AttributeError):
                continue

        # Check for next page
        next_page = soup.find(id="ctl00_catalogBody_nextPageLinkText")
        if not next_page:
            break

        # Build POST data for next page
        page_data = {"__EVENTTARGET": "ctl00$catalogBody$nextPageLinkText"}
        for field in [
            "__EVENTARGUMENT",
            "__EVENTVALIDATION",
            "__VIEWSTATE",
            "__VIEWSTATEGENERATOR",
        ]:
            el = soup.find(id=field)
            if el:
                page_data[field] = el.attrs.get("value", "")

        post_data = page_data
        page += 1

    return entries


def download_cab(url, output_dir):
    """Download a .cab file from the given URL into output_dir. Returns the file path."""
    filename = url.split("/")[-1]
    # Sanitize filename
    filename = re.sub(r"[^\w\-.]", "_", filename)
    if not filename.lower().endswith(".cab"):
        filename += ".cab"

    filepath = os.path.join(output_dir, filename)

    print(f"  Downloading: {filename}")
    try:
        resp = SESSION.get(url, stream=True, timeout=120)
        resp.raise_for_status()
        total = int(resp.headers.get("content-length", 0))
        downloaded = 0

        with open(filepath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
                downloaded += len(chunk)
                if total > 0:
                    pct = downloaded * 100 // total
                    bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
                    print(
                        f"\r  [{bar}] {pct}% ({downloaded // 1024} KB / {total // 1024} KB)",
                        end="",
                        flush=True,
                    )
        print()

        size_mb = os.path.getsize(filepath) / (1024 * 1024)
        print(f"  Saved: {filepath} ({size_mb:.2f} MB)")
        return filepath
    except requests.RequestException as e:
        print(f"  [!] Download failed: {e}")
        return None


def extract_cab(cab_path, output_dir=None):
    """Extract a .cab file. Tries multiple methods in order:
    1. cabextract (Linux/macOS)
    2. Windows expand command
    3. Python cabarchive package
    4. Pure-Python CAB parser (no dependencies)
    """
    if output_dir is None:
        base = os.path.splitext(os.path.basename(cab_path))[0]
        output_dir = os.path.join(os.path.dirname(cab_path), base)

    os.makedirs(output_dir, exist_ok=True)

    # Method 1: cabextract (Linux/macOS)
    if shutil.which("cabextract"):
        try:
            result = subprocess.run(
                ["cabextract", "-d", output_dir, cab_path],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode == 0:
                print(f"  Extracted to: {output_dir}")
                _extract_nested_cabs(output_dir)
                scan_extracted_files(output_dir)
                return output_dir
        except subprocess.TimeoutExpired:
            print("  [!] cabextract timed out")

    # Method 2: Windows expand command
    if platform.system() == "Windows":
        try:
            result = subprocess.run(
                ["expand", cab_path, "-F:*", output_dir],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode == 0:
                print(f"  Extracted to: {output_dir}")
                _extract_nested_cabs(output_dir)
                scan_extracted_files(output_dir)
                return output_dir
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

    # Method 3: Python cabarchive package
    try:
        import cabarchive

        cab = cabarchive.CabArchive(cab_path)
        for name in cab:
            dest = os.path.join(output_dir, name)
            parent = os.path.dirname(dest)
            if parent:
                os.makedirs(parent, exist_ok=True)
            with open(dest, "wb") as f:
                f.write(cab[name].buf)
        print(f"  Extracted to: {output_dir}")
        _extract_nested_cabs(output_dir)
        scan_extracted_files(output_dir)
        return output_dir
    except ImportError:
        pass
    except Exception as e:
        print(f"  [!] cabarchive failed: {e}")

    # Method 4: Pure-Python CAB parser (no dependencies needed)
    try:
        _extract_cab_pure_python(cab_path, output_dir)
        print(f"  Extracted to: {output_dir}")
        _extract_nested_cabs(output_dir)
        scan_extracted_files(output_dir)
        return output_dir
    except Exception as e:
        print(f"  [!] Extraction failed: {e}")
        return None


def _extract_cab_pure_python(cab_path, output_dir):
    """Pure-Python CAB extractor supporting NONE and MSZIP compression."""
    with open(cab_path, "rb") as f:
        data = f.read()

    # CAB header: signature, reserved, cabinet size, reserved, files offset,
    # reserved, version minor, version major, num folders, num files, flags
    if data[:4] != b"MSCF":
        raise ValueError("Not a valid CAB file (missing MSCF signature)")

    cab_size = struct.unpack_from("<I", data, 8)[0]
    files_offset = struct.unpack_from("<I", data, 16)[0]
    num_folders = struct.unpack_from("<H", data, 26)[0]
    num_files = struct.unpack_from("<H", data, 28)[0]
    flags = struct.unpack_from("<H", data, 30)[0]

    offset = 36

    # Handle reserved fields in header
    header_reserve = 0
    folder_reserve = 0
    data_reserve = 0
    if flags & 0x0004:  # cfhdrRESERVE_PRESENT
        header_reserve = struct.unpack_from("<H", data, offset)[0]
        folder_reserve = struct.unpack_from("<B", data, offset + 2)[0]
        data_reserve = struct.unpack_from("<B", data, offset + 3)[0]
        offset += 4 + header_reserve

    # Skip previous cabinet name if present
    if flags & 0x0001:  # cfhdrPREV_CABINET
        while data[offset] != 0:
            offset += 1
        offset += 1  # skip null terminator
        while data[offset] != 0:
            offset += 1
        offset += 1

    # Skip next cabinet name if present
    if flags & 0x0002:  # cfhdrNEXT_CABINET
        while data[offset] != 0:
            offset += 1
        offset += 1
        while data[offset] != 0:
            offset += 1
        offset += 1

    # Parse folders (CFFOLDER structures)
    folders = []
    for _ in range(num_folders):
        data_offset = struct.unpack_from("<I", data, offset)[0]
        num_data_blocks = struct.unpack_from("<H", data, offset + 4)[0]
        compress_type = struct.unpack_from("<H", data, offset + 6)[0]
        folders.append((data_offset, num_data_blocks, compress_type))
        offset += 8 + folder_reserve

    # Parse files (CFFILE structures) starting at files_offset
    files = []
    offset = files_offset
    for _ in range(num_files):
        usize = struct.unpack_from("<I", data, offset)[0]
        uoffset = struct.unpack_from("<I", data, offset + 4)[0]
        folder_index = struct.unpack_from("<H", data, offset + 8)[0]
        # date and time at offset+10 and offset+12
        attrs = struct.unpack_from("<H", data, offset + 14)[0]
        offset += 16
        # Read null-terminated filename
        name_start = offset
        while data[offset] != 0:
            offset += 1
        name = data[name_start:offset].decode("utf-8", errors="replace")
        offset += 1  # skip null terminator
        files.append((name, usize, uoffset, folder_index))

    # Decompress data blocks per folder
    folder_data = {}
    for folder_idx, (data_off, num_blocks, comp_type) in enumerate(folders):
        block_offset = data_off
        raw_data = bytearray()
        for _ in range(num_blocks):
            # CFDATA: checksum(4), compressed_size(2), uncompressed_size(2)
            # + data_reserve bytes + compressed data
            _checksum = struct.unpack_from("<I", data, block_offset)[0]
            comp_size = struct.unpack_from("<H", data, block_offset + 4)[0]
            uncomp_size = struct.unpack_from("<H", data, block_offset + 6)[0]
            block_offset += 8 + data_reserve
            block_data = data[block_offset : block_offset + comp_size]
            block_offset += comp_size

            if comp_type == 0:  # NONE
                raw_data.extend(block_data)
            elif comp_type == 1:  # MSZIP
                if block_data[:2] == b"CK":
                    block_data = block_data[2:]
                try:
                    decompressed = zlib.decompress(block_data, -zlib.MAX_WBITS)
                    raw_data.extend(decompressed)
                except zlib.error:
                    raw_data.extend(block_data)
            else:
                raise ValueError(
                    f"Unsupported compression type {comp_type} (LZX/Quantum). "
                    "Please install cabextract or cabarchive: pip install cabarchive"
                )

        folder_data[folder_idx] = bytes(raw_data)

    # Write files
    for name, usize, uoffset, folder_idx in files:
        # Normalize path separators
        name = name.replace("\\", os.sep).replace("/", os.sep)
        dest = os.path.join(output_dir, name)
        parent = os.path.dirname(dest)
        if parent:
            os.makedirs(parent, exist_ok=True)
        fdata = folder_data[folder_idx]
        with open(dest, "wb") as out:
            out.write(fdata[uoffset : uoffset + usize])


def _extract_nested_cabs(directory):
    """Recursively extract any .cab files found inside an extracted directory."""
    for root, _dirs, files in os.walk(directory):
        for f in files:
            if f.lower().endswith(".cab"):
                nested_path = os.path.join(root, f)
                nested_out = os.path.join(root, os.path.splitext(f)[0])
                print(f"  Extracting nested cab: {f}")
                extract_cab(nested_path, nested_out)


# Driver file extensions to look for, grouped by importance
DRIVER_FILE_TYPES = {
    ".sys": "Driver",
    ".inf": "Setup Info",
    ".cat": "Catalog/Signature",
    ".dll": "Library",
    ".exe": "Executable",
    ".mui": "Language Resource",
    ".man": "Manifest",
}


def _format_size(size_bytes):
    """Format byte count to human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


def scan_extracted_files(directory):
    """Scan extracted directory for driver files and print a summary."""
    if not directory or not os.path.isdir(directory):
        return

    found = {}  # ext -> list of (relative_path, size)
    other_files = []
    total_files = 0

    for root, _dirs, files in os.walk(directory):
        for f in files:
            fpath = os.path.join(root, f)
            rel = os.path.relpath(fpath, directory)
            size = os.path.getsize(fpath)
            ext = os.path.splitext(f)[1].lower()
            total_files += 1
            if ext in DRIVER_FILE_TYPES:
                found.setdefault(ext, []).append((rel, size))
            else:
                other_files.append((rel, size))

    if total_files == 0:
        print("  [!] No files found after extraction.")
        return

    print(f"\n  --- Driver File Summary ({total_files} files) ---")

    # Show .sys files first (the main driver binaries)
    if ".sys" in found:
        print(f"  [+] DRIVER FILES (.sys): {len(found['.sys'])} found")
        for rel, size in found[".sys"]:
            print(f"      >> {rel}  ({_format_size(size)})")
    else:
        print("  [-] No .sys driver files found")

    # Show .inf files (needed for driver installation)
    if ".inf" in found:
        print(f"  [+] Setup Info (.inf): {len(found['.inf'])} found")
        for rel, size in found[".inf"]:
            print(f"      {rel}  ({_format_size(size)})")

    # Show remaining driver-related file types
    for ext in [".cat", ".dll", ".exe", ".mui", ".man"]:
        if ext in found:
            label = DRIVER_FILE_TYPES[ext]
            print(f"  [+] {label} ({ext}): {len(found[ext])} found")
            for rel, size in found[ext]:
                print(f"      {rel}  ({_format_size(size)})")

    if other_files:
        print(f"  [i] Other files: {len(other_files)}")
        for rel, size in other_files:
            print(f"      {rel}  ({_format_size(size)})")

    print()


def _file_hash(filepath):
    """Compute SHA-256 hash of a file."""
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def deduplicate_sys_files(directory):
    """Find and remove duplicate .sys files across all extracted subdirectories.
    Duplicates are identified by filename + content hash.
    Keeps the newest copy (by modification time) and removes the rest.
    """
    if not directory or not os.path.isdir(directory):
        return

    # Collect all .sys files: group by lowercase filename
    sys_files = {}  # filename.lower() -> list of (full_path, size, mtime, hash)
    for root, _dirs, files in os.walk(directory):
        for f in files:
            if f.lower().endswith(".sys"):
                fpath = os.path.join(root, f)
                stat = os.stat(fpath)
                sys_files.setdefault(f.lower(), []).append(
                    (fpath, stat.st_size, stat.st_mtime)
                )

    removed_count = 0
    removed_size = 0

    for fname, copies in sys_files.items():
        if len(copies) < 2:
            continue

        # Compute hashes to group truly identical files
        by_hash = {}  # hash -> list of (path, size, mtime)
        for fpath, size, mtime in copies:
            fhash = _file_hash(fpath)
            by_hash.setdefault(fhash, []).append((fpath, size, mtime))

        for fhash, group in by_hash.items():
            if len(group) < 2:
                continue

            # Keep the newest copy, remove the rest
            group.sort(key=lambda x: x[2], reverse=True)
            keep_path = group[0][0]
            keep_rel = os.path.relpath(keep_path, directory)

            for dup_path, dup_size, _ in group[1:]:
                dup_rel = os.path.relpath(dup_path, directory)
                try:
                    os.remove(dup_path)
                    removed_count += 1
                    removed_size += dup_size
                    print(f"  [x] Removed duplicate: {dup_rel}")
                except OSError as e:
                    print(f"  [!] Could not remove {dup_rel}: {e}")

            if removed_count:
                print(f"  [=] Kept: {keep_rel}")

    if removed_count > 0:
        print(
            f"\n  --- Deduplication: removed {removed_count} duplicate .sys file(s) "
            f"({_format_size(removed_size)} freed) ---\n"
        )
    else:
        print("\n  --- No duplicate .sys files found ---\n")


def display_results(entries):
    """Print search results as a numbered list."""
    if not entries:
        return

    print(f"\n{'='*80}")
    print(f" Found {len(entries)} result(s)")
    print(f"{'='*80}\n")

    for i, entry in enumerate(entries, 1):
        title_wrapped = textwrap.fill(
            entry.title, width=70, subsequent_indent="       "
        )
        print(f"  [{i:3d}] {title_wrapped}")
        date_str = (
            entry.last_updated.strftime("%Y-%m-%d") if entry.last_updated else "N/A"
        )
        print(
            f"       Size: {entry.size_str}  |  Date: {date_str}  |  Class: {entry.classification}"
        )
        if entry.products:
            print(f"       Products: {', '.join(entry.products[:3])}")
        print()


def interactive_mode(entries, output_dir):
    """Let the user select which entries to download and extract."""
    while True:
        print("Enter selection (e.g. 1,3,5 or 1-5 or 'all' or 'q' to quit):")
        try:
            choice = input("> ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break

        if choice in ("q", "quit", "exit"):
            break

        if choice == "all":
            selected = list(range(len(entries)))
        else:
            selected = _parse_selection(choice, len(entries))

        if not selected:
            print("[!] Invalid selection. Try again.")
            continue

        for idx in selected:
            entry = entries[idx]
            print(f"\n--- Processing: {entry.title} ---")
            urls = entry.get_download_urls()
            if not urls:
                print("  [!] No download URLs found for this entry.")
                continue

            for url in urls:
                cab_path = download_cab(url, output_dir)
                if cab_path:
                    extract_cab(cab_path)

        deduplicate_sys_files(output_dir)
        print("Done! Select more entries or 'q' to quit.")


def _parse_selection(text, total):
    """Parse user selection like '1,3,5' or '1-5' into zero-based indices."""
    indices = set()
    parts = text.replace(" ", "").split(",")
    for part in parts:
        if "-" in part:
            try:
                start, end = part.split("-", 1)
                for i in range(int(start), int(end) + 1):
                    if 1 <= i <= total:
                        indices.add(i - 1)
            except ValueError:
                continue
        else:
            try:
                i = int(part)
                if 1 <= i <= total:
                    indices.add(i - 1)
            except ValueError:
                continue
    return sorted(indices)


def batch_mode(entries, output_dir, selections):
    """Download and extract specific entries (non-interactive)."""
    selected = _parse_selection(selections, len(entries))
    if not selected:
        print("[!] No valid selections. Use numbers like '1,3' or '1-5' or 'all'.")
        return

    for idx in selected:
        entry = entries[idx]
        print(f"\n--- Processing: {entry.title} ---")
        urls = entry.get_download_urls()
        if not urls:
            print("  [!] No download URLs found.")
            continue

        for url in urls:
            cab_path = download_cab(url, output_dir)
            if cab_path:
                extract_cab(cab_path)

    deduplicate_sys_files(output_dir)


def extract_local_cabs(paths, output_dir):
    """Extract local .cab files or all .cab files in a directory."""
    for path in paths:
        path = os.path.abspath(path)
        if os.path.isdir(path):
            for root, _dirs, files in os.walk(path):
                for f in files:
                    if f.lower().endswith(".cab"):
                        cab_path = os.path.join(root, f)
                        print(f"\n--- Extracting: {cab_path} ---")
                        out = os.path.join(
                            output_dir, os.path.splitext(f)[0]
                        )
                        extract_cab(cab_path, out)
        elif os.path.isfile(path):
            print(f"\n--- Extracting: {path} ---")
            out = os.path.join(output_dir, os.path.splitext(os.path.basename(path))[0])
            extract_cab(path, out)
        else:
            print(f"[!] Not found: {path}")

    deduplicate_sys_files(output_dir)


def main():
    parser = argparse.ArgumentParser(
        description="Search, download, and extract driver .cab files from the Windows Update Catalog.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples:
              Search and interactively download:
                %(prog)s "Realtek Audio"

              Search and auto-download first result:
                %(prog)s "NVIDIA" --download 1

              Download all results to a custom folder:
                %(prog)s "Intel Bluetooth" --download all --output ./intel_bt

              Just list results (no download):
                %(prog)s "USB 3.0 Host Controller" --list-only

              Fetch multiple pages of results:
                %(prog)s "Surface Pro" --pages 3

              Extract local .cab files:
                %(prog)s --extract driver.cab another.cab
                %(prog)s --extract ./cab_folder/
        """),
    )
    parser.add_argument(
        "search", nargs="?", help="Search query for the Windows Update Catalog"
    )
    parser.add_argument(
        "-o",
        "--output",
        default="./downloads",
        help="Output directory for downloaded/extracted files (default: ./downloads)",
    )
    parser.add_argument(
        "-d",
        "--download",
        metavar="SEL",
        help="Download selections without interactive prompt (e.g. '1,3', '1-5', 'all')",
    )
    parser.add_argument(
        "-p",
        "--pages",
        type=int,
        default=1,
        help="Number of result pages to fetch (default: 1, ~25 results per page)",
    )
    parser.add_argument(
        "--list-only",
        action="store_true",
        help="Only list results, do not download",
    )
    parser.add_argument(
        "-e",
        "--extract",
        nargs="+",
        metavar="PATH",
        help="Extract local .cab file(s) or all .cab files in a directory",
    )

    args = parser.parse_args()

    output_dir = os.path.abspath(args.output)
    os.makedirs(output_dir, exist_ok=True)

    # Local extraction mode
    if args.extract:
        extract_local_cabs(args.extract, output_dir)
        print(f"\n[*] Extracted files saved to: {output_dir}")
        return

    if not args.search:
        parser.print_help()
        print("\n[!] Please provide a search query or use --extract for local .cab files.")
        sys.exit(1)

    print(f"[*] Searching Windows Update Catalog for: '{args.search}'")
    entries = search_catalog(args.search, max_pages=args.pages)

    if not entries:
        print("[!] No results found. Try a different search term.")
        sys.exit(1)

    display_results(entries)

    if args.list_only:
        sys.exit(0)

    if args.download:
        if args.download.lower() == "all":
            batch_mode(
                entries,
                output_dir,
                ",".join(str(i) for i in range(1, len(entries) + 1)),
            )
        else:
            batch_mode(entries, output_dir, args.download)
    else:
        interactive_mode(entries, output_dir)

    print(f"\n[*] Files saved to: {output_dir}")


if __name__ == "__main__":
    main()
