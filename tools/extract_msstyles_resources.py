from __future__ import annotations

import pathlib
import struct
import sys


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: extract_msstyles_resources.py <msstyles> <out-dir>")
        return 2

    source = pathlib.Path(sys.argv[1])
    out_dir = pathlib.Path(sys.argv[2])
    data = source.read_bytes()
    out_dir.mkdir(parents=True, exist_ok=True)

    pe_offset = struct.unpack_from("<I", data, 0x3C)[0]
    section_count = struct.unpack_from("<H", data, pe_offset + 6)[0]
    optional_header_size = struct.unpack_from("<H", data, pe_offset + 20)[0]
    optional_header = pe_offset + 24
    magic = struct.unpack_from("<H", data, optional_header)[0]
    data_directories = optional_header + (112 if magic == 0x20B else 96)
    resource_rva, _resource_size = struct.unpack_from("<II", data, data_directories + 16)
    section_offset = optional_header + optional_header_size

    sections = []
    for index in range(section_count):
        offset = section_offset + index * 40
        virtual_size, virtual_address, raw_size, raw_pointer = struct.unpack_from("<IIII", data, offset + 8)
        sections.append((virtual_address, max(virtual_size, raw_size), raw_pointer))

    def rva_to_offset(rva: int) -> int:
        for virtual_address, size, raw_pointer in sections:
            if virtual_address <= rva < virtual_address + size:
                return raw_pointer + (rva - virtual_address)
        raise ValueError(f"RVA not mapped: 0x{rva:x}")

    resource_base = rva_to_offset(resource_rva)

    def entry_name(value: int) -> str:
        if value & 0x80000000:
            name_offset = resource_base + (value & 0x7FFFFFFF)
            length = struct.unpack_from("<H", data, name_offset)[0]
            raw_name = data[name_offset + 2 : name_offset + 2 + length * 2]
            return raw_name.decode("utf-16le", errors="ignore")
        return str(value & 0xFFFF)

    resources: list[tuple[list[str], bytes]] = []

    def walk(directory_relative_offset: int, parts: list[str]) -> None:
        directory_offset = resource_base + directory_relative_offset
        named_count, id_count = struct.unpack_from("<HH", data, directory_offset + 12)

        for entry_index in range(named_count + id_count):
            entry_offset = directory_offset + 16 + entry_index * 8
            name_value, offset_value = struct.unpack_from("<II", data, entry_offset)
            name = entry_name(name_value)
            target = offset_value & 0x7FFFFFFF

            if offset_value & 0x80000000:
                walk(target, parts + [name])
            else:
                data_entry_offset = resource_base + target
                rva, size, _codepage, _reserved = struct.unpack_from("<IIII", data, data_entry_offset)
                blob_offset = rva_to_offset(rva)
                resources.append((parts + [name], data[blob_offset : blob_offset + size]))

    walk(0, [])

    counts: dict[str, int] = {}
    for parts, blob in resources:
        resource_type = parts[0] if parts else "root"
        counts[resource_type] = counts.get(resource_type, 0) + 1
        extension = ".bin"

        if blob.startswith(b"\x89PNG"):
            extension = ".png"
        elif blob.startswith(b"BM"):
            extension = ".bmp"
        elif len(blob) >= 4 and struct.unpack_from("<I", blob, 0)[0] in (12, 40, 108, 124):
            extension = ".dib"

        safe_name = "__".join(
            "".join(character if character.isalnum() or character in "._-" else "_" for character in part)
            for part in parts
        )
        (out_dir / f"{safe_name}{extension}").write_bytes(blob)

    print(f"extracted {len(resources)} resources")
    for resource_type, count in sorted(counts.items()):
        print(f"{resource_type}: {count}")
    print(out_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
