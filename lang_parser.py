"""
TWoM .lang binary parser.

Format (confirmed from french.lang):
  [data_size: uint32 LE] [record_count: uint32 LE]
  per record:
    [key_len: uint16 LE] [key: ASCII bytes]
    [val_len: uint16 LE] [value: UTF-16 LE, val_len chars incl. trailing CR]
"""
import struct
import csv
from dataclasses import dataclass
from pathlib import Path


@dataclass
class LangRecord:
    key: str
    value: str  # display text (CR stripped)


def load_lang(path: Path) -> list[LangRecord]:
    data = path.read_bytes()
    if len(data) < 8:
        raise ValueError("File too small.")
    record_count = struct.unpack_from('<I', data, 4)[0]
    if not (1 <= record_count <= 200_000):
        raise ValueError(f"Unexpected record count: {record_count}")
    pos, n, records = 8, len(data), []
    for i in range(record_count):
        if pos + 4 > n:
            raise ValueError(f"Unexpected EOF at record {i}")
        kl = struct.unpack_from('<H', data, pos)[0]; pos += 2
        key = data[pos:pos+kl].decode('ascii', errors='replace'); pos += kl
        vl = struct.unpack_from('<H', data, pos)[0]; pos += 2
        val = data[pos:pos+vl*2].decode('utf-16-le', errors='replace').rstrip('\r'); pos += vl*2
        records.append(LangRecord(key=key, value=val))
    return records


def build_lang(records: list[LangRecord], translations: dict[str, str]) -> bytes:
    body = bytearray()
    for rec in records:
        text  = translations.get(rec.key, rec.value)
        key_b = rec.key.encode('ascii')
        val_s = text + '\r'
        body += struct.pack('<H', len(key_b)) + key_b
        body += struct.pack('<H', len(val_s)) + val_s.encode('utf-16-le')
    header = struct.pack('<II', len(body) + 4, len(records))
    return bytes(header) + bytes(body)


# ── CSV helpers ──────────────────────────────────────────────────────────────

CSV_COLS = ['key', 'source', 'translation']

def save_csv(path: Path, records: list[LangRecord],
             translations: dict[str, str], source_lang: str = 'English') -> None:
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(CSV_COLS)
        for rec in records:
            w.writerow([rec.key, rec.value, translations.get(rec.key, '')])


def load_csv(path: Path) -> tuple[list[LangRecord], dict[str, str]]:
    """Returns (records_as_LangRecord_with_source, translations_dict)."""
    records, translations = [], {}
    with open(path, 'r', newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        if not {'key', 'source', 'translation'}.issubset(set(reader.fieldnames or [])):
            raise ValueError("CSV must have columns: key, source, translation")
        for row in reader:
            rec = LangRecord(key=row['key'], value=row['source'])
            records.append(rec)
            if row['translation'].strip():
                translations[row['key']] = row['translation']
    return records, translations
