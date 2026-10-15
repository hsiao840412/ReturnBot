#!/usr/bin/env python3
"""Build public notes from the approved, newest-first version history."""
import argparse
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent.parent
HEADER = re.compile(r'v(\d+\.\d+(?:\.\d+)?)')


def read_history(path):
    sections = []
    for block in re.split(r'(?m)(?=^v\d)', path.read_text(encoding='utf-8').strip()):
        if not block.strip():
            continue
        lines = block.strip().splitlines()
        match = HEADER.fullmatch(lines[0])
        if not match or len(lines) < 3 or lines[1] != '':
            raise ValueError('Each version requires a version title, blank line, and bullets')
        if not all(line.startswith('- ') and line[2:].strip() for line in lines[2:]):
            raise ValueError('Release changes must be nonempty bullet points')
        sections.append((match[1], '\n'.join(lines) + '\n'))
    versions = [tuple(map(int, v.split('.'))) + (0,) * (3 - len(v.split('.'))) for v, _ in sections]
    if not versions or any(a <= b for a, b in zip(versions, versions[1:])):
        raise ValueError('Versions must be unique and ordered newest first')
    return sections


def generate(version, history=ROOT / 'packaging/RELEASE_NOTES.md', output=ROOT / 'dist'):
    sections = read_history(history)
    if sections[0][0] != version:
        raise ValueError('Packaged version must match the newest release notes')
    output.mkdir(parents=True, exist_ok=True)
    (output / 'GITHUB_RELEASE_NOTES.md').write_text(sections[0][1], encoding='utf-8')
    full = '\n'.join(section for _, section in sections)
    (output / 'APP_RELEASE_NOTES.md').write_text(full, encoding='utf-8')
    return full


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('version')
    generate(parser.parse_args().version)
