from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
for path, start, end in [
    ('packages/pptx-engine/src/builtin-layouts.ts', 200, 216),
    ('packages/pptx-engine/src/index.ts', 1626, 1642),
    ('packages/pptx-engine/src/layout.ts', 236, 250),
    ('packages/pptx-engine/src/slide-transfer.ts', 235, 246),
]:
    lines = (ROOT / path).read_text(encoding='utf-8').splitlines()
    print(f'--- {path}:{start}-{end} ---')
    for no in range(start, min(end, len(lines)) + 1):
        print(f'{no:04d}: {lines[no-1]}')
