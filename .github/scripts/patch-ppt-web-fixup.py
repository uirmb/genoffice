from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

def read(path): return (ROOT / path).read_text(encoding='utf-8')
def write(path, text): (ROOT / path).write_text(text, encoding='utf-8')

# Normalize malformed helper calls left by the broad Buffer migration. This parser
# removes an accidental encoding argument from encodeUtf8/decodeBase64 while
# respecting nested parentheses and quoted strings.
def matching_paren(text, open_pos):
    depth = 0
    quote = None
    escape = False
    for i in range(open_pos, len(text)):
        ch = text[i]
        if quote:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == quote:
                quote = None
            continue
        if ch in "'\"`":
            quote = ch
            continue
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth == 0:
                return i
    raise RuntimeError('unbalanced call')

def split_top_level(value):
    parts = []
    start = 0
    depths = {'(': 0, '[': 0, '{': 0}
    pairs = {')': '(', ']': '[', '}': '{'}
    quote = None
    escape = False
    for i, ch in enumerate(value):
        if quote:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == quote:
                quote = None
            continue
        if ch in "'\"`":
            quote = ch
            continue
        if ch in depths:
            depths[ch] += 1
        elif ch in pairs:
            depths[pairs[ch]] -= 1
        elif ch == ',' and all(v == 0 for v in depths.values()):
            parts.append(value[start:i].strip())
            start = i + 1
    parts.append(value[start:].strip())
    return parts

def normalize_call_args(text, fn, forbidden):
    cursor = 0
    while True:
        token = fn + '('
        pos = text.find(token, cursor)
        if pos < 0:
            return text
        open_pos = pos + len(fn)
        close = matching_paren(text, open_pos)
        inner = text[open_pos + 1:close]
        args = split_top_level(inner)
        if len(args) > 1 and args[-1].strip(" \t\r\n'\"") in forbidden:
            replacement = fn + '(' + ', '.join(args[:-1]) + ')'
            text = text[:pos] + replacement + text[close + 1:]
            cursor = pos + len(replacement)
        else:
            cursor = close + 1

for file in (ROOT / 'packages/pptx-engine/src').glob('*.ts'):
    text = file.read_text(encoding='utf-8')
    fixed = normalize_call_args(text, 'encodeUtf8', {'utf8'})
    fixed = normalize_call_args(fixed, 'decodeBase64', {'base64'})
    fixed = normalize_call_args(fixed, 'encodeAscii', {'ascii'})
    if fixed != text:
        file.write_text(fixed, encoding='utf-8')

# App lifecycle labels belong to App(), whose i18n hook returns only lang.
path = 'apps/slides/src/renderer/App.tsx'
text = read(path)
if "import { slidesWebLifecycleLabels } from './web-labels'" not in text:
    text = text.replace("import * as fileActions from './file-actions'\n", "import * as fileActions from './file-actions'\nimport { slidesWebLifecycleLabels } from './web-labels'\n", 1)
text = text.replace(
    "export function App() {\n  const { lang } = useI18n()\n",
    "export function App() {\n  const { lang } = useI18n()\n  const webLabels = slidesWebLifecycleLabels(lang)\n",
    1,
)
write(path, text)

# Ribbon has helper components using useI18n before Ribbon itself; scope labels to
# the actual Ribbon Props function instead of the first hook in the module.
path = 'apps/slides/src/renderer/components/Ribbon.tsx'
text = read(path)
if "import { slidesWebLifecycleLabels } from '../web-labels'" not in text:
    marker = "import { useI18n, type StringKey } from '../i18n/locale'\n"
    text = text.replace(marker, marker + "import { slidesWebLifecycleLabels } from '../web-labels'\n", 1)
needle = "}: Props) {\n  const { t } = useI18n()\n"
if needle in text:
    text = text.replace(
        needle,
        "}: Props) {\n  const { t, lang } = useI18n()\n  const webLabels = slidesWebLifecycleLabels(lang)\n",
        1,
    )
elif "}: Props) {\n  const { t, lang } = useI18n()\n" in text and "const webLabels = slidesWebLifecycleLabels(lang)" not in text[text.find("}: Props) {"):]:
    text = text.replace("}: Props) {\n  const { t, lang } = useI18n()\n", "}: Props) {\n  const { t, lang } = useI18n()\n  const webLabels = slidesWebLifecycleLabels(lang)\n", 1)
write(path, text)

# Keep the dynamic setMode class toggle; only initialize the startup class as edit.
path = 'apps/slides/src/web/slides-api.ts'
text = read(path)
setmode = "    mode = nextMode\n    document.documentElement.classList.remove('office-view-mode')\n"
if setmode in text:
    text = text.replace(setmode, "    mode = nextMode\n    document.documentElement.classList.toggle('office-view-mode', mode === 'view')\n", 1)
final = "\n  document.documentElement.classList.toggle('office-view-mode', mode === 'view')\n\n  return {\n"
if final in text:
    text = text.replace(final, "\n  document.documentElement.classList.remove('office-view-mode')\n\n  return {\n", 1)
write(path, text)

print('PPT Web fixup applied')
