"""Re-apply AddInUtils refactor to VisualAddIn.cpp without touching Cyrillic literals."""
from pathlib import Path
import subprocess

repo = Path(r"e:\VSProjects\visualRecognitionAddIn_Cursor")
path = repo / "src" / "VisualAddIn.cpp"

text = subprocess.check_output(
    ["git", "-C", str(repo), "show", "HEAD:src/VisualAddIn.cpp"],
    encoding="utf-8",
)
lines = text.splitlines(keepends=True)

start = end = None
for i, line in enumerate(lines):
    if start is None and line.strip() == "namespace {":
        start = i
    if start is not None and line.strip() == "} // namespace" and i > start:
        end = i
        break

if start is None or end is None:
    raise SystemExit("anonymous namespace block not found")

insert_at = start
new_lines = lines[:insert_at]
new_lines.extend(
    [
        '#include "AddInUtils.h"\n',
        "\n",
        "using namespace addin;\n",
        "\n",
    ]
)
new_lines.extend(lines[end + 1 :])

# Replace duplicate includes if AddInUtils already present
out = "".join(new_lines)
out = out.replace('#include "AddInUtils.h"\n\n#include "AddInUtils.h"\n', '#include "AddInUtils.h"\n')
if '#include "AddInUtils.h"' not in out:
    out = out.replace(
        '#include "GeminiPdfClient.h"\n',
        '#include "AddInUtils.h"\n#include "AnthropicMessagesClient.h"\n#include "GeminiPdfClient.h"\n',
        1,
    )

path.write_bytes(out.encode("utf-8"))  # no BOM
print(f"Wrote {path} ({len(out.encode('utf-8'))} bytes, no BOM)")
