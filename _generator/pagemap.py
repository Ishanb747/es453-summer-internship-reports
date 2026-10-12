"""Find the printed page number of each index marker in the rendered PDFs."""
import json, subprocess, sys
sys.path.insert(0, sys.argv[1])
import report1 as R
import build_specs as S

def pages(pdf):
    n = int([l for l in subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout.splitlines() if l.startswith("Pages")][0].split()[1])
    out = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"], capture_output=True, text=True).stdout
        lines = [l.strip() for l in t.splitlines() if l.strip()]
        out.append(lines)
    return out

def resolve(pdf, spec):
    pg = pages(pdf)
    res = {}
    for key in ("contents", "annexure"):
        for label, marker in spec.get(key, []):
            for i, lines in enumerate(pg):
                if i < 2:  # cover + index
                    continue
                if any(l.startswith(marker) for l in lines[:4]):
                    res[marker] = lines[-1] if lines else ""
                    break
            else:
                res[marker] = "?"
    return res

out = {"r1": resolve(sys.argv[2], R.REPORT1_INDEX), "r2": resolve(sys.argv[3], S.REPORT2_INDEX)}
json.dump(out, open(sys.argv[4], "w"), indent=1)
print(json.dumps(out, indent=1))
