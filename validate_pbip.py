#!/usr/bin/env python3
"""
validate_pbip.py — static integrity checker for a Power BI Project (PBIP).

Cannot execute DAX (no Analysis Services on this host), but catches the failure
modes that hand-authored PBIP actually ships with:

  E1  a visual/page/report references a table Entity that the model has no table for
  E2  a visual/page/report references a Property (measure/column) that Entity does not expose
  E3  a measure's DAX calls [Name] that is neither a measure nor a column in the model
  E4  a measure's DAX calls Table[Column] where Table or Column does not exist
  E5  pages.json pageOrder / activePageName is out of sync with the page folders
  E6  report.json / page.json points at a RegisteredResources file that is missing on disk
  W*  the same, downgraded to warnings where a false positive is plausible

Exit code is non-zero if any Ex fires.
"""
import json, re, sys, pathlib, collections

def find(root, pat):
    return sorted(pathlib.Path(root).rglob(pat))

# ---------- TMDL model parse -------------------------------------------------
IDENT = r"(?:'([^']+)'|([^\s'=]+))"

def ident(m):
    return m.group(1) if m.group(1) is not None else m.group(2)

def parse_model(model_dir):
    """Return (entities:set, props:{table:set(names)}, measures:set, columns:{table:set})."""
    props = collections.defaultdict(set)
    measures, entities = set(), set()
    columns = collections.defaultdict(set)
    for tf in find(model_dir, "*.tmdl"):
        if tf.parent.name != "tables":
            continue
        table = None
        for raw in tf.read_text(encoding="utf-8").splitlines():
            line = raw.rstrip()
            if re.match(r"^(createOrReplace\s+)?table\s+", line):
                m = re.match(r"^(?:createOrReplace\s+)?table\s+" + IDENT, line)
                if m:
                    table = ident(m); entities.add(table)
                continue
            mm = re.match(r"^\s+measure\s+" + IDENT, line)
            if mm and table:
                n = ident(mm); props[table].add(n); measures.add(n); continue
            mc = re.match(r"^\s+column\s+" + IDENT, line)
            if mc and table:
                n = ident(mc); props[table].add(n); columns[table].add(n)
    return entities, props, measures, columns

# ---------- report JSON reference walk ---------------------------------------
def walk_refs(node, out):
    """Collect (Entity, Property) from every Measure/Column/Hierarchy leaf."""
    if isinstance(node, dict):
        if "Property" in node and "Expression" in node and isinstance(node["Expression"], dict):
            ent = (node["Expression"].get("SourceRef") or {}).get("Entity")
            if ent is not None:
                out.append((ent, node["Property"]))
        for v in node.values():
            walk_refs(v, out)
    elif isinstance(node, list):
        for v in node:
            walk_refs(v, out)

# ---------- DAX measure body parse -------------------------------------------
def parse_measure_bodies(measures_file):
    """Yield (measure_name, body_text)."""
    text = pathlib.Path(measures_file).read_text(encoding="utf-8")
    lines = text.splitlines()
    out, cur, buf = [], None, []
    for raw in lines:
        m = re.match(r"^\s+measure\s+" + IDENT, raw)
        if m:
            if cur:
                out.append((cur, "\n".join(buf)))
            cur = ident(m); buf = [raw]
        elif cur is not None:
            # a new top-level item (column/partition) closes the measure
            if re.match(r"^\s+(column|partition)\s+", raw) or re.match(r"^\S", raw):
                out.append((cur, "\n".join(buf))); cur = None; buf = []
            else:
                buf.append(raw)
    if cur:
        out.append((cur, "\n".join(buf)))
    return out

def main(pbip_root):
    root = pathlib.Path(pbip_root)
    model_dir = next(root.glob("*.SemanticModel/definition"), None)
    report_def = next(root.glob("*.Report/definition"), None)
    assert model_dir and report_def, f"not a PBIP: {pbip_root}"

    entities, props, measures, columns = parse_model(model_dir)
    all_columns = {c for s in columns.values() for c in s}

    errors, warnings = [], []

    # --- E1/E2 report references
    ref_files = find(report_def, "*.json")
    total_refs = 0
    for jf in ref_files:
        try:
            data = json.loads(jf.read_text(encoding="utf-8"))
        except Exception as e:
            errors.append(f"E0 invalid JSON {jf.relative_to(root)}: {e}"); continue
        refs = []
        walk_refs(data, refs)
        for ent, prop in refs:
            total_refs += 1
            if ent not in entities:
                errors.append(f"E1 {jf.relative_to(root)} -> unknown table '{ent}'")
            elif prop not in props[ent]:
                errors.append(f"E2 {jf.relative_to(root)} -> '{ent}' has no measure/column '{prop}'")

    # --- E3/E4 DAX cross references
    mfile = next(model_dir.glob("tables/_Measures.tmdl"), None)
    if mfile:
        known = measures | all_columns
        for name, body in parse_measure_bodies(mfile):
            # strip TMDL comments, then blank out string literals so text
            # like "... & MID(x) & ..." is not mistaken for a [column] ref
            dax = re.sub(r"^\s*///.*$", "", body, flags=re.M)
            dax = re.sub(r'"(?:[^"]|"")*"', '""', dax)
            # bare [X] references (exclude [@temp] ADDCOLUMNS/SELECTCOLUMNS locals)
            for ref in re.findall(r"(?<![\w'])\[([^\]@][^\]]*)\]", dax):
                if ref not in known:
                    warnings.append(f"W3 measure '{name}' calls [{ref}] — not a known measure/column")
            # Table[Column] references
            for m in re.finditer(r"(?<![\w])(?:'([^']+)'|([A-Za-z_][\w]*))\[([^\]@][^\]]*)\]", dax):
                tbl = m.group(1) or m.group(2); col = m.group(3)
                if tbl not in entities:
                    warnings.append(f"W4 measure '{name}' references {tbl}[{col}] — unknown table")
                elif col not in props[tbl]:
                    warnings.append(f"W4 measure '{name}' references {tbl}[{col}] — no such column/measure")

    # --- E5 page wiring
    pages_dir = report_def / "pages"
    folders = {p.name for p in pages_dir.iterdir() if p.is_dir()}
    pj = json.loads((pages_dir / "pages.json").read_text(encoding="utf-8"))
    order = set(pj.get("pageOrder", []))
    for f in folders - order:
        errors.append(f"E5 page folder '{f}' missing from pages.json pageOrder")
    for o in order - folders:
        errors.append(f"E5 pageOrder lists '{o}' but no folder exists")
    if pj.get("activePageName") not in folders:
        errors.append(f"E5 activePageName '{pj.get('activePageName')}' has no page folder")
    # each page.json name matches its folder
    for f in folders:
        p = json.loads((pages_dir / f / "page.json").read_text(encoding="utf-8"))
        if p.get("name") != f:
            warnings.append(f"W5 page.json name '{p.get('name')}' != folder '{f}'")

    # --- E6 registered resources exist
    rj = json.loads((report_def / "report.json").read_text(encoding="utf-8"))
    resdir = report_def.parent / "StaticResources" / "RegisteredResources"
    for pkg in rj.get("resourcePackages", []):
        for it in pkg.get("items", []):
            if resdir and not (resdir / it["path"]).exists():
                errors.append(f"E6 report.json resource missing on disk: {it['path']}")

    name = root.name
    print(f"\n===== {name} =====")
    print(f"  tables={len(entities)}  measures={len(measures)}  columns={len(all_columns)}")
    print(f"  report json files={len(ref_files)}  field references checked={total_refs}")
    print(f"  ERRORS={len(errors)}  WARNINGS={len(warnings)}")
    for e in errors:
        print("  ✗", e)
    for w in warnings[:80]:
        print("  ·", w)
    if len(warnings) > 80:
        print(f"  … +{len(warnings)-80} more warnings")
    return len(errors)

if __name__ == "__main__":
    rc = 0
    for arg in sys.argv[1:]:
        rc += main(arg)
    sys.exit(1 if rc else 0)
