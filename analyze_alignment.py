"""Analyze SafiPrint.rpt for misaligned headers/objects."""
from collections import defaultdict
from crystalreports import CrystalReport

AREA_NAMES = {
    1: "Report Header", 2: "Page Header", 3: "Group Header",
    4: "Detail", 5: "Group Footer", 6: "Report Footer", 7: "Page Footer",
}

with CrystalReport("SafiPrint.rpt") as rpt:
    print(f"Report: {rpt}")
    print(f"Margins (L,R,T,B): {rpt.get_margins()}")
    print()

    # Group objects by section
    sections = defaultdict(list)
    for obj in rpt.objects:
        sections[obj.section_code].append(obj)

    # Analyze each section
    for code in sorted(sections.keys()):
        area = code // 6000
        sub = (code - area * 6000) // 50
        area_name = AREA_NAMES.get(area, "?")
        label = f"{area_name}" + (f" #{sub}" if sub > 0 else "")
        objects = sections[code]
        height = rpt.get_section_height(code)

        print(f"{'='*70}")
        print(f"Section {code} — {label} (height={height}, {len(objects)} objects)")
        print(f"{'='*70}")

        if not objects:
            continue

        # Print all objects sorted by left position
        objects_sorted = sorted(objects, key=lambda o: (o.top, o.left))
        for obj in objects_sorted:
            w = obj.right - obj.left
            h = obj.bottom - obj.top
            # Note: bottom < top means the object uses inverted coords (height)
            print(f"  {obj.name:35s} {obj.object_type:8s} "
                  f"L={obj.left:5d} T={obj.top:5d} R={obj.right:5d} B={obj.bottom:5d} "
                  f"(w={w:5d} h={obj.bottom - obj.top:5d})")

        # --- Alignment analysis ---
        # Group objects by approximate top position (within 20 twips tolerance)
        TOLERANCE = 20  # twips (~0.35mm)
        rows = []
        used = set()
        for obj in objects_sorted:
            if id(obj) in used:
                continue
            row = [obj]
            used.add(id(obj))
            for other in objects_sorted:
                if id(other) in used:
                    continue
                if abs(other.top - obj.top) <= TOLERANCE:
                    row.append(other)
                    used.add(id(other))
            if len(row) > 1:
                rows.append(row)

        # Check each row for misalignment
        misaligned = []
        for row in rows:
            tops = [o.top for o in row]
            bottoms = [o.bottom for o in row]
            if len(set(tops)) > 1:
                # Objects on same row have different top values
                avg_top = sum(tops) // len(tops)
                for obj in row:
                    if obj.top != avg_top:
                        diff = obj.top - avg_top
                        misaligned.append((obj, "top", diff, avg_top))
            if len(set(bottoms)) > 1:
                avg_bottom = sum(bottoms) // len(bottoms)
                for obj in row:
                    if abs(obj.bottom - avg_bottom) > TOLERANCE:
                        diff = obj.bottom - avg_bottom
                        misaligned.append((obj, "bottom", diff, avg_bottom))

        if misaligned:
            print()
            print(f"  *** MISALIGNED OBJECTS:")
            for obj, edge, diff, expected in misaligned:
                direction = "te laag" if diff > 0 else "te hoog"
                print(f"      {obj.name:30s} {edge}={getattr(obj, edge):5d} "
                      f"(verwacht ~{expected}, {direction} met {abs(diff)} twips)")
        print()

    # --- Overall summary ---
    print(f"\n{'='*70}")
    print("SAMENVATTING — Alignment check per sectie-rij")
    print(f"{'='*70}")

    total_issues = 0
    for code in sorted(sections.keys()):
        area = code // 6000
        sub = (code - area * 6000) // 50
        area_name = AREA_NAMES.get(area, "?")
        label = f"{area_name}" + (f" #{sub}" if sub > 0 else "")
        objects = sections[code]
        if len(objects) < 2:
            continue

        objects_sorted = sorted(objects, key=lambda o: (o.top, o.left))
        rows = []
        used = set()
        for obj in objects_sorted:
            if id(obj) in used:
                continue
            row = [obj]
            used.add(id(obj))
            for other in objects_sorted:
                if id(other) in used:
                    continue
                if abs(other.top - obj.top) <= TOLERANCE:
                    row.append(other)
                    used.add(id(other))
            if len(row) > 1:
                rows.append(row)

        section_issues = []
        for row in rows:
            tops = [o.top for o in row]
            if len(set(tops)) > 1:
                names = [o.name for o in row]
                min_t, max_t = min(tops), max(tops)
                section_issues.append(
                    f"    Rij (top~{min_t}): {', '.join(names)} — "
                    f"verschil {max_t - min_t} twips"
                )

        if section_issues:
            print(f"\n  {label} (section {code}):")
            for issue in section_issues:
                print(issue)
                total_issues += 1

    if total_issues == 0:
        print("\n  Alles staat recht!")
    else:
        print(f"\n  Totaal: {total_issues} rijen met alignment issues")
