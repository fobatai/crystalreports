"""Fix alignment and set professional fonts in SafiPrint.rpt.

Pass 1: Align objects that are slightly off within the same row.
Pass 2: Kopjes (Text objects) -> Calibri 9
Pass 3: Data (Field objects) -> Arial 8
"""
import shutil
import os
from collections import defaultdict, Counter
from crystalreports import CrystalReport

INPUT = "SafiPrint.rpt"
OUTPUT = "SafiPrint_fixed.rpt"
TEMP = "SafiPrint_temp_save.rpt"
TOLERANCE = 20  # twips

CONTAINER_TYPES = {"Box", "Line"}
AREA_NAMES = {1: "RH", 2: "PH", 3: "GH", 4: "D",
              5: "GF", 6: "RF", 7: "PF"}


def section_label(code):
    area = code // 6000
    sub = (code - area * 6000) // 50
    return f"{AREA_NAMES.get(area, '?')}#{sub}"


fixes = []

with CrystalReport(INPUT) as rpt:
    # Group objects by section
    sections = defaultdict(list)
    for obj in rpt.objects:
        sections[obj.section_code].append(obj)

    # =========================================================
    # PASS 1: Fix alignment
    # =========================================================
    print("=== PASS 1: Alignment fixes ===\n")

    for code in sorted(sections.keys()):
        objects = sections[code]
        if len(objects) < 2:
            continue

        containers = [o for o in objects if o.object_type in CONTAINER_TYPES]
        content = [o for o in objects if o.object_type not in CONTAINER_TYPES]

        for group_name, group in [("container", containers), ("content", content)]:
            if len(group) < 2:
                continue

            sorted_objs = sorted(group, key=lambda o: o.top)
            rows = []
            used = set()

            for obj in sorted_objs:
                if obj.handle in used:
                    continue
                row = [obj]
                used.add(obj.handle)
                for other in sorted_objs:
                    if other.handle in used:
                        continue
                    if abs(other.top - obj.top) <= TOLERANCE:
                        row.append(other)
                        used.add(other.handle)
                if len(row) > 1:
                    rows.append(row)

            for row in rows:
                tops = [o.top for o in row]
                if len(set(tops)) <= 1:
                    continue

                top_counts = Counter(tops)
                target_top = top_counts.most_common(1)[0][0]
                if top_counts.most_common(1)[0][1] == 1:
                    target_top = round(sum(tops) / len(tops))

                for obj in row:
                    if obj.top != target_top:
                        delta = target_top - obj.top
                        new_bottom = obj.bottom + delta
                        label = section_label(code)
                        try:
                            rpt.move_object(
                                obj.handle,
                                obj.left, target_top,
                                obj.right, new_bottom,
                                section_code=code,
                            )
                            fixes.append(
                                f"  {label:8s} {obj.name:35s} "
                                f"top {obj.top:5d} -> {target_top:5d} ({delta:+d})"
                            )
                        except Exception as e:
                            fixes.append(
                                f"  {label:8s} {obj.name:35s} "
                                f"SKIP ({e})"
                            )

    for f in fixes:
        print(f)
    print(f"\n  {len(fixes)} objecten verwerkt")

    # =========================================================
    # PASS 1b: Fix box pairs (kopje + content box alignment)
    # =========================================================
    print("\n=== PASS 1b: Box pairs rechttrekken ===\n")

    box_fixes = []
    # Target left/right for consistent box edges
    TARGET_LEFT = 129
    TARGET_RIGHT = 10170

    for code in sorted(sections.keys()):
        objects = sections[code]
        label = section_label(code)
        boxes = sorted(
            [o for o in objects if o.object_type == "Box"],
            key=lambda o: o.top,
        )
        if len(boxes) < 2:
            continue

        # Identify kopje box (top, usually B~298) and content box
        if len(boxes) == 3:
            # GH#1 pattern: 2 nested boxes at top + content below
            # Use the outermost kopje box (largest height of top two)
            top_two = boxes[:2]
            kopje = max(top_two, key=lambda b: b.bottom - b.top)
            inner = min(top_two, key=lambda b: b.bottom - b.top)
            content = boxes[2]
        elif len(boxes) == 2:
            kopje, content = boxes[0], boxes[1]
            inner = None
            # Skip layered boxes (content starts above kopje bottom)
            if content.top < kopje.top:
                box_fixes.append(f"  {label:8s} SKIP layered boxes")
                continue
        else:
            continue

        # Fix content box top = kopje box bottom (no gap, no overlap)
        gap = content.top - kopje.bottom
        need_fix_top = gap != 0
        need_fix_lr = (content.left != TARGET_LEFT
                       or content.right != TARGET_RIGHT)
        need_fix_kopje_lr = (kopje.left != TARGET_LEFT
                             or kopje.right != TARGET_RIGHT)

        # Determine effective right edge: try TARGET_RIGHT, but
        # cross-section boxes can't have their right changed — in that
        # case match the kopje to the content box's actual right.
        effective_right = TARGET_RIGHT

        if need_fix_top or need_fix_lr:
            new_top = kopje.bottom
            height = content.bottom - content.top
            try:
                rpt.move_object(
                    content.handle,
                    TARGET_LEFT, new_top,
                    TARGET_RIGHT, new_top + height,
                    section_code=code,
                )
                # Re-read actual coordinates (cross-section boxes
                # may not accept right-edge changes)
                actual = [
                    o for o in rpt.get_objects_in_section(code)
                    if o.handle == content.handle
                ]
                if actual and actual[0].right != TARGET_RIGHT:
                    effective_right = actual[0].right
                parts = []
                if need_fix_top:
                    parts.append(f"top {content.top}->{new_top} (gap {gap:+d})")
                if need_fix_lr:
                    if effective_right != TARGET_RIGHT:
                        parts.append(
                            f"L->{TARGET_LEFT} (R={effective_right} cross-sec)")
                    else:
                        parts.append(f"L/R->{TARGET_LEFT}/{TARGET_RIGHT}")
                box_fixes.append(
                    f"  {label:8s} {content.name:10s} {', '.join(parts)}"
                )
            except Exception as e:
                box_fixes.append(f"  {label:8s} {content.name:10s} SKIP ({e})")

        # Fix kopje box left/right — match the effective right edge
        kopje_right = effective_right
        need_fix_kopje_lr = (kopje.left != TARGET_LEFT
                             or kopje.right != kopje_right)
        if need_fix_kopje_lr:
            try:
                rpt.move_object(
                    kopje.handle,
                    TARGET_LEFT, kopje.top,
                    kopje_right, kopje.bottom,
                    section_code=code,
                )
                box_fixes.append(
                    f"  {label:8s} {kopje.name:10s} L/R->{TARGET_LEFT}/{kopje_right}"
                )
            except Exception as e:
                box_fixes.append(f"  {label:8s} {kopje.name:10s} SKIP ({e})")

        # Fix inner nested box if present (GH#1)
        if inner and (inner.left != TARGET_LEFT
                      or inner.right != kopje_right):
            try:
                rpt.move_object(
                    inner.handle,
                    TARGET_LEFT, inner.top,
                    kopje_right, inner.bottom,
                    section_code=code,
                )
                box_fixes.append(
                    f"  {label:8s} {inner.name:10s} L/R->{TARGET_LEFT}/{kopje_right}"
                )
            except Exception as e:
                box_fixes.append(f"  {label:8s} {inner.name:10s} SKIP ({e})")

    for f in box_fixes:
        print(f)
    print(f"\n  {len(box_fixes)} box-fixes verwerkt")

    # =========================================================
    # PASS 2: Kopjes -> Calibri 9
    # =========================================================
    print("\n=== PASS 2: Kopjes -> Calibri 9 ===\n")

    kopje_fixes = []
    for code in sorted(sections.keys()):
        objects = sections[code]
        label = section_label(code)
        texts = [o for o in objects if o.object_type == "Text"]
        if not texts:
            continue

        # scope=2 sets ALL Text objects in the section
        try:
            rpt.set_section_font(
                code, face_name="Calibri", point_size=9, scope=2,
            )
            names = [t.name for t in texts]
            kopje_fixes.append(
                f"  {label:8s} Calibri 9 -> {', '.join(names)}"
            )
        except Exception as e:
            kopje_fixes.append(f"  {label:8s} SKIP ({e})")

    for f in kopje_fixes:
        print(f)
    print(f"\n  {len(kopje_fixes)} secties verwerkt")

    # =========================================================
    # PASS 3: Data fields -> Arial 8
    # =========================================================
    print("\n=== PASS 3: Data fields -> Arial 8 ===\n")

    field_fixes = []
    for code in sorted(sections.keys()):
        objects = sections[code]
        label = section_label(code)
        fields = [o for o in objects if o.object_type == "Field"]
        if not fields:
            continue

        # scope=1 sets ALL Field objects in the section
        try:
            rpt.set_section_font(
                code, face_name="Arial", point_size=8, scope=1,
            )
            names = [f.name for f in fields]
            field_fixes.append(
                f"  {label:8s} Arial 8  -> {', '.join(names)}"
            )
        except Exception as e:
            field_fixes.append(f"  {label:8s} SKIP ({e})")

    for f in field_fixes:
        print(f)
    print(f"\n  {len(field_fixes)} secties verwerkt")

    # =========================================================
    # PASS 4: Page Footer fields -> Arial 8 (individual fallback)
    # =========================================================
    # PESetFont(scope=1/2) fails on PF with error 572, so we
    # use set_field_font per object instead.
    print("\n=== PASS 4: Page Footer fields -> Arial 8 ===\n")

    pf_fixes = []
    pf_code = 42000
    if pf_code in sections:
        for obj in sections[pf_code]:
            if obj.object_type == "Field":
                try:
                    rpt.set_field_font(obj.handle, face_name="Arial",
                                       point_size=8)
                    pf_fixes.append(f"  PF#0     Arial 8  -> {obj.name}")
                except Exception as e:
                    pf_fixes.append(f"  PF#0     {obj.name} SKIP ({e})")

    for f in pf_fixes:
        print(f)
    print(f"\n  {len(pf_fixes)} objecten verwerkt")

    # =========================================================
    # Save
    # =========================================================
    print(f"\nSaving to {OUTPUT}...")
    rpt.save(TEMP)
    print("Done!")

# Move temp to final output
if os.path.exists(TEMP):
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)
    os.rename(TEMP, OUTPUT)
    print(f"Output: {OUTPUT} ({os.path.getsize(OUTPUT):,} bytes)")
