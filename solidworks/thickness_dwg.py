import re
from pathlib import Path

DXF_FOLDER_NAME = "_parts_dxf"


class ThicknessDwgBuilder:
    """
    Puts the flat patterns of all sheet metal parts of the same thickness
    into ONE drawing file (one file per thickness).

    Steps:
      1. every part's flat pattern is exported by SolidWorks to a temporary DXF
      2. the DXFs of one thickness are merged (arranged in rows) with ezdxf
      3. the merged DXF is converted to DWG (needs ODA File Converter);
         if the converter is not installed the merged DXF is kept instead
    """

    def __init__(
        self,
        exporter,
        gap=20.0,
        max_row_width=6000.0,
        repeat_by_quantity=False,
        oda_path=None,
        dwg_version="R2018",
    ):
        self.exporter = exporter
        self.gap = gap
        self.max_row_width = max_row_width
        self.repeat_by_quantity = repeat_by_quantity
        self.oda_path = oda_path
        self.dwg_version = dwg_version

    # =====================================================
    # HELPERS
    # =====================================================

    @staticmethod
    def _safe(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def _format_thickness(value):
        return f"{value:g}"

    # =====================================================
    # STEP 1: EXPORT EVERY PART TO DXF
    # =====================================================

    def _export_part_dxfs(self, parts, dxf_folder):
        dxf_folder.mkdir(parents=True, exist_ok=True)

        total = len(parts)
        exported = []
        failed = []
        used = set()

        print()
        print("=" * 60)
        print("STEP 1/2: EXPORT FLAT PATTERNS (DXF)")
        print("=" * 60)

        for index, item in enumerate(parts, start=1):
            name = item["name"]
            label = f"[{index}/{total}] {name}"

            model = item.get("model")
            if model is None:
                print(f"{label} -> FAILED: no model")
                failed.append((name, "no model"))
                continue

            base = self._safe(name)
            unique = base
            counter = 2
            while unique.lower() in used:
                unique = f"{base}_{counter}"
                counter += 1
            used.add(unique.lower())

            dxf_path = dxf_folder / f"{unique}.dxf"

            try:
                ok, message = self.exporter.export_part(model, dxf_path)
            except Exception as e:
                ok, message = False, f"error: {e}"

            if ok:
                print(f"{label} -> OK")
                exported.append({"item": item, "dxf": dxf_path})
            else:
                print(f"{label} -> FAILED: {message}")
                failed.append((name, message))

        return exported, failed

    # =====================================================
    # STEP 2: MERGE ONE THICKNESS INTO ONE DRAWING
    # =====================================================

    def _merge(self, entries, out_dxf):
        import ezdxf
        from ezdxf import bbox
        from ezdxf.addons.importer import Importer

        target = ezdxf.new("R2010")
        target_msp = target.modelspace()

        if "PART_LABELS" not in target.layers:
            target.layers.add("PART_LABELS", color=3)

        cursor_x = 0.0
        cursor_y = 0.0
        row_height = 0.0
        placed = 0
        skipped = []
        units_set = False

        for entry in entries:
            item = entry["item"]
            copies = max(1, int(item["quantity"])) if self.repeat_by_quantity else 1

            for copy_index in range(copies):
                source = ezdxf.readfile(str(entry["dxf"]))
                source_msp = source.modelspace()

                if not units_set:
                    try:
                        target.units = source.units
                    except Exception:
                        pass
                    units_set = True

                box = bbox.extents(source_msp)
                if not box.has_data:
                    skipped.append(item["name"])
                    break

                width = box.size.x
                height = box.size.y
                label_space = max(10.0, self.gap)

                if cursor_x > 0 and cursor_x + width > self.max_row_width:
                    cursor_x = 0.0
                    cursor_y += row_height + self.gap
                    row_height = 0.0

                dx = cursor_x - box.extmin.x
                dy = cursor_y - box.extmin.y

                entities = list(source_msp)
                for entity in entities:
                    entity.translate(dx, dy, 0)

                importer = Importer(source, target)
                importer.import_entities(entities, target_msp)
                importer.finalize()

                label = item["name"]
                if self.repeat_by_quantity and copies > 1:
                    label += f"  ({copy_index + 1}/{copies})"
                elif not self.repeat_by_quantity:
                    label += f"  x{item['quantity']}"

                target_msp.add_text(
                    label,
                    height=max(4.0, label_space * 0.5),
                    dxfattribs={
                        "layer": "PART_LABELS",
                        "insert": (cursor_x, cursor_y + height + 2.0),
                    },
                )

                row_height = max(row_height, height + label_space)
                cursor_x += width + self.gap
                placed += 1

        target.saveas(str(out_dxf))
        return placed, skipped

    # =====================================================
    # STEP 3: DXF -> DWG (optional)
    # =====================================================

    def _convert_to_dwg(self, dxf_path):
        try:
            import ezdxf
            from ezdxf.addons import odafc

            if self.oda_path:
                try:
                    ezdxf.options.set(
                        "odafc-addon", "win_exec_path", str(self.oda_path)
                    )
                except Exception:
                    pass

            if not odafc.is_installed():
                return None, "ODA File Converter not found"

            dwg_path = Path(dxf_path).with_suffix(".dwg")
            odafc.convert(
                str(dxf_path), str(dwg_path), version=self.dwg_version, replace=True
            )

            if dwg_path.exists() and dwg_path.stat().st_size > 0:
                return dwg_path, "OK"

            return None, "converter produced no file"

        except Exception as e:
            return None, f"conversion error: {e}"

    # =====================================================
    # BUILD
    # =====================================================

    def build(self, sheet_metal_parts, output_folder, base_name="sheet_metal"):
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        try:
            import ezdxf  # noqa: F401
        except ImportError:
            print()
            print("The 'ezdxf' package is required to merge drawings.")
            print("Install it with:  pip install ezdxf")
            return {"files": [], "failed_parts": [], "error": "ezdxf missing"}

        parts = sorted(
            [p for p in sheet_metal_parts if p.get("thickness") is not None],
            key=lambda p: (p["thickness"], p["name"]),
        )

        exported, failed = self._export_part_dxfs(
            parts, output_folder / DXF_FOLDER_NAME
        )

        groups = {}
        for entry in exported:
            groups.setdefault(entry["item"]["thickness"], []).append(entry)

        print()
        print("=" * 60)
        print("STEP 2/2: ONE DRAWING PER THICKNESS")
        print("=" * 60)

        files = []

        for thickness in sorted(groups):
            entries = groups[thickness]
            tag = self._format_thickness(thickness)
            out_dxf = output_folder / f"{self._safe(base_name)}_{tag}mm.dxf"

            quantity = sum(int(e["item"]["quantity"]) for e in entries)

            try:
                placed, skipped = self._merge(entries, out_dxf)
            except Exception as e:
                print(f"{tag} mm -> MERGE FAILED: {e}")
                failed.extend(
                    (e2["item"]["name"], f"merge failed: {e}") for e2 in entries
                )
                continue

            dwg_path, message = self._convert_to_dwg(out_dxf)
            final_path = dwg_path if dwg_path else out_dxf

            if dwg_path:
                print(
                    f"{tag} mm -> {len(entries)} parts "
                    f"(total quantity {quantity}) -> {dwg_path.name}"
                )
            else:
                print(
                    f"{tag} mm -> {len(entries)} parts "
                    f"(total quantity {quantity}) -> {out_dxf.name} "
                    f"(DXF only: {message})"
                )

            if skipped:
                print(f"    empty drawings skipped: {', '.join(skipped)}")

            files.append(
                {
                    "thickness": thickness,
                    "unique_parts": len(entries),
                    "quantity": quantity,
                    "path": final_path,
                    "is_dwg": dwg_path is not None,
                }
            )

        print()
        print("=" * 60)
        print("DRAWING SUMMARY")
        print("=" * 60)
        for f in files:
            kind = "DWG" if f["is_dwg"] else "DXF"
            print(
                f"{f['thickness']:>8g} mm | parts: {f['unique_parts']:<3} | "
                f"qty: {f['quantity']:<4} | {kind}: {f['path'].name}"
            )

        if any(not f["is_dwg"] for f in files):
            print()
            print("Some files are DXF because ODA File Converter was not found.")
            print("Install the free 'ODA File Converter' and run again to get DWG,")
            print("or open the DXF in AutoCAD and use Save As -> DWG.")

        if failed:
            print()
            print("Parts that could not be exported:")
            for name, message in failed:
                print(f"  - {name}: {message}")

        return {"files": files, "failed_parts": failed, "error": None}
