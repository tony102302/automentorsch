# ============================================================================
# Sheet Copy Script (pywin32)
# Creates a new sheet named Schematic2 under Schematic1 and copies components + nets
# ============================================================================
import csv
import json
import os
import win32com.client

VDM_COMP = 128
VDM_NET = 32
VDM_LABEL = 256
VDM_TEXT = 4
VDM_LINE = 1
VD_ALL = 0
VD_WIRE = 0
VDJ_LOW = 0
VDJ_HIGH = 1
VDLABELVISIBLE = 1
SHORT_NAME = 1


def get_active_app():
    try:
        return win32com.client.GetActiveObject("ViewDraw.Application")
    except Exception:
        return None


def stringlist_to_list(sl):
    items = []
    if sl is None:
        return items
    try:
        count = sl.GetCount()
        for i in range(1, count + 1):
            items.append(str(sl.GetItem(i)))
        return items
    except Exception:
        pass
    try:
        count = sl.Count
        for i in range(1, count + 1):
            items.append(str(sl.Item(i)))
    except Exception:
        return items
    return items


def iter_collection(coll):
    try:
        for obj in coll:
            yield obj
        return
    except Exception:
        pass
    try:
        count = coll.GetCount()
        for i in range(1, count + 1):
            yield coll.GetItem(i)
        return
    except Exception:
        pass
    try:
        count = coll.Count
        for i in range(1, count + 1):
            yield coll.Item(i)
    except Exception:
        return


def open_sheet(sheets, schematic_name, sheet_name):
    try:
        return sheets.Open(schematic_name, sheet_name)
    except Exception:
        pass
    try:
        if isinstance(sheet_name, str) and sheet_name.isdigit():
            return sheets.Open(schematic_name, int(sheet_name))
    except Exception:
        pass
    return None


def insert_sheet(sheets, schematic_name, sheet_name):
    ok = False
    try:
        ok = sheets.InsertSheet(schematic_name, sheet_name)
    except Exception:
        ok = False
    if not ok:
        try:
            if isinstance(sheet_name, str) and sheet_name.isdigit():
                ok = sheets.InsertSheet(schematic_name, int(sheet_name))
        except Exception:
            ok = False
    return ok


def get_view_from_doc(doc):
    try:
        views = doc.GetViews()
    except Exception:
        return None
    if views is None:
        return None
    try:
        return views.Item(1)
    except Exception:
        pass
    try:
        for v in views:
            return v
    except Exception:
        pass
    return None


def get_location(obj):
    try:
        return obj.GetLocation()
    except Exception:
        try:
            return obj.GetLocation
        except Exception:
            return None


def get_segments(net):
    try:
        return net.GetSegments()
    except Exception:
        try:
            return net.GetSegments
        except Exception:
            return None


def get_symbol_info(comp):
    part = ""
    sym_name = ""
    try:
        sym_block = comp.SymbolBlock
    except Exception:
        sym_block = None
    if sym_block is None:
        return part, sym_name
    try:
        part = str(sym_block.LibraryName).strip()
    except Exception:
        part = ""
    try:
        sym_name = str(sym_block.GetName(SHORT_NAME)).strip()
    except Exception:
        sym_name = ""
    return part, sym_name


def find_attribute(obj, name):
    try:
        return obj.FindAttribute(name)
    except Exception:
        pass
    try:
        attrs = obj.Attributes
    except Exception:
        return None
    try:
        return attrs.Item(name)
    except Exception:
        pass
    lname = str(name).lower()
    for attr in iter_collection(attrs):
        try:
            if str(attr.Name).lower() == lname:
                return attr
        except Exception:
            pass
    return None


def find_attributes_by_name(obj, name):
    attrs_list = []
    try:
        coll = obj.Attributes
    except Exception:
        return attrs_list
    lname = str(name).lower()
    for attr in iter_collection(coll):
        try:
            if str(attr.Name).lower() == lname:
                attrs_list.append(attr)
        except Exception:
            pass
    return attrs_list


def get_attribute(comp, name):
    return find_attribute(comp, name)


def get_attribute_value(attr):
    if attr is None:
        return ""
    for prop in ("EitherValue", "InstanceValue", "Value"):
        try:
            val = str(getattr(attr, prop)).strip()
            if val:
                return val
        except Exception:
            pass
    try:
        text = str(attr.TextString).strip()
    except Exception:
        text = ""
    if "=" in text:
        return text.split("=", 1)[1].strip()
    return ""


def set_attribute_value(attr, value):
    if attr is None:
        return False
    for prop in ("EitherValue", "InstanceValue", "Value"):
        try:
            setattr(attr, prop, value)
            return True
        except Exception:
            pass
    try:
        attr.TextString = f"{attr.Name}={value}"
        return True
    except Exception:
        return False


def attribute_value_from_data(data):
    for key in ("EitherValue", "InstanceValue", "Value"):
        val = data.get(key)
        if val is not None and str(val).strip():
            return str(val)
    text = str(data.get("TextString", "")).strip()
    if "=" in text:
        return text.split("=", 1)[1].strip()
    return ""


def get_component_value(comp):
    for name in ("Value", "VALUE"):
        attr = get_attribute(comp, name)
        if attr is None:
            continue
        val = get_attribute_value(attr)
        if val:
            return val, attr
    return "", None


def set_component_value(comp, value, src_attr=None):
    if not value:
        return False
    dst_attrs = []
    if src_attr is not None:
        try:
            dst_attrs = find_attributes_by_name(comp, src_attr.Name)
        except Exception:
            dst_attrs = []
    if not dst_attrs:
        dst_attrs = find_attributes_by_name(comp, "Value")
    if not dst_attrs:
        dst_attrs = find_attributes_by_name(comp, "VALUE")
    if not dst_attrs:
        try:
            comp.AddOat(f"Value={value}")
            return True
        except Exception:
            return False
    ok = False
    for dst_attr in dst_attrs:
        if set_attribute_value(dst_attr, value):
            ok = True
            if src_attr is not None:
                try:
                    dst_attr.Visible = src_attr.Visible
                except Exception:
                    pass
    return ok


def attribute_to_dict(attr):
    data = {}
    try:
        data["Name"] = str(attr.Name)
    except Exception:
        data["Name"] = ""
    for prop in ("Value", "EitherValue", "InstanceValue", "TextString"):
        try:
            data[prop] = str(getattr(attr, prop))
        except Exception:
            pass
    for prop in ("Visible", "NameVisible", "ValueVisible", "Orientation", "Size"):
        try:
            data[prop] = getattr(attr, prop)
        except Exception:
            pass
    try:
        origin = attr.Origin
        data["OriginX"] = int(origin.X)
        data["OriginY"] = int(origin.Y)
    except Exception:
        pass
    return data


def collect_attributes(obj):
    attrs = []
    try:
        coll = obj.Attributes
    except Exception:
        return attrs
    for attr in iter_collection(coll):
        data = attribute_to_dict(attr)
        if data.get("Name"):
            attrs.append(data)
    return attrs


def add_attribute(attrs_obj, data):
    name = str(data.get("Name", "")).strip()
    if not name:
        return None
    value = attribute_value_from_data(data)
    name_visible = bool(data.get("NameVisible", True))
    value_visible = bool(data.get("ValueVisible", True))
    try:
        return attrs_obj.Add(name, value, name_visible, value_visible, True)
    except Exception:
        return None


def apply_attributes(obj, attrs_data):
    if not attrs_data:
        return
    try:
        attrs_obj = obj.Attributes
    except Exception:
        attrs_obj = None
    for data in attrs_data:
        name = str(data.get("Name", "")).strip()
        if not name:
            continue
        attr = find_attribute(obj, name)
        if attr is None and attrs_obj is not None:
            attr = add_attribute(attrs_obj, data)
        if attr is None:
            continue
        value = attribute_value_from_data(data)
        if value:
            set_attribute_value(attr, value)
        for prop in ("Visible", "NameVisible", "ValueVisible", "Orientation", "Size"):
            if prop in data:
                try:
                    setattr(attr, prop, data[prop])
                except Exception:
                    pass
        if "OriginX" in data and "OriginY" in data:
            try:
                attr.SetLocation(int(data["OriginX"]), int(data["OriginY"]))
            except Exception:
                pass


def convert_oats(oats):
    lines = []
    for line in oats.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(None, 1)
        if len(parts) == 2 and parts[0].isdigit():
            vis = parts[0]
            rest = parts[1]
        else:
            vis = "1"
            rest = line
        lines.append(f"{vis} 1 {rest}\r")
    return "".join(lines)


def copy_components(src_view, dst_block):
    comps = src_view.Query(VDM_COMP, VD_ALL)
    for comp in comps:
        loc = get_location(comp)
        if loc is None:
            continue
        part, sym_name = get_symbol_info(comp)
        if not part or not sym_name:
            continue
        new_comp = dst_block.AddSymbolInstance(part, sym_name, int(loc.X), int(loc.Y))
        if new_comp is None:
            continue
        try:
            new_comp.Orientation = comp.Orientation
        except Exception:
            pass
        try:
            new_comp.Scale = comp.Scale
        except Exception:
            pass
        try:
            new_comp.Refdes = comp.Refdes
        except Exception:
            pass
        try:
            oats = comp.GetBatchOats()
            if oats:
                new_comp.AddBatchOats(convert_oats(oats))
        except Exception:
            pass
        try:
            value, src_attr = get_component_value(comp)
            set_component_value(new_comp, value, src_attr)
        except Exception:
            pass
        try:
            attrs_data = collect_attributes(comp)
            apply_attributes(new_comp, attrs_data)
        except Exception:
            pass


def segment_midpoint(seg):
    try:
        p_low = seg.Location(VDJ_LOW)
        p_high = seg.Location(VDJ_HIGH)
        return int((p_low.X + p_high.X) / 2), int((p_low.Y + p_high.Y) / 2)
    except Exception:
        return None, None


def segment_key(x1, y1, x2, y2):
    if (x1, y1) <= (x2, y2):
        return (x1, y1, x2, y2)
    return (x2, y2, x1, y1)


def point_on_segment(x, y, x1, y1, x2, y2, tol=1):
    if x1 == x2:
        return abs(x - x1) <= tol and min(y1, y2) - tol <= y <= max(y1, y2) + tol
    if y1 == y2:
        return abs(y - y1) <= tol and min(x1, x2) - tol <= x <= max(x1, x2) + tol
    return point_to_segment_distance(x, y, x1, y1, x2, y2) <= tol


def try_add_label(net, seg, name, x, y, orient=None, size=None):
    try:
        lbl = net.AddLabel(seg, name, int(x), int(y))
    except Exception:
        return False
    if lbl is None:
        return False
    if orient is not None:
        try:
            lbl.Orientation = orient
        except Exception:
            pass
    if size is not None:
        try:
            lbl.Size = size
        except Exception:
            pass
    try:
        lbl.SetLocation(int(x), int(y))
    except Exception:
        pass
    try:
        lbl.TextString = name
    except Exception:
        pass
    try:
        lbl.Visible = VDLABELVISIBLE
    except Exception:
        pass
    return True


def net_has_label(net, name):
    segs = get_segments(net)
    if segs is None:
        return False
    for seg in iter_collection(segs):
        lbl = None
        try:
            lbl = net.GetLabel(seg)
        except Exception:
            lbl = None
        if lbl is None:
            try:
                lbl = net.GetConnectedLabel(seg)
            except Exception:
                lbl = None
        if lbl is not None:
            if _label_text_from_label(lbl) == name:
                return True
    return False


def _label_text_from_label(lbl):
    text = ""
    try:
        text = str(lbl.TextString).strip()
    except Exception:
        text = ""
    if not text or text.startswith("$"):
        try:
            text = str(lbl.ResolvedName).strip()
        except Exception:
            pass
    return text


def get_net_labels(net):
    segs = get_segments(net)
    if segs is None:
        return []

    labels = []
    seen = set()

    for seg in iter_collection(segs):
        try:
            p_low = seg.Location(VDJ_LOW)
            p_high = seg.Location(VDJ_HIGH)
        except Exception:
            continue

        lbl = None
        try:
            lbl = net.GetLabel(seg)
        except Exception:
            lbl = None

        if lbl is not None:
            name = _label_text_from_label(lbl)
            loc = get_location(lbl)
            if not name or name.startswith("$") or loc is None:
                continue
            try:
                orient = lbl.Orientation
            except Exception:
                orient = None
            try:
                size = lbl.Size
            except Exception:
                size = None
            key = (name, int(loc.X), int(loc.Y))
            if key in seen:
                continue
            seen.add(key)
            labels.append(
                (
                    name,
                    int(loc.X),
                    int(loc.Y),
                    p_low.X,
                    p_low.Y,
                    p_high.X,
                    p_high.Y,
                    orient,
                    size,
                )
            )

    if labels:
        return labels

    for seg in iter_collection(segs):
        try:
            p_low = seg.Location(VDJ_LOW)
            p_high = seg.Location(VDJ_HIGH)
        except Exception:
            continue
        lbl = None
        try:
            lbl = net.GetConnectedLabel(seg)
        except Exception:
            lbl = None
        if lbl is None:
            continue
        name = _label_text_from_label(lbl)
        loc = get_location(lbl)
        if not name or name.startswith("$") or loc is None:
            continue
        if not point_on_segment(loc.X, loc.Y, p_low.X, p_low.Y, p_high.X, p_high.Y):
            continue
        try:
            orient = lbl.Orientation
        except Exception:
            orient = None
        try:
            size = lbl.Size
        except Exception:
            size = None
        key = (name, int(loc.X), int(loc.Y))
        if key in seen:
            continue
        seen.add(key)
        labels.append(
            (
                name,
                int(loc.X),
                int(loc.Y),
                p_low.X,
                p_low.Y,
                p_high.X,
                p_high.Y,
                orient,
                size,
            )
        )

    return labels


def copy_nets(src_view, dst_block):
    nets = src_view.Query(VDM_NET, VD_ALL)
    labels_added = 0
    for net in iter_collection(nets):
        segs = get_segments(net)
        if segs is None:
            continue

        seg_list = []
        for seg in iter_collection(segs):
            try:
                p_low = seg.Location(VDJ_LOW)
                p_high = seg.Location(VDJ_HIGH)
                seg_list.append((p_low.X, p_low.Y, p_high.X, p_high.Y))
            except Exception:
                continue

        labels = get_net_labels(net)
        last_net = None
        for x1, y1, x2, y2 in seg_list:
            try:
                last_net = dst_block.AddNet(
                    int(x1), int(y1), int(x2), int(y2), None, None, VD_WIRE
                )
            except Exception:
                pass

        try:
            attrs_data = collect_attributes(net)
            if last_net is not None:
                apply_attributes(last_net, attrs_data)
        except Exception:
            pass

        if labels and last_net is not None:
            try:
                dst_segs = get_segments(last_net)
                if dst_segs is None:
                    continue
                dst_seg_list = []
                for s in iter_collection(dst_segs):
                    try:
                        p_low = s.Location(VDJ_LOW)
                        p_high = s.Location(VDJ_HIGH)
                    except Exception:
                        continue
                    dst_seg_list.append(
                        (
                            s,
                            p_low.X,
                            p_low.Y,
                            p_high.X,
                            p_high.Y,
                            segment_key(p_low.X, p_low.Y, p_high.X, p_high.Y),
                        )
                    )

                for name, lx, ly, sx1, sy1, sx2, sy2, orient, size in labels:
                    src_key = segment_key(sx1, sy1, sx2, sy2)
                    chosen = None
                    for s, x1, y1, x2, y2, key in dst_seg_list:
                        if key == src_key:
                            chosen = s
                            break
                    if chosen is None:
                        for s, x1, y1, x2, y2, _ in dst_seg_list:
                            if point_on_segment(lx, ly, x1, y1, x2, y2, tol=1):
                                chosen = s
                                break
                    if chosen is None:
                        best = None
                        best_dist = None
                        for s, x1, y1, x2, y2, _ in dst_seg_list:
                            dist = point_to_segment_distance(lx, ly, x1, y1, x2, y2)
                            if best_dist is None or dist < best_dist:
                                best_dist = dist
                                best = s
                        chosen = best
                    if chosen is not None and not net_has_label(last_net, name):
                        if try_add_label(last_net, chosen, name, lx, ly, orient, size):
                            labels_added += 1
            except Exception:
                pass
    return labels_added


def count_collection(coll):
    try:
        return int(coll.Count)
    except Exception:
        pass
    try:
        return int(coll.GetCount())
    except Exception:
        pass
    count = 0
    try:
        for _ in coll:
            count += 1
    except Exception:
        return 0
    return count


def get_parent(obj):
    try:
        return obj.Parent
    except Exception:
        return None


def resolve_net_from_label(lbl):
    parent = get_parent(lbl)
    if parent is None:
        return None
    if get_segments(parent) is not None:
        return parent
    parent2 = get_parent(parent)
    if parent2 is not None and get_segments(parent2) is not None:
        return parent2
    return None


def point_to_segment_distance(x, y, x1, y1, x2, y2):
    if x1 == x2:
        if min(y1, y2) <= y <= max(y1, y2):
            return abs(x - x1)
        return min(((x - x1) ** 2 + (y - y1) ** 2) ** 0.5,
                   ((x - x2) ** 2 + (y - y2) ** 2) ** 0.5)
    if y1 == y2:
        if min(x1, x2) <= x <= max(x1, x2):
            return abs(y - y1)
        return min(((x - x1) ** 2 + (y - y1) ** 2) ** 0.5,
                   ((x - x2) ** 2 + (y - y2) ** 2) ** 0.5)
    # general fallback
    return min(((x - x1) ** 2 + (y - y1) ** 2) ** 0.5,
               ((x - x2) ** 2 + (y - y2) ** 2) ** 0.5)




def delete_empty_sheet(sheets, schematic_name, sheet_name):
    doc = open_sheet(sheets, schematic_name, sheet_name)
    if doc is None:
        return False
    view = get_view_from_doc(doc)
    if view is None:
        return False
    comps = view.Query(VDM_COMP, VD_ALL)
    nets = view.Query(VDM_NET, VD_ALL)
    if count_collection(comps) == 0 and count_collection(nets) == 0:
        try:
            return bool(sheets.DeleteSheet(schematic_name, sheet_name))
        except Exception:
            try:
                if isinstance(sheet_name, str) and sheet_name.isdigit():
                    return bool(sheets.DeleteSheet(schematic_name, int(sheet_name)))
            except Exception:
                pass
    return False


def clear_sheet(view):
    block = view.Block
    if block is None:
        return
    try:
        block.DeSelectAll()
    except Exception:
        pass
    mask = VDM_COMP | VDM_NET | VDM_LABEL | VDM_TEXT | VDM_LINE
    objs = view.Query(mask, VD_ALL)
    for obj in iter_collection(objs):
        try:
            obj.Selected = True
        except Exception:
            pass
    try:
        block.DeleteSelected()
    except Exception:
        pass


def get_sheet_object_count(view):
    comps = view.Query(VDM_COMP, VD_ALL)
    nets = view.Query(VDM_NET, VD_ALL)
    return count_collection(comps) + count_collection(nets)


def choose_source_sheet(sheets, schematic_name, preferred=None):
    try:
        sheet_list = stringlist_to_list(sheets.GetAvailableSheets(schematic_name))
    except Exception:
        sheet_list = []
    if not sheet_list:
        return None
    if preferred and preferred in sheet_list:
        return preferred
    best_sheet = sheet_list[0]
    best_count = -1
    for name in sheet_list:
        doc = open_sheet(sheets, schematic_name, name)
        if doc is None:
            continue
        view = get_view_from_doc(doc)
        if view is None:
            continue
        count = get_sheet_object_count(view)
        if count > best_count:
            best_count = count
            best_sheet = name
    return best_sheet


def _open_csv_writer(path, fieldnames):
    try:
        f = open(path, "w", newline="", encoding="utf-8")
        return f, path
    except PermissionError:
        base, ext = os.path.splitext(path)
        fallback = f"{base}_tmp{ext}"
        f = open(fallback, "w", newline="", encoding="utf-8")
        return f, fallback


def export_components(view, path):
    rows = []
    comps = view.Query(VDM_COMP, VD_ALL)
    for comp in iter_collection(comps):
        loc = get_location(comp)
        if loc is None:
            continue
        part, sym_name = get_symbol_info(comp)
        if not part or not sym_name:
            continue
        row = {
            "Refdes": getattr(comp, "Refdes", ""),
            "Partition": part,
            "Symbol": sym_name,
            "X": int(loc.X),
            "Y": int(loc.Y),
        }
        try:
            row["Orientation"] = comp.Orientation
        except Exception:
            row["Orientation"] = ""
        try:
            row["Scale"] = comp.Scale
        except Exception:
            row["Scale"] = ""
        attrs = collect_attributes(comp)
        row["Attributes"] = json.dumps(attrs, ensure_ascii=False)
        rows.append(row)

    f, used_path = _open_csv_writer(path, fieldnames=None)
    with f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "Refdes",
                "Partition",
                "Symbol",
                "X",
                "Y",
                "Orientation",
                "Scale",
                "Attributes",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)
    return used_path


def export_nets(view, path):
    rows = []
    nets = view.Query(VDM_NET, VD_ALL)
    for net in iter_collection(nets):
        segs = get_segments(net)
        if segs is None:
            continue
        seg_list = []
        for seg in iter_collection(segs):
            try:
                p_low = seg.Location(VDJ_LOW)
                p_high = seg.Location(VDJ_HIGH)
                seg_list.append((p_low.X, p_low.Y, p_high.X, p_high.Y))
            except Exception:
                continue
        labels = []
        for name, lx, ly, sx1, sy1, sx2, sy2, orient, size in get_net_labels(net):
            labels.append(
                {
                    "Name": name,
                    "X": int(lx),
                    "Y": int(ly),
                    "SegX1": int(sx1),
                    "SegY1": int(sy1),
                    "SegX2": int(sx2),
                    "SegY2": int(sy2),
                    "Orientation": orient,
                    "Size": size,
                }
            )
        attrs = collect_attributes(net)
        rows.append(
            {
                "Segments": json.dumps(seg_list, ensure_ascii=False),
                "Labels": json.dumps(labels, ensure_ascii=False),
                "Attributes": json.dumps(attrs, ensure_ascii=False),
            }
        )

    f, used_path = _open_csv_writer(path, fieldnames=None)
    with f:
        writer = csv.DictWriter(f, fieldnames=["Segments", "Labels", "Attributes"])
        writer.writeheader()
        writer.writerows(rows)
    return used_path


def import_components(path, dst_block):
    with open(path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            part = row.get("Partition", "")
            sym = row.get("Symbol", "")
            if not part or not sym:
                continue
            try:
                x = int(float(row.get("X", "0")))
                y = int(float(row.get("Y", "0")))
            except Exception:
                x, y = 0, 0
            new_comp = dst_block.AddSymbolInstance(part, sym, x, y)
            if new_comp is None:
                continue
            refdes = row.get("Refdes", "")
            if refdes:
                try:
                    new_comp.Refdes = refdes
                except Exception:
                    pass
            try:
                ori = row.get("Orientation", "")
                if ori != "":
                    new_comp.Orientation = int(float(ori))
            except Exception:
                pass
            try:
                scale = row.get("Scale", "")
                if scale != "":
                    new_comp.Scale = float(scale)
            except Exception:
                pass
            attrs_text = row.get("Attributes", "")
            if attrs_text:
                try:
                    attrs_data = json.loads(attrs_text)
                    apply_attributes(new_comp, attrs_data)
                    value = ""
                    for attr in attrs_data:
                        try:
                            if str(attr.get("Name", "")).lower() == "value":
                                value = attribute_value_from_data(attr)
                                break
                        except Exception:
                            pass
                    if value:
                        set_component_value(new_comp, value)
                except Exception:
                    pass


def import_nets(path, dst_block):
    with open(path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                seg_list = json.loads(row.get("Segments", "[]"))
            except Exception:
                seg_list = []
            if not seg_list:
                continue
            last_net = None
            for x1, y1, x2, y2 in seg_list:
                try:
                    last_net = dst_block.AddNet(
                        int(x1), int(y1), int(x2), int(y2), None, None, VD_WIRE
                    )
                except Exception:
                    pass
            if last_net is None:
                continue
            attrs_text = row.get("Attributes", "")
            if attrs_text:
                try:
                    attrs_data = json.loads(attrs_text)
                    apply_attributes(last_net, attrs_data)
                except Exception:
                    pass
            labels_text = row.get("Labels", "")
            if labels_text:
                try:
                    labels = json.loads(labels_text)
                except Exception:
                    labels = []
                dst_segs = get_segments(last_net)
                if dst_segs is None:
                    continue
                dst_seg_list = []
                for s in iter_collection(dst_segs):
                    try:
                        p_low = s.Location(VDJ_LOW)
                        p_high = s.Location(VDJ_HIGH)
                    except Exception:
                        continue
                    dst_seg_list.append(
                        (
                            s,
                            p_low.X,
                            p_low.Y,
                            p_high.X,
                            p_high.Y,
                            segment_key(p_low.X, p_low.Y, p_high.X, p_high.Y),
                        )
                    )
                for lbl in labels:
                    name = str(lbl.get("Name", "")).strip()
                    if not name:
                        continue
                    lx = int(lbl.get("X", 0))
                    ly = int(lbl.get("Y", 0))
                    sx1 = int(lbl.get("SegX1", 0))
                    sy1 = int(lbl.get("SegY1", 0))
                    sx2 = int(lbl.get("SegX2", 0))
                    sy2 = int(lbl.get("SegY2", 0))
                    orient = lbl.get("Orientation")
                    size = lbl.get("Size")
                    src_key = segment_key(sx1, sy1, sx2, sy2)
                    chosen = None
                    for s, x1, y1, x2, y2, key in dst_seg_list:
                        if key == src_key:
                            chosen = s
                            break
                    if chosen is None:
                        for s, x1, y1, x2, y2, _ in dst_seg_list:
                            if point_on_segment(lx, ly, x1, y1, x2, y2, tol=1):
                                chosen = s
                                break
                    if chosen is None:
                        best = None
                        best_dist = None
                        for s, x1, y1, x2, y2, _ in dst_seg_list:
                            dist = point_to_segment_distance(lx, ly, x1, y1, x2, y2)
                            if best_dist is None or dist < best_dist:
                                best_dist = dist
                                best = s
                        chosen = best
                    if chosen is not None and not net_has_label(last_net, name):
                        try_add_label(last_net, chosen, name, lx, ly, orient, size)


def main():
    app = get_active_app()
    if app is None:
        print("Please open Xpedition Designer and a schematic page first.")
        return

    schematic_name = "Schematic1"
    dst_sheet = "Schematic2"

    try:
        sheets = app.SchematicSheetDocuments()
    except Exception:
        sheets = app.SchematicSheetDocuments

    schems = stringlist_to_list(sheets.GetAvailableSchematics())
    if not any(s.lower() == schematic_name.lower() for s in schems):
        if schems:
            schematic_name = schems[0]
        else:
            print("Cannot find any schematic.")
            return

    src_sheet = choose_source_sheet(sheets, schematic_name, preferred="2")
    if not src_sheet:
        print("Cannot find source sheet.")
        return

    insert_sheet(sheets, schematic_name, dst_sheet)

    # Source view
    src_doc = open_sheet(sheets, schematic_name, src_sheet)
    if src_doc is None:
        print("Cannot open source sheet.")
        return
    src_view = get_view_from_doc(src_doc)
    if src_view is None:
        print("Cannot resolve source view.")
        return

    # Destination view (same schematic, new sheet)
    dst_doc = open_sheet(sheets, schematic_name, dst_sheet)
    if dst_doc is None:
        print("Cannot open destination sheet.")
        return

    dst_view = get_view_from_doc(dst_doc)
    if dst_view is None:
        print("Cannot get destination view.")
        return

    dst_block = dst_view.Block
    if dst_block is None:
        print("Cannot access destination block.")
        return

    base_dir = os.path.dirname(os.path.abspath(__file__))
    parts_csv = os.path.join(base_dir, "parts.csv")
    nets_csv = os.path.join(base_dir, "net.csv")

    app.SetRedraw(False)
    try:
        parts_csv_used = export_components(src_view, parts_csv)
        nets_csv_used = export_nets(src_view, nets_csv)
        clear_sheet(dst_view)
        import_components(parts_csv_used, dst_block)
        import_nets(nets_csv_used, dst_block)
    finally:
        app.SetRedraw(True)

    try:
        dst_view.Refresh()
    except Exception:
        pass

    msg = f"{dst_sheet} created and copied from {schematic_name}:{src_sheet} via CSV."
    if parts_csv_used != parts_csv or nets_csv_used != nets_csv:
        msg += f" (fallback CSV: {os.path.basename(parts_csv_used)}, {os.path.basename(nets_csv_used)})"
    print(msg)


if __name__ == "__main__":
    main()
