# ============================================================================
# Voltage Divider Example (pywin32)
# Based on draw_voltage_divider.vbs
# Creates: 5V -> R1(pin1) -> net2.5v -> R2(pin1) -> GND (R2 pin2)
# Partition: Discrete, Device: R0603, Symbol: RES.1
# ============================================================================
import win32com.client

VD_WIRE = 0
VDLOWERLEFT = 0
VDUPPERRIGHT = 3
VDJ_LOW = 0
VDJ_HIGH = 1
VDLABELINVISIBLE = 0
VDLABELVISIBLE = 1


def get_active_app():
    try:
        return win32com.client.GetActiveObject("ViewDraw.Application")
    except Exception:
        return None


def get_connections(comp):
    try:
        return comp.GetConnections()
    except Exception:
        return comp.GetConnections


def get_pin_number(pin):
    try:
        return str(pin.Number).strip()
    except Exception:
        try:
            return str(pin.Number()).strip()
        except Exception:
            return ""


def find_pin_by_number(comp, number_str):
    conns = get_connections(comp)
    # Try iteration first
    try:
        for conn in conns:
            pin = conn.CompPin
            if get_pin_number(pin) == number_str:
                return pin
    except Exception:
        pass
    # Fallback to indexed access
    try:
        count = conns.Count
        for i in range(1, count + 1):
            conn = conns.Item(i)
            pin = conn.CompPin
            if get_pin_number(pin) == number_str:
                return pin
    except Exception:
        pass
    return None


def get_pin_location(pin):
    try:
        return pin.GetLocation()
    except Exception:
        return pin.GetLocation


def get_two_pins_by_location(comp):
    conns = get_connections(comp)
    pins = []
    try:
        for conn in conns:
            pins.append(conn.CompPin)
            if len(pins) >= 2:
                break
    except Exception:
        pass
    if len(pins) < 2:
        try:
            count = conns.Count
            for i in range(1, count + 1):
                pins.append(conns.Item(i).CompPin)
                if len(pins) >= 2:
                    break
        except Exception:
            pass
    if len(pins) < 2:
        return None, None

    p1, p2 = pins[0], pins[1]
    loc1 = get_pin_location(p1)
    loc2 = get_pin_location(p2)
    # Deterministic order: higher Y first
    if loc1.Y >= loc2.Y:
        return p1, p2
    return p2, p1


def get_attr_value(attr):
    try:
        return str(attr.Value).strip()
    except Exception:
        try:
            return str(attr.Value()).strip()
        except Exception:
            try:
                return str(attr.InstanceValue).strip()
            except Exception:
                try:
                    return str(attr.InstanceValue()).strip()
                except Exception:
                    return ""


def hide_device_attribute(comp, required_value=None):
    try:
        attr = comp.FindAttribute("DEVICE")
        if attr is None:
            return False
        if required_value is not None:
            if get_attr_value(attr) != required_value:
                return False
        try:
            attr.Visibility = 0
        except Exception:
            try:
                attr.Visible = 0
            except Exception:
                try:
                    attr.NameVisible = 0
                    attr.ValueVisible = 0
                except Exception:
                    pass
        return True
    except Exception:
        return False


def set_component_attribute(comp, name, value):
    try:
        attr = comp.FindAttribute(name)
    except Exception:
        attr = None
    if attr is not None:
        try:
            attr.Value = value
            return True
        except Exception:
            pass
        try:
            attr.InstanceValue = value
            return True
        except Exception:
            pass
        try:
            attr.TextString = value
            return True
        except Exception:
            pass
    try:
        comp.AddOat(f"{name}={value}")
        return True
    except Exception:
        return False


def get_attr_name(attr):
    try:
        return str(attr.Name).strip()
    except Exception:
        try:
            return str(attr.Name()).strip()
        except Exception:
            return ""


def normalize_value_attribute(comp, value):
    attrs = []
    try:
        for attr in comp.Attributes:
            if get_attr_name(attr).lower() == "value":
                attrs.append(attr)
    except Exception:
        try:
            count = comp.Attributes.Count
            for i in range(1, count + 1):
                attr = comp.Attributes.Item(i)
                if get_attr_name(attr).lower() == "value":
                    attrs.append(attr)
        except Exception:
            attrs = []

    if not attrs:
        set_component_attribute(comp, "Value", value)
        return

    for attr in attrs:
        try:
            attr.Value = value
        except Exception:
            try:
                attr.InstanceValue = value
            except Exception:
                try:
                    attr.TextString = value
                except Exception:
                    pass

    for attr in attrs:
        if get_attr_name(attr) != "Value":
            try:
                attr.Delete()
            except Exception:
                try:
                    attr.Visible = 0
                except Exception:
                    pass


def iter_collection(coll):
    try:
        for obj in coll:
            yield obj
        return
    except Exception:
        pass
    try:
        count = coll.Count
        for i in range(1, count + 1):
            yield coll.Item(i)
    except Exception:
        return


def hide_device_r0603_in_design(app):
    try:
        design_name = app.GetActiveDesign()
    except Exception:
        design_name = ""
    if not design_name:
        return 0
    try:
        comps = app.DesignComponents("", design_name, -1, "", True)
    except Exception:
        return 0
    changed = 0
    for comp in iter_collection(comps):
        if hide_device_attribute(comp, "R0603"):
            changed += 1
    return changed


def add_net_with_label(block, x1, y1, x2, y2, pin1, pin2, name, label_x, label_y):
    def segment_midpoint(seg):
        try:
            p_low = seg.Location(VDJ_LOW)
            p_high = seg.Location(VDJ_HIGH)
            return int((p_low.X + p_high.X) / 2), int((p_low.Y + p_high.Y) / 2)
        except Exception:
            return None, None

    def find_label_on_net(net_obj):
        try:
            segs = net_obj.GetSegments()
            for i in range(1, segs.Count + 1):
                seg = segs.Item(i)
                lbl = net_obj.GetLabel(seg)
                if lbl is not None:
                    return lbl
        except Exception:
            return None
        return None

    def ensure_label_visible(lbl_obj, text=None):
        if lbl_obj is None:
            return
        try:
            if text:
                lbl_obj.TextString = text
        except Exception:
            pass
        try:
            lbl_obj.Visible = VDLABELVISIBLE
        except Exception:
            try:
                lbl_obj.Visible = True
            except Exception:
                pass

    def get_net_name(net_obj):
        if net_obj is None:
            return ""
        try:
            segs = net_obj.GetSegments()
            if segs.Count <= 0:
                return ""
            seg = segs.Item(1)
        except Exception:
            return ""
        try:
            lbl = net_obj.GetLabel(seg)
            if lbl is not None:
                text = str(lbl.TextString).strip()
                if text:
                    return text
        except Exception:
            pass
        try:
            text = net_obj.GetConnectedNetName(seg)
            if text:
                text = str(text).strip()
                if text.startswith("$"):
                    return ""
                return text
        except Exception:
            pass
        return ""

    def get_net_from_pin(pin):
        if pin is None:
            return None
        try:
            conn = pin.Connection
            if conn is None:
                return None
            return conn.Net
        except Exception:
            try:
                return pin.Connection.Net
            except Exception:
                return None

    try:
        if int(x1) == int(x2) or int(y1) == int(y2):
            net = block.AddNet(int(x1), int(y1), int(x2), int(y2), pin1, pin2, VD_WIRE)
        else:
            elbow_x, elbow_y = int(x1), int(y2)
            net1 = block.AddNet(int(x1), int(y1), elbow_x, elbow_y, pin1, None, VD_WIRE)
            net2 = block.AddNet(elbow_x, elbow_y, int(x2), int(y2), None, pin2, VD_WIRE)
            net = net2 if net2 is not None else net1
    except Exception:
        net = None
    if net is None:
        net = get_net_from_pin(pin1)
    if net is None:
        net = get_net_from_pin(pin2)
    if net is None:
        return

    existing_name = get_net_name(net)
    existing_label = find_label_on_net(net)
    if existing_label is not None:
        ensure_label_visible(existing_label, existing_name or name)
    if existing_name:
        print(f"Net already named: {existing_name}")
        return
    try:
        segments = net.GetSegments()
        if segments.Count <= 0:
            return

        def try_add(seg, lx, ly):
            try:
                lbl = net.AddLabel(seg, name, int(lx), int(ly))
                if lbl is not None:
                    try:
                        lbl.TextString = name
                    except Exception:
                        pass
                    try:
                        lbl.Visible = VDLABELVISIBLE
                    except Exception:
                        pass
            except Exception:
                return False
            return bool(get_net_name(net))

        for i in range(1, segments.Count + 1):
            seg = segments.Item(i)
            mx, my = segment_midpoint(seg)
            if mx is None:
                continue
            if try_add(seg, mx, my):
                return

        seg1 = segments.Item(1)
        if try_add(seg1, label_x, label_y):
            return

        # Fallback: place label directly on the segment geometry
        if int(x1) == int(x2) or int(y1) == int(y2):
            if int(x1) == int(x2):
                lx, ly = x1, int((y1 + y2) / 2)
            else:
                lx, ly = int((x1 + x2) / 2), y1
            try_add(seg1, lx, ly)
            return

        # L-shape fallback
        elbow_x, elbow_y = int(x1), int(y2)
        mid1_x, mid1_y = x1, int((y1 + y2) / 2)
        mid2_x, mid2_y = int((x1 + x2) / 2), y2

        if segments.Count >= 2:
            seg2 = segments.Item(2)
            if try_add(seg1, mid1_x, mid1_y):
                return
            if try_add(seg2, mid2_x, mid2_y):
                return
            return

        # If only one segment is reported, try both midpoints on it
        if try_add(seg1, mid1_x, mid1_y):
            return
        try_add(seg1, mid2_x, mid2_y)

        if not get_net_name(net):
            print(f"Net label add failed: {name}")
    except Exception:
        pass


def main():
    app = get_active_app()
    if app is None:
        print("Please open Xpedition Designer and a schematic page first.")
        return

    view = app.ActiveView
    if view is None:
        print("No active schematic view.")
        return

    block = view.Block
    if block is None:
        print("Cannot access current Block.")
        return

    app.SetRedraw(False)
    try:
        # 200 mil = 0.2 inch = 20 units (1 unit = 0.01 inch = 10 mil)
        y_offset = 20
        base_x, base_y = 100, 100 + y_offset

        comp_r1 = block.AddSymbolInstance("Discrete", "RES.1", base_x, base_y)
        if comp_r1 is None:
            print("Cannot add R1. Check symbol Discrete/RES.1.")
            return
        comp_r1.Refdes = "R1"
        normalize_value_attribute(comp_r1, "4.7K")
        set_component_attribute(comp_r1, "DEVICE", "R0603")
        hide_device_attribute(comp_r1)

        # Determine spacing using R1 bounding box
        bbox_ll = comp_r1.GetBboxPoint(VDLOWERLEFT)
        bbox_ur = comp_r1.GetBboxPoint(VDUPPERRIGHT)
        r1_height = bbox_ur.Y - bbox_ll.Y
        if r1_height <= 0:
            r1_height = 100

        # Place R2 below R1 so R1 is the upper resistor (add a small gap)
        gap = 10
        r2_x = base_x
        r2_y = base_y - r1_height - gap

        comp_r2 = block.AddSymbolInstance("Discrete", "RES.1", r2_x, r2_y)
        if comp_r2 is None:
            print("Cannot add R2. Check symbol Discrete/RES.1.")
            return
        comp_r2.Refdes = "R2"
        normalize_value_attribute(comp_r2, "4.7K")
        set_component_attribute(comp_r2, "DEVICE", "R0603")
        hide_device_attribute(comp_r2)

        # Get pins by number (fallback to location order)
        r1_pin1 = find_pin_by_number(comp_r1, "1")
        r1_pin2 = find_pin_by_number(comp_r1, "2")
        if r1_pin1 is None or r1_pin2 is None:
            r1_pin1, r1_pin2 = get_two_pins_by_location(comp_r1)

        r2_pin1 = find_pin_by_number(comp_r2, "1")
        r2_pin2 = find_pin_by_number(comp_r2, "2")
        if r2_pin1 is None or r2_pin2 is None:
            r2_pin1, r2_pin2 = get_two_pins_by_location(comp_r2)

        if None in (r1_pin1, r1_pin2, r2_pin1, r2_pin2):
            print("Cannot resolve resistor pins.")
            return

        # net5v -> R1 pin1
        loc_r1_p1 = get_pin_location(r1_pin1)
        loc_r1_p2 = get_pin_location(r1_pin2)

        # Tail length based on resistor body length (approx)
        body_ratio = 0.4
        pin_dx = abs(loc_r1_p1.X - loc_r1_p2.X)
        pin_dy = abs(loc_r1_p1.Y - loc_r1_p2.Y)
        pin_span = max(pin_dx, pin_dy)
        net_len = max(20, int(pin_span * body_ratio))
        if loc_r1_p1.Y >= loc_r1_p2.Y:
            tail_y = loc_r1_p1.Y + net_len
        else:
            tail_y = loc_r1_p1.Y - net_len
        add_net_with_label(
            block,
            loc_r1_p1.X,
            loc_r1_p1.Y,
            loc_r1_p1.X,
            tail_y,
            r1_pin1,
            None,
            "net5v",
            loc_r1_p1.X + 50,
            tail_y,
        )

        # net2.5v connects R1 pin2 to R2 pin1
        loc_r2_p1 = get_pin_location(r2_pin1)
        add_net_with_label(
            block,
            loc_r1_p2.X,
            loc_r1_p2.Y,
            loc_r2_p1.X,
            loc_r2_p1.Y,
            r1_pin2,
            r2_pin1,
            "net2.5v",
            loc_r1_p2.X + 50,
            int((loc_r1_p2.Y + loc_r2_p1.Y) / 2),
        )

        # netgnd -> R2 pin2
        loc_r2_p2 = get_pin_location(r2_pin2)
        if loc_r2_p2.Y <= loc_r2_p1.Y:
            tail_y = loc_r2_p2.Y - net_len
        else:
            tail_y = loc_r2_p2.Y + net_len
        add_net_with_label(
            block,
            loc_r2_p2.X,
            loc_r2_p2.Y,
            loc_r2_p2.X,
            tail_y,
            r2_pin2,
            None,
            "netgnd",
            loc_r2_p2.X + 50,
            tail_y,
        )

        try:
            view.Refresh()
        except Exception:
            pass

        hidden = hide_device_r0603_in_design(app)
        print(
            "Voltage divider completed: net5v -> R1(1) -> net2.5v -> R2(1) -> netgnd"
        )
        if hidden > 0:
            print(f"DEVICE=R0603 hidden on {hidden} component(s).")
    except Exception as exc:
        print(f"Script error: {exc}")
    finally:
        app.SetRedraw(True)


if __name__ == "__main__":
    main()
