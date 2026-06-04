#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер SVG -> PNG без внешних зависимостей.
Рендерит ограниченное подмножество SVG (используемое в рисунках ВКР)
напрямую через системную библиотеку libcairo (ctypes).

Поддерживается: viewBox, g (transform=translate), rect (rx/ry, скругление),
line, polyline, polygon, circle, ellipse, path (M/L/H/V/C/Q/Z, abs+rel),
text (text-anchor, font-size, font-weight, font-style, fill, fill-opacity),
маркеры-стрелки (marker-start/marker-end -> треугольник), CSS-классы из <style>,
наследование стилей через группы, цвета #rgb/#rrggbb и базовые именованные.
"""
import sys, os, re, math, ctypes
import xml.etree.ElementTree as ET

# ---------- cairo через ctypes ----------
cairo = ctypes.CDLL("libcairo.so.2")
c_double = ctypes.c_double
c_int = ctypes.c_int
c_void_p = ctypes.c_void_p
c_char_p = ctypes.c_char_p


class TextExtents(ctypes.Structure):
    _fields_ = [("x_bearing", c_double), ("y_bearing", c_double),
                ("width", c_double), ("height", c_double),
                ("x_advance", c_double), ("y_advance", c_double)]


def _decl(name, restype, argtypes):
    fn = getattr(cairo, name)
    fn.restype = restype
    fn.argtypes = argtypes
    return fn


cairo_image_surface_create = _decl("cairo_image_surface_create", c_void_p, [c_int, c_int, c_int])
cairo_create = _decl("cairo_create", c_void_p, [c_void_p])
cairo_surface_write_to_png = _decl("cairo_surface_write_to_png", c_int, [c_void_p, c_char_p])
cairo_destroy = _decl("cairo_destroy", None, [c_void_p])
cairo_surface_destroy = _decl("cairo_surface_destroy", None, [c_void_p])
cairo_set_source_rgba = _decl("cairo_set_source_rgba", None, [c_void_p, c_double, c_double, c_double, c_double])
cairo_paint = _decl("cairo_paint", None, [c_void_p])
cairo_save = _decl("cairo_save", None, [c_void_p])
cairo_restore = _decl("cairo_restore", None, [c_void_p])
cairo_translate = _decl("cairo_translate", None, [c_void_p, c_double, c_double])
cairo_scale = _decl("cairo_scale", None, [c_void_p, c_double, c_double])
cairo_set_line_width = _decl("cairo_set_line_width", None, [c_void_p, c_double])
cairo_set_dash = _decl("cairo_set_dash", None, [c_void_p, ctypes.POINTER(c_double), c_int, c_double])
cairo_move_to = _decl("cairo_move_to", None, [c_void_p, c_double, c_double])
cairo_line_to = _decl("cairo_line_to", None, [c_void_p, c_double, c_double])
cairo_curve_to = _decl("cairo_curve_to", None, [c_void_p, c_double, c_double, c_double, c_double, c_double, c_double])
cairo_close_path = _decl("cairo_close_path", None, [c_void_p])
cairo_new_path = _decl("cairo_new_path", None, [c_void_p])
cairo_new_sub_path = _decl("cairo_new_sub_path", None, [c_void_p])
cairo_arc = _decl("cairo_arc", None, [c_void_p, c_double, c_double, c_double, c_double, c_double])
cairo_rectangle = _decl("cairo_rectangle", None, [c_void_p, c_double, c_double, c_double, c_double])
cairo_fill = _decl("cairo_fill", None, [c_void_p])
cairo_fill_preserve = _decl("cairo_fill_preserve", None, [c_void_p])
cairo_stroke = _decl("cairo_stroke", None, [c_void_p])
cairo_set_line_join = _decl("cairo_set_line_join", None, [c_void_p, c_int])
cairo_set_line_cap = _decl("cairo_set_line_cap", None, [c_void_p, c_int])
cairo_select_font_face = _decl("cairo_select_font_face", None, [c_void_p, c_char_p, c_int, c_int])
cairo_set_font_size = _decl("cairo_set_font_size", None, [c_void_p, c_double])
cairo_text_extents = _decl("cairo_text_extents", None, [c_void_p, c_char_p, ctypes.POINTER(TextExtents)])
cairo_show_text = _decl("cairo_show_text", None, [c_void_p, c_char_p])

CAIRO_FORMAT_ARGB32 = 0
SLANT_NORMAL, SLANT_ITALIC = 0, 1
WEIGHT_NORMAL, WEIGHT_BOLD = 0, 1
FONT_FAMILY = b"Noto Sans"

NAMED = {
    "white": (1, 1, 1), "black": (0, 0, 0), "red": (1, 0, 0),
    "green": (0, 0.5, 0), "blue": (0, 0, 1), "none": None,
    "gray": (.5, .5, .5), "grey": (.5, .5, .5),
}


def parse_color(c):
    if c is None:
        return None
    c = c.strip()
    if c == "" or c.lower() == "none":
        return None
    if c.startswith("#"):
        h = c[1:]
        if len(h) == 3:
            h = "".join(ch * 2 for ch in h)
        if len(h) >= 6:
            try:
                return (int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0)
            except ValueError:
                return (0, 0, 0)
    return NAMED.get(c.lower(), (0, 0, 0))


def fnum(v, default=0.0):
    if v is None:
        return default
    m = re.match(r"\s*(-?\d*\.?\d+(?:[eE][-+]?\d+)?)", str(v))
    return float(m.group(1)) if m else default


def collapse(t):
    if t is None:
        return ""
    return re.sub(r"\s+", " ", t).strip()


def local(tag):
    return tag.split("}")[-1]


def parse_style_block(css):
    rules = {}
    for sel, body in re.findall(r"\.([\w\-]+)\s*\{([^}]*)\}", css):
        props = {}
        for decl in body.split(";"):
            if ":" in decl:
                k, v = decl.split(":", 1)
                props[k.strip()] = v.strip()
        rules[sel] = props
    return rules


def set_source(cr, color, alpha=1.0):
    if color is None:
        return False
    cairo_set_source_rgba(cr, color[0], color[1], color[2], alpha)
    return True


def apply_dash(cr, dash):
    if not dash:
        cairo_set_dash(cr, None, 0, 0)
        return
    parts = [float(x) for x in re.split(r"[,\s]+", dash.strip()) if x]
    if not parts:
        cairo_set_dash(cr, None, 0, 0)
        return
    arr = (c_double * len(parts))(*parts)
    cairo_set_dash(cr, arr, len(parts), 0)


PRES_ATTRS = ["fill", "stroke", "stroke-width", "stroke-dasharray", "fill-opacity",
              "opacity", "font-size", "font-weight", "font-style", "font-family",
              "text-anchor"]


def resolve_style(el, parent, class_styles):
    st = dict(parent)
    cls = el.get("class")
    if cls:
        for c in cls.split():
            st.update(class_styles.get(c, {}))
    for a in PRES_ATTRS:
        v = el.get(a)
        if v is not None:
            st[a] = v
    return st


def font_size_of(st):
    return fnum(st.get("font-size", 13), 13)


def stroke_setup(cr, st):
    col = parse_color(st.get("stroke"))
    if col is None:
        return False
    sw = fnum(st.get("stroke-width", 1), 1)
    cairo_set_line_width(cr, sw)
    apply_dash(cr, st.get("stroke-dasharray"))
    set_source(cr, col, fnum(st.get("opacity", 1), 1))
    return True


def fill_setup(cr, st, default_black=True):
    fv = st.get("fill")
    if fv is None:
        if not default_black:
            return False, None
        col = (0, 0, 0)
    else:
        col = parse_color(fv)
    if col is None:
        return False, None
    alpha = fnum(st.get("fill-opacity", 1), 1) * fnum(st.get("opacity", 1), 1)
    return True, (col, alpha)


def rounded_rect(cr, x, y, w, h, r):
    r = min(r, w / 2.0, h / 2.0)
    if r <= 0:
        cairo_rectangle(cr, x, y, w, h)
        return
    cairo_new_sub_path(cr)
    cairo_arc(cr, x + w - r, y + r, r, -math.pi / 2, 0)
    cairo_arc(cr, x + w - r, y + h - r, r, 0, math.pi / 2)
    cairo_arc(cr, x + r, y + h - r, r, math.pi / 2, math.pi)
    cairo_arc(cr, x + r, y + r, r, math.pi, 1.5 * math.pi)
    cairo_close_path(cr)


def draw_arrow(cr, tip, direction, color):
    dx, dy = direction
    L = math.hypot(dx, dy)
    if L == 0:
        return
    ux, uy = dx / L, dy / L
    size, half = 10.0, 4.0
    bx, by = tip[0] - ux * size, tip[1] - uy * size
    px, py = -uy, ux
    cairo_new_path(cr)
    cairo_move_to(cr, tip[0], tip[1])
    cairo_line_to(cr, bx + px * half, by + py * half)
    cairo_line_to(cr, bx - px * half, by - py * half)
    cairo_close_path(cr)
    set_source(cr, color, 1.0)
    cairo_fill(cr)


def marker_color(url, markers):
    if not url:
        return None
    m = re.search(r"url\(#([\w\-]+)\)", url)
    if not m:
        return None
    return markers.get(m.group(1), (0.13, 0.2, 0.27))


TOKEN_RE = re.compile(r"[MmLlHhVvCcSsQqTtAaZz]|-?\d*\.?\d+(?:[eE][-+]?\d+)?")


def draw_path_d(cr, d):
    """Рисует path. Возвращает (first_pt, first_next, last_pt, last_prev) для маркеров."""
    toks = TOKEN_RE.findall(d)
    i, n = 0, len(toks)
    cur = (0.0, 0.0)
    start = (0.0, 0.0)
    cmd = None
    first_pt = None
    first_next = None
    last_pt = None
    last_prev = None

    def num():
        nonlocal i
        v = float(toks[i]); i += 1
        return v

    while i < n:
        t = toks[i]
        if re.match(r"[A-Za-z]", t):
            cmd = t; i += 1
        # implicit repetition keeps cmd; for M/m subsequent coords are L/l
        c = cmd
        if c in ("M", "m"):
            x = num(); y = num()
            if c == "m" and last_pt is not None:
                x += cur[0]; y += cur[1]
            cur = (x, y); start = cur
            cairo_move_to(cr, x, y)
            if first_pt is None:
                first_pt = cur
            last_prev = last_pt
            last_pt = cur
            cmd = "L" if c == "M" else "l"
        elif c in ("L", "l"):
            x = num(); y = num()
            if c == "l":
                x += cur[0]; y += cur[1]
            prev = cur
            cur = (x, y)
            cairo_line_to(cr, x, y)
            if first_next is None:
                first_next = cur
            last_prev = prev
            last_pt = cur
        elif c in ("H", "h"):
            x = num()
            if c == "h":
                x += cur[0]
            prev = cur
            cur = (x, cur[1])
            cairo_line_to(cr, cur[0], cur[1])
            if first_next is None:
                first_next = cur
            last_prev = prev
            last_pt = cur
        elif c in ("V", "v"):
            y = num()
            if c == "v":
                y += cur[1]
            prev = cur
            cur = (cur[0], y)
            cairo_line_to(cr, cur[0], cur[1])
            if first_next is None:
                first_next = cur
            last_prev = prev
            last_pt = cur
        elif c in ("C", "c"):
            x1 = num(); y1 = num(); x2 = num(); y2 = num(); x = num(); y = num()
            if c == "c":
                x1 += cur[0]; y1 += cur[1]; x2 += cur[0]; y2 += cur[1]; x += cur[0]; y += cur[1]
            cairo_curve_to(cr, x1, y1, x2, y2, x, y)
            if first_next is None:
                first_next = (x1, y1)
            last_prev = (x2, y2)
            cur = (x, y); last_pt = cur
        elif c in ("Q", "q"):
            qx = num(); qy = num(); x = num(); y = num()
            if c == "q":
                qx += cur[0]; qy += cur[1]; x += cur[0]; y += cur[1]
            # quadratic -> cubic
            c1 = (cur[0] + 2.0 / 3 * (qx - cur[0]), cur[1] + 2.0 / 3 * (qy - cur[1]))
            c2 = (x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y))
            cairo_curve_to(cr, c1[0], c1[1], c2[0], c2[1], x, y)
            if first_next is None:
                first_next = (qx, qy)
            last_prev = (qx, qy)
            cur = (x, y); last_pt = cur
        elif c in ("Z", "z"):
            cairo_close_path(cr)
            cur = start
        else:
            # неподдерживаемые (A/S/T): пропустить число, чтобы не зациклиться
            try:
                num()
            except Exception:
                i += 1
    return first_pt, first_next, last_pt, last_prev


SKIP = {"defs", "style", "marker", "title", "desc", "metadata", "clipPath", "linearGradient", "radialGradient"}


def render(cr, el, parent_style, class_styles, markers):
    tag = local(el.tag)
    if tag in SKIP:
        return
    st = resolve_style(el, parent_style, class_styles)

    pushed = False
    tr = el.get("transform")
    if tr:
        m = re.search(r"translate\(\s*(-?\d*\.?\d+)[,\s]+(-?\d*\.?\d+)\s*\)", tr)
        ms = re.search(r"scale\(\s*(-?\d*\.?\d+)(?:[,\s]+(-?\d*\.?\d+))?\s*\)", tr)
        cairo_save(cr); pushed = True
        if m:
            cairo_translate(cr, float(m.group(1)), float(m.group(2)))
        if ms:
            sx = float(ms.group(1)); sy = float(ms.group(2)) if ms.group(2) else sx
            cairo_scale(cr, sx, sy)

    if tag in ("svg", "g", "a"):
        for ch in el:
            render(cr, ch, st, class_styles, markers)
        if pushed:
            cairo_restore(cr)
        return

    if tag == "rect":
        x = fnum(el.get("x")); y = fnum(el.get("y"))
        w = fnum(el.get("width")); h = fnum(el.get("height"))
        r = el.get("rx", el.get("ry"))
        rr = fnum(r) if r is not None else 0
        ok, fc = fill_setup(cr, st, default_black=True)
        if ok:
            rounded_rect(cr, x, y, w, h, rr)
            set_source(cr, fc[0], fc[1]); cairo_fill(cr)
        if stroke_setup(cr, st):
            rounded_rect(cr, x, y, w, h, rr); cairo_stroke(cr)

    elif tag == "circle":
        cx = fnum(el.get("cx")); cy = fnum(el.get("cy")); rad = fnum(el.get("r"))
        ok, fc = fill_setup(cr, st, default_black=True)
        if ok:
            cairo_new_path(cr); cairo_arc(cr, cx, cy, rad, 0, 2 * math.pi)
            set_source(cr, fc[0], fc[1]); cairo_fill(cr)
        if stroke_setup(cr, st):
            cairo_new_path(cr); cairo_arc(cr, cx, cy, rad, 0, 2 * math.pi); cairo_stroke(cr)

    elif tag == "ellipse":
        cx = fnum(el.get("cx")); cy = fnum(el.get("cy"))
        rx = fnum(el.get("rx")); ry = fnum(el.get("ry"))
        def ell():
            cairo_save(cr); cairo_translate(cr, cx, cy); cairo_scale(cr, rx, ry)
            cairo_new_path(cr); cairo_arc(cr, 0, 0, 1, 0, 2 * math.pi); cairo_restore(cr)
        ok, fc = fill_setup(cr, st, default_black=True)
        if ok:
            ell(); set_source(cr, fc[0], fc[1]); cairo_fill(cr)
        if stroke_setup(cr, st):
            ell(); cairo_stroke(cr)

    elif tag == "line":
        x1 = fnum(el.get("x1")); y1 = fnum(el.get("y1"))
        x2 = fnum(el.get("x2")); y2 = fnum(el.get("y2"))
        if stroke_setup(cr, st):
            cairo_new_path(cr); cairo_move_to(cr, x1, y1); cairo_line_to(cr, x2, y2); cairo_stroke(cr)
        mc = marker_color(el.get("marker-end"), markers)
        if mc:
            draw_arrow(cr, (x2, y2), (x2 - x1, y2 - y1), mc)
        mc = marker_color(el.get("marker-start"), markers)
        if mc:
            draw_arrow(cr, (x1, y1), (x1 - x2, y1 - y2), mc)

    elif tag in ("polyline", "polygon"):
        nums = [float(v) for v in re.split(r"[,\s]+", el.get("points", "").strip()) if v]
        pts = list(zip(nums[0::2], nums[1::2]))
        if pts:
            def build():
                cairo_new_path(cr); cairo_move_to(cr, pts[0][0], pts[0][1])
                for p in pts[1:]:
                    cairo_line_to(cr, p[0], p[1])
                if tag == "polygon":
                    cairo_close_path(cr)
            if tag == "polygon":
                ok, fc = fill_setup(cr, st, default_black=True)
                if ok:
                    build(); set_source(cr, fc[0], fc[1]); cairo_fill(cr)
            else:
                ok, fc = fill_setup(cr, st, default_black=False)
                if ok:
                    build(); set_source(cr, fc[0], fc[1]); cairo_fill(cr)
            if stroke_setup(cr, st):
                build(); cairo_stroke(cr)

    elif tag == "path":
        d = el.get("d", "")
        info = (None, None, None, None)
        ok, fc = fill_setup(cr, st, default_black=False)
        if ok:
            cairo_new_path(cr); draw_path_d(cr, d)
            set_source(cr, fc[0], fc[1]); cairo_fill(cr)
        if stroke_setup(cr, st):
            cairo_new_path(cr); info = draw_path_d(cr, d); cairo_stroke(cr)
        else:
            cairo_new_path(cr); info = draw_path_d(cr, d); cairo_new_path(cr)
        fp, fn, lp, lpv = info
        mc = marker_color(el.get("marker-end"), markers)
        if mc and lp and lpv:
            draw_arrow(cr, lp, (lp[0] - lpv[0], lp[1] - lpv[1]), mc)
        mc = marker_color(el.get("marker-start"), markers)
        if mc and fp and fn:
            draw_arrow(cr, fp, (fp[0] - fn[0], fp[1] - fn[1]), mc)

    elif tag == "text":
        x = fnum(el.get("x")); y = fnum(el.get("y"))
        txt = collapse(el.text)
        # учесть tspan без собственных координат как продолжение
        for ch in el:
            if local(ch.tag) == "tspan":
                txt = (txt + " " + collapse(ch.text)).strip()
                if ch.tail:
                    txt = (txt + " " + collapse(ch.tail)).strip()
        if txt:
            slant = SLANT_ITALIC if str(st.get("font-style", "")).strip() == "italic" else SLANT_NORMAL
            fw = str(st.get("font-weight", "normal")).strip()
            weight = WEIGHT_BOLD if (fw == "bold" or (fw.isdigit() and int(fw) >= 600)) else WEIGHT_NORMAL
            cairo_select_font_face(cr, FONT_FAMILY, slant, weight)
            cairo_set_font_size(cr, font_size_of(st))
            ext = TextExtents()
            data = txt.encode("utf-8")
            cairo_text_extents(cr, data, ctypes.byref(ext))
            anchor = str(st.get("text-anchor", "start")).strip()
            if anchor == "middle":
                x -= ext.x_advance / 2.0
            elif anchor == "end":
                x -= ext.x_advance
            fv = st.get("fill")
            col = (0, 0, 0) if fv is None else parse_color(fv)
            if col is None:
                col = (0, 0, 0)
            set_source(cr, col, fnum(st.get("fill-opacity", 1), 1))
            cairo_move_to(cr, x, y)
            cairo_show_text(cr, data)

    if pushed:
        cairo_restore(cr)


def collect_markers(root):
    markers = {}
    for el in root.iter():
        if local(el.tag) == "marker":
            mid = el.get("id")
            color = (0.13, 0.2, 0.27)
            for ch in el.iter():
                f = ch.get("fill")
                if f and parse_color(f):
                    color = parse_color(f); break
            if mid:
                markers[mid] = color
    return markers


def collect_css(root):
    css = ""
    for el in root.iter():
        if local(el.tag) == "style":
            css += (el.text or "")
    return parse_style_block(css)


def convert(svg_path, png_path, scale=3.0):
    tree = ET.parse(svg_path)
    root = tree.getroot()
    vb = root.get("viewBox")
    if vb:
        parts = [float(v) for v in re.split(r"[,\s]+", vb.strip())]
        _, _, W, H = parts
    else:
        W = fnum(root.get("width"), 720); H = fnum(root.get("height"), 480)
    pw, ph = int(math.ceil(W * scale)), int(math.ceil(H * scale))
    surface = cairo_image_surface_create(CAIRO_FORMAT_ARGB32, pw, ph)
    cr = cairo_create(surface)
    # белый фон
    cairo_set_source_rgba(cr, 1, 1, 1, 1)
    cairo_paint(cr)
    cairo_scale(cr, scale, scale)
    cairo_set_line_join(cr, 1)  # round
    cairo_set_line_cap(cr, 1)   # round

    class_styles = collect_css(root)
    markers = collect_markers(root)
    base = {"fill": None, "stroke": None, "stroke-width": "1",
            "font-size": root.get("font-size", "13"),
            "font-family": "Noto Sans", "font-weight": "normal",
            "font-style": "normal", "text-anchor": "start"}
    render(cr, root, base, class_styles, markers)

    status = cairo_surface_write_to_png(surface, png_path.encode("utf-8"))
    cairo_destroy(cr)
    cairo_surface_destroy(surface)
    return status, pw, ph


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else "figures_png"
    os.makedirs(out_dir, exist_ok=True)
    targets = []
    if os.path.isdir("figures"):
        for f in sorted(os.listdir("figures")):
            if f.lower().endswith(".svg"):
                targets.append(os.path.join("figures", f))
    for f in sorted(os.listdir(".")):
        if f.lower().endswith(".svg"):
            targets.append(f)
    for svg in targets:
        name = os.path.splitext(os.path.basename(svg))[0] + ".png"
        png = os.path.join(out_dir, name)
        try:
            status, w, h = convert(svg, png)
            ok = "OK" if status == 0 else f"ERR({status})"
            print(f"{ok:8} {svg}  ->  {png}  [{w}x{h}]")
        except Exception as e:
            print(f"FAIL     {svg}: {e}")


if __name__ == "__main__":
    main()
