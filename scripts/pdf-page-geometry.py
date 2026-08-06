"""Page geometry from a Word PDF's content stream, in top-down CSS px at 96 DPI.

Images come from the CTM that scales the unit image square (`cm` then `Do`), so
the reported box is the placed box, not the ink. Text comes from Tm/Td, so the
reported y is the BASELINE, not a line-box top -- the two differ by the ascent
and must never be compared directly against a DOM line top.
"""
import sys
import pikepdf
from pikepdf import Name, Operator

PT_TO_PX = 96.0 / 72.0


def mul(m, n):
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + e * 0 + E, e * B + f * D + F)


def run(path, pageno):
    pdf = pikepdf.open(path)
    page = pdf.pages[pageno - 1]
    box = page.mediabox
    ph = float(box[3]) - float(box[1])
    ctm = (1, 0, 0, 1, 0, 0)
    stack = []
    tm = None
    out = []
    for operands, op in pikepdf.parse_content_stream(page):
        o = str(op)
        if o == "q":
            stack.append(ctm)
        elif o == "Q":
            ctm = stack.pop() if stack else (1, 0, 0, 1, 0, 0)
        elif o == "cm":
            ctm = mul(tuple(float(x) for x in operands), ctm)
        elif o == "Do":
            a, b, c, d, e, f = ctm
            w, h = abs(a), abs(d)
            top = ph - (f + h)
            out.append(("IMG", str(operands[0]), top * PT_TO_PX, h * PT_TO_PX, e * PT_TO_PX, w * PT_TO_PX))
        elif o == "BT":
            tm = (1, 0, 0, 1, 0, 0)
        elif o in ("Tm",):
            tm = tuple(float(x) for x in operands)
        elif o in ("Td", "TD") and tm:
            tx, ty = float(operands[0]), float(operands[1])
            tm = mul((1, 0, 0, 1, tx, ty), tm)
        elif o in ("Tj", "TJ") and tm:
            txt = ""
            if o == "Tj":
                txt = bytes(operands[0]).decode("latin-1", "replace")
            else:
                for it in operands[0]:
                    if isinstance(it, pikepdf.String):
                        txt += bytes(it).decode("latin-1", "replace")
            if txt.strip():
                out.append(("TXT", txt.strip()[:46], (ph - tm[5]) * PT_TO_PX, 0.0, tm[4] * PT_TO_PX, 0.0))
    out.sort(key=lambda r: r[2])
    print(f"# {path} page {pageno}  mediabox {float(box[2]):.1f}x{ph:.1f}pt = {float(box[2])*PT_TO_PX:.1f}x{ph*PT_TO_PX:.1f}px")
    for kind, label, top, h, x, w in out:
        if kind == "IMG":
            print(f"  IMG  top={top:8.2f} h={h:7.2f} bottom={top+h:8.2f} x={x:7.2f} w={w:7.2f}  {label}")
        else:
            print(f"  TXT  baseline={top:8.2f}                    x={x:7.2f}          {label!r}")


run(sys.argv[1], int(sys.argv[2]))
