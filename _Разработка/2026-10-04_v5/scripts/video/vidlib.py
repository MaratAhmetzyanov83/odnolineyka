# -*- coding: utf-8 -*-
"""vidlib.py — движок видеогайда: снимки областей листов Excel (через PDF), кадры 1920×1080, озвучка SAPI, сборка ffmpeg.

Снимок: область листа печатается в PDF с известным масштабом (поля 0,2", A3 альбомная), PDF растрируется (PyMuPDF),
координаты ячеек пересчитываются в пиксели — рамки подсветки попадают точно в ячейки. Масштаб по каждой оси уточняется
по линиям сетки калибровочной печати (Excel печатает столбцы шире, чем по их ширине на экране).
"""
import os
import subprocess
import wave

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFont

W, H = 1920, 1080
TOP, BOT, PAD = 104, 132, 34
DPI = 220
MARGIN_IN = 0.2
C_TITLE = (31, 58, 95)
C_ACCENT = (255, 122, 0)
C_BG = (243, 245, 248)
FONT = r"C:\Windows\Fonts\segoeui.ttf"
FONT_B = r"C:\Windows\Fonts\segoeuib.ttf"


def font(size, bold=False):
    return ImageFont.truetype(FONT_B if bold else FONT, size)


class Shot:
    """снимок области: изображение и пересчёт адресов ячеек в прямоугольники на нём (kx, ky — пикселей на пункт ширины / высоты)."""

    def __init__(self, img, ws, area, kx, ky):
        self.img, self.ws, self.area, self.kx, self.ky = img, ws, area, kx, ky

    def rect(self, addr):
        """положение по сумме ширин и высот (Left / Top после смены ширины столбца Excel отдаёт устаревшими)."""
        ws, a, kx, ky = self.ws, self.area, self.kx, self.ky
        r = ws.Range(addr)
        x0 = cols_w(ws, a.Column, r.Column) * kx
        y0 = rows_h(ws, a.Row, r.Row) * ky
        return (x0, y0, x0 + cols_w(ws, r.Column, r.Column + r.Columns.Count) * kx, y0 + rows_h(ws, r.Row, r.Row + r.Rows.Count) * ky)


def cols_w(ws, c0, c1):
    return sum(ws.Columns(c).Width for c in range(c0, c1))


def rows_h(ws, r0, r1):
    return sum(ws.Rows(r).Height for r in range(r0, r1))


def _lines(g, axis, lo, hi):
    """центры тонких тёмных линий (сетка, границы ячеек): axis=1 — вертикальные (x), 0 — горизонтальные (y); lo:hi — полоса поперёк линий."""
    D = 4
    s = (g[lo:hi, :] if axis == 1 else g[:, lo:hi].T).astype(np.int16)
    t = (s[:, D:-D] < s[:, :-2 * D] - 20) & (s[:, D:-D] < s[:, 2 * D:] - 20)
    out, grp = [], []
    for i in np.where(t.mean(axis=0) > 0.25)[0]:
        if grp and i - grp[-1] > 1:
            out.append(sum(grp) / len(grp) + D)
            grp = []
        grp.append(i)
    if grp:
        out.append(sum(grp) / len(grp) + D)
    return np.array(out)


def _fit(b, L, o):
    """масштаб s, при котором линии сетки L ложатся на расчётные границы o + b·s.

    Excel печатает столбцы шире, чем по Columns().Width (ширина — в символах шрифта, а метрики шрифта у принтера свои; у нас ~+4,5 %),
    высоту строк — почти точно. Один масштаб на ось: остаток — около пикселя на всю ширину A3.
    """
    if len(L) < 2 or b[-1] <= 0:
        return 1.0
    L = np.sort(L)
    best, bs = -1.0, 1.0
    for s in np.arange(0.95, 1.12, 0.0002):
        p = o + 1 + b * s
        i = np.clip(np.searchsorted(L, p), 1, len(L) - 1)
        d = np.minimum(np.abs(L[i] - p), np.abs(L[i - 1] - p))
        sc = np.clip(2.5 - d, 0, None).sum()
        if sc > best:
            best, bs = sc, s
    return bs


def _export(ws, pdf):
    if os.path.exists(pdf):
        os.remove(pdf)
    ws.ExportAsFixedFormat(0, pdf, 0, False, False)
    return fitz.open(pdf)


def _raster(doc, png):
    doc[0].get_pixmap(dpi=DPI).save(png)
    doc.close()
    return Image.open(png).convert("RGB")


def _page(xl, ws, area_addr, gridlines):
    """печать области: A3 альбомная, поля 0,2", без колонтитулов и заголовков; вернуть (PageSetup, поле в пт)."""
    ps = ws.PageSetup
    ps.PrintArea = area_addr
    try:
        ps.PaperSize = 8                                    # A3
    except Exception:
        pass
    ps.Orientation = 2
    m = xl.InchesToPoints(MARGIN_IN)
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = m
    ps.HeaderMargin = ps.FooterMargin = 0
    for h in ("LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter"):
        setattr(ps, h, "")
    ps.CenterHorizontally = False
    ps.CenterVertically = False
    ps.PrintGridlines = gridlines
    ps.PrintHeadings = False
    return ps, m


_CAL = {}
CAL_W = [2, 7, 13, 4, 21, 9, 30, 5, 11, 17, 3, 8]          # служебная сетка: ширины столбцов (символов) и высоты строк (пт)
CAL_H = [15, 9.6, 27, 12.6, 18, 36]


def _calib(xl, wb, zoom, tmp):
    """масштаб печати по осям при данном Zoom: (sx, sy) — во сколько раз печать больше расчёта по Width / Height.

    Сетка на листах с заливкой (таблицы) не печатается, поэтому калибровка — по служебной книге с тем же шрифтом «Обычный»
    (от него зависит ширина столбца) и чистой сеткой; книга закрывается без сохранения.
    """
    f = wb.Styles("Normal").Font
    key = (f.Name, f.Size, zoom)
    if key in _CAL:
        return _CAL[key]
    act = xl.ActiveWorkbook
    nb = xl.Workbooks.Add()
    try:
        nf = nb.Styles("Normal").Font
        nf.Name, nf.Size = key[0], key[1]
        ws = nb.Worksheets(1)
        lim_w = (1190.55 - 2 * 14.4) / (zoom / 100.0) / 1.1     # сколько влезет на лист A3 при этом масштабе (с запасом)
        lim_h = (841.89 - 2 * 14.4) / (zoom / 100.0) / 1.03
        cw, rh = [], []
        while len(cw) < 60:
            c = len(cw) + 1
            ws.Columns(c).ColumnWidth = CAL_W[(c - 1) % len(CAL_W)]
            if sum(cw) + ws.Columns(c).Width > lim_w and len(cw) >= 3:
                break
            cw.append(ws.Columns(c).Width)
        while len(rh) < 60:
            r = len(rh) + 1
            ws.Rows(r).RowHeight = CAL_H[(r - 1) % len(CAL_H)]
            if sum(rh) + ws.Rows(r).Height > lim_h and len(rh) >= 3:
                break
            rh.append(ws.Rows(r).Height)
        ws.Cells(1, 1).Value = "."
        ws.Cells(len(rh), len(cw)).Value = "."
        ps, _ = _page(xl, ws, ws.Range(ws.Cells(1, 1), ws.Cells(len(rh), len(cw))).Address, True)
        ps.Zoom = zoom
        g = np.asarray(_raster(_export(ws, os.path.join(tmp, "_cal.pdf")), os.path.join(tmp, "_cal.png")).convert("L"))
        k = zoom / 100.0 * DPI / 72.0
        o = MARGIN_IN * DPI
        bx = np.concatenate([[0.0], np.cumsum(cw)]) * k
        by = np.concatenate([[0.0], np.cumsum(rh)]) * k
        sx = _fit(bx, _lines(g, 1, int(o), min(g.shape[0], int(o + by[-1] * 1.05) + 6)), o)
        sy = _fit(by, _lines(g, 0, int(o), min(g.shape[1], int(o + bx[-1] * sx) + 6)), o)
    finally:
        nb.Close(False)
        try:
            act.Activate()
        except Exception:
            pass
    _CAL[key] = (sx, sy)
    return sx, sy


def shot(xl, ws, area_addr, tmp, name, gridlines=True):
    """напечатать область в PDF и растрировать; масштаб по осям — по служебной сетке при том же Zoom; вернуть Shot."""
    ps, m = _page(xl, ws, area_addr, gridlines)
    a = ws.Range(area_addr)
    cw = [ws.Columns(c).Width for c in range(a.Column, a.Column + a.Columns.Count)]
    rh = [ws.Rows(r).Height for r in range(a.Row, a.Row + a.Rows.Count)]
    aw, ah = sum(cw) * 1.06, sum(rh) * 1.01                # печать шире расчёта — с запасом, чтобы влезло на лист
    pw, ph = 1190.55, 841.89
    zoom = int(min(100, (pw - 2 * m) / aw * 100, (ph - 2 * m) / ah * 100))
    pdf = os.path.join(tmp, name + ".pdf")
    for _ in range(12):                                    # если не влезло на страницу (принтер не дал A3) — меньше масштаб
        ps.Zoom = max(10, zoom)
        doc = _export(ws, pdf)
        if doc.page_count == 1 or zoom <= 10:
            break
        pw_, ph_ = doc[0].rect.width, doc[0].rect.height
        doc.close()
        zoom = int(min(zoom * 0.92, (pw_ - 2 * m) / aw * 100, (ph_ - 2 * m) / ah * 100))
    img = _raster(doc, os.path.join(tmp, name + ".png"))
    z = ps.Zoom
    sx, sy = _calib(xl, ws.Parent, z, tmp)
    k = z / 100.0 * DPI / 72.0
    o = MARGIN_IN * DPI
    img = img.crop((int(o), int(o), int(o + sum(cw) * k * sx) + 2, int(o + sum(rh) * k * sy) + 2))
    return Shot(img, ws, a, k * sx, k * sy)


def wrap(draw, text, fnt, width):
    words, lines, cur = text.split(), [], ""
    for w_ in words:
        t = (cur + " " + w_).strip()
        if draw.textlength(t, font=fnt) <= width:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    return lines


def frame(out, title, caption, sh=None, marks=(), note=None, crop=None):
    """кадр: заголовок главы, снимок (вписан), рамки подсветки, подпись внизу."""
    im = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, W, TOP), fill=C_TITLE)
    d.text((PAD + 6, TOP // 2), title, font=font(46, True), fill=(255, 255, 255), anchor="lm")
    d.text((W - PAD - 6, TOP // 2), "Однолинейка v5", font=font(28), fill=(170, 190, 215), anchor="rm")
    if sh is not None:
        src = sh.img
        rects = [a if isinstance(a, tuple) else sh.rect(a) for a in marks]
        if crop:                                           # кадрирование по прямоугольнику ячеек (адрес)
            cx0, cy0, cx1, cy1 = crop if isinstance(crop, tuple) else sh.rect(crop)
            cx0, cy0 = max(0, cx0 - 12), max(0, cy0 - 12)
            cx1, cy1 = min(src.width, cx1 + 12), min(src.height, cy1 + 12)
            src = src.crop((int(cx0), int(cy0), int(cx1), int(cy1)))
            rects = [(x0 - cx0, y0 - cy0, x1 - cx0, y1 - cy0) for x0, y0, x1, y1 in rects]
        bw, bh = W - 2 * PAD, H - TOP - BOT - 2 * PAD
        s = min(bw / src.width, bh / src.height, 1.6)
        nw, nh = int(src.width * s), int(src.height * s)
        x, y = (W - nw) // 2, TOP + PAD + (bh - nh) // 2
        d.rectangle((x - 3, y - 3, x + nw + 3, y + nh + 3), fill=(205, 210, 218))
        im.paste(src.resize((nw, nh), Image.LANCZOS), (x, y))
        if rects:
            ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            od = ImageDraw.Draw(ov)
            for x0, y0, x1, y1 in rects:
                bx = (x + x0 * s - 6, y + y0 * s - 6, x + x1 * s + 6, y + y1 * s + 6)
                od.rounded_rectangle(bx, radius=10, fill=(255, 160, 0, 46), outline=C_ACCENT + (255,), width=6)
            im = Image.alpha_composite(im.convert("RGBA"), ov).convert("RGB")
            d = ImageDraw.Draw(im)
    if note:                                               # крупная надпись на пустом кадре (титул, итог)
        f1 = font(64, True)
        lines = wrap(d, note, f1, W - 300)
        y = (H - len(lines) * 84) // 2 - 20
        for ln in lines:
            d.text((W // 2, y), ln, font=f1, fill=C_TITLE, anchor="ma")
            y += 84
    d.rectangle((0, H - BOT, W, H), fill=(255, 255, 255))
    d.line((0, H - BOT, W, H - BOT), fill=(214, 219, 226), width=2)
    f2 = font(34)
    lines = wrap(d, caption, f2, W - 2 * PAD - 20)[:2]
    y = H - BOT + (BOT - len(lines) * 46) // 2
    for ln in lines:
        d.text((PAD + 10, y), ln, font=f2, fill=(30, 36, 44))
        y += 46
    im.save(out)


SAY = [("УЗО.1", "УЗО точка один"), ("DLN", "Дэ Эл Эн"), ("AutoCAD", "Автокад"), ("Data Link", "Дата Линк"), ("SXDRAW", "Эс Икс Дро"),
       ("кВт", "киловатт"), (" мА", " миллиампер"), ("мм²", "квадратных миллиметров"), (" мм", " миллиметров"), ("DIN", "дин"),
       ("DALI", "Дали"), ("Relay", "реле"), ("LED", "светодиодные"), ("ВРУ", "вэ эр у"), ("ПТО", "пэ тэ о"), ("БП", "бэ пэ"),
       ("МК-5-1", "эм ка пять один"), ("A4", "А четыре"), ("C3", "цэ три"), ("24V 480W", "двадцать четыре вольта, четыреста восемьдесят ватт"),
       ("1,3", "один и три десятых"), ("«", ""), ("»", ""), ("—", ","), ("ΔU", "потери напряжения"), ("Iкз", "ток короткого замыкания")]


def say_text(t):
    for a, b in SAY:
        t = t.replace(a, b)
    return t


def tts(text, wav, rate=1):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    v = win32com.client.Dispatch("SAPI.SpVoice")
    for t in v.GetVoices():
        if "Irina" in t.GetDescription():
            v.Voice = t
    v.Rate = rate
    st = win32com.client.Dispatch("SAPI.SpFileStream")
    fmt = win32com.client.Dispatch("SAPI.SpAudioFormat")
    fmt.Type = 22                                          # 22 кГц, 16 бит, моно
    st.Format = fmt
    st.Open(wav, 3, False)
    v.AudioOutputStream = st
    v.Speak(text, 0)
    st.Close()
    with wave.open(wav) as w_:
        return w_.getnframes() / float(w_.getframerate())


def ffmpeg(args):
    r = subprocess.run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error"] + args, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(r.stderr[-800:])


def clip(png, wav, dur, out, tail=0.7):
    ffmpeg(["-loop", "1", "-framerate", "25", "-i", png, "-i", wav, "-filter_complex", "[1:a]apad=pad_dur=%.2f[a]" % tail,
            "-map", "0:v", "-map", "[a]", "-c:v", "libx264", "-preset", "medium", "-tune", "stillimage", "-crf", "20", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-t", "%.2f" % (dur + tail), out])
    return dur + tail


def srt_time(t):
    h, r = divmod(t, 3600)
    m_, s = divmod(r, 60)
    return "%02d:%02d:%02d,%03d" % (h, m_, int(s), int((s - int(s)) * 1000))
