import re, html, os, shutil, zipfile
esc = lambda s: html.escape(s, quote=False)


def col2n(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n


def sheet_path(root, name):
    wb = open(root + "/xl/workbook.xml", encoding="utf-8").read()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wb).group(1)
    rels = open(root + "/xl/_rels/workbook.xml.rels", encoding="utf-8").read()
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return root + "/xl/" + tgt.group(1).lstrip("/").replace("xl/", "")


class Sheet:
    def __init__(self, path):
        self.path = path; self.s = open(path, encoding="utf-8").read()

    def _row_span(self, r):
        m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % r, self.s, flags=re.S)
        return m

    def get(self, ref):
        m = re.search(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>.*?</c>)' % ref, self.s, flags=re.S)
        return m

    def style(self, ref):
        m = self.get(ref)
        if m:
            st = re.search(r'\ss="(\d+)"', m.group(0))
            return st.group(1) if st else None
        return None

    def put(self, ref, cellxml):
        m = self.get(ref)
        if m:
            self.s = self.s[:m.start()] + cellxml + self.s[m.end():]; return
        col = re.match(r"[A-Z]+", ref).group(0); r = int(ref[len(col):])
        rm = self._row_span(r)
        if rm is None:
            # вставить пустую строку в нужное место
            nxt = None
            for m in re.finditer(r'<row r="(\d+)"', self.s):
                if int(m.group(1)) > r:
                    nxt = m.start(); break
            newrow = f'<row r="{r}"></row>'
            if nxt is None:
                k = self.s.find("</sheetData>")
                self.s = self.s[:k] + newrow + self.s[k:]
            else:
                self.s = self.s[:nxt] + newrow + self.s[nxt:]
            rm = self._row_span(r)
        row = rm.group(0)
        if row.endswith("/>"):
            new = row[:-2] + ">" + cellxml + "</row>"
        else:
            cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', row))
            pos = None
            for c in cells:
                if col2n(c.group(1)) > col2n(col):
                    pos = c.start(); break
            if pos is None:
                pos = row.rfind("</row>")
            new = row[:pos] + cellxml + row[pos:]
        self.s = self.s[:rm.start()] + new + self.s[rm.end():]

    def formula(self, ref, f, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}><f>{esc(f)}</f></c>')

    def text(self, ref, t, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa} t="inlineStr"><is><t xml:space="preserve">{esc(t)}</t></is></c>')

    def num(self, ref, v, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}><v>{v}</v></c>')

    def clear(self, ref):
        st = self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}/>')

    def is_empty(self, ref):
        m = self.get(ref)
        return m is None or ("<f" not in m.group(0) and "<v" not in m.group(0) and "<is>" not in m.group(0))

    def save(self):
        open(self.path, "w", encoding="utf-8").write(self.s)


# shared strings (to read values)
def load_sst(WD):
    global SST
    ssx = open(WD + "/xl/sharedStrings.xml", encoding="utf-8").read()
    SST = [re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<si>(.*?)</si>", ssx, flags=re.S)]
    SST = [html.unescape(x) for x in SST]


def value(sh, ref):
    m = sh.get(ref)
    if not m: return None
    c = m.group(0)
    v = re.search(r"<v>(.*?)</v>", c)
    if 't="s"' in c and v: return SST[int(v.group(1))]
    t = re.search(r"<t[^>]*>(.*?)</t>", c, flags=re.S)
    if t: return html.unescape(t.group(1))
    return v.group(1) if v else None


