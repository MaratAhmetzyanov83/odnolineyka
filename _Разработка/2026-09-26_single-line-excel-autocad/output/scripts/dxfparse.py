"""Minimal DXF reader (no deps). Produces sections -> list of entities (list of (code,value))."""
import pickle, sys, collections

PATH = "/mnt/user-data/uploads/Desktop/02_Однолинейка.dxf"

def pairs(path):
    with open(path, "rb") as f:
        data = f.read().decode("utf-8", errors="replace").splitlines()
    it = iter(data)
    for code in it:
        try:
            val = next(it)
        except StopIteration:
            return
        yield int(code.strip()), val.rstrip("\r")

def parse(path):
    sections = {}
    cur_sec = None
    ents = None
    ent = None
    expect_name = False
    for code, val in pairs(path):
        if code == 0 and val == "SECTION":
            expect_name = True
            continue
        if expect_name and code == 2:
            cur_sec = val; ents = []; sections[cur_sec] = ents; expect_name = False; ent = None
            continue
        if code == 0 and val == "ENDSEC":
            cur_sec = None; ent = None
            continue
        if cur_sec is None:
            continue
        if code == 0:
            ent = [(0, val)]
            ents.append(ent)
        elif ent is not None:
            ent.append((code, val))
        else:
            ents.append([(code, val)])
    return sections

if __name__ == "__main__":
    s = parse(PATH)
    for k, v in s.items():
        print(k, len(v))
    pickle.dump(s, open("/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/dxf.pkl", "wb"))
