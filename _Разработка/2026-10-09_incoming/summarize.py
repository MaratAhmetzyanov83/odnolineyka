import pathlib,re,json,sys
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent
s=(p/'acad_dump.txt').read_text(encoding='cp1251',errors='replace')
es=[]
for raw in re.split(r'(?=\(\(-1 \. <)',s):
    if not raw.startswith('((-1'): continue
    def st(k):
        m=re.search(r'\('+str(k)+r' \. "((?:\\.|[^"\\])*)"\)',raw)
        return m.group(1) if m else None
    def pt(k):
        m=re.search(r'\('+str(k)+r' ([-\d.eE]+) ([-\d.eE]+)(?: ([-\d.eE]+))?\)',raw)
        return [float(x) for x in m.groups() if x] if m else None
    r={'h':st(5),'t':st(0),'l':st(8),'p':pt(10),'end':pt(11),'name':st(2),'txt':st(1)}
    if r['h'] in sys.argv[1:]: print(raw)
    if r['t']=='INSERT': r['details']=raw
    es.append(r)
(p/'parsed_entities.json').write_text(json.dumps(es,ensure_ascii=False),encoding='utf-8')
for r in es:
    if len(sys.argv)==1 and r['p'] and 55<=r['p'][0]<=180 and 1405<=r['p'][1]<=1577:
        print(json.dumps(r,ensure_ascii=False))
