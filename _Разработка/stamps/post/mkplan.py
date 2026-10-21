import json,os,sys
# mkplan.py <insp.json> <вход-папка> <выход-папка> <list.txt> [режим]
# Режим для всех файлов: 1 — рамка из образца на весь файл, 0 — только поля (и исправление шифра).
# По умолчанию — как в FRAME. Лист 8 лежит в другой папке (ХБК) — его строку пишет mkplan8.
d=json.load(open(sys.argv[1],encoding='utf-8'));src,dst,out=sys.argv[2:5];mode=sys.argv[5] if len(sys.argv)>5 else None
FRAME={"01_Титул.DWG":"1","02_Однолинейка.dwg":"1","02_Однолинейка действующая.dwg":"1","05_КЖ.dwg":"1","06_СО.dwg":"1","Развертки.dwg":"1","03_Чертежи ЭОМ_Пост охраны.dwg":"0"}
def group(s): return " СО" if s['Path'].startswith('/6_') else " ВОР" if s['Path'].startswith('/7_') else ""
def sheet(s): g=group(s); return {"Layout":s['Layout'],"Number":s['Number'],"Offset":"Сквозная нумерация"+g,"Count":"Листов"+g}
plan={}
for s in d['Set']['Sheets']:
  f=os.path.basename(s['File'].replace(chr(92),'/'))
  if f not in FRAME or s['Title']=='Обложка':continue
  plan.setdefault(f,[]).append(sheet(s))
with open(out,'w',encoding='utf-8') as w:
  for f,p in plan.items(): w.write(f"{os.path.join(src,f)}\t{os.path.join(dst,f)}\t{mode or FRAME[f]}\t{json.dumps(p,ensure_ascii=False)}\r\n")
print({k:len(v) for k,v in plan.items()})
