import json,os,sys
# mkplan.py <insp.json> <вход-папка> <выход-папка> <list.txt>
d=json.load(open(sys.argv[1],encoding='utf-8'));src,dst,out=sys.argv[2:5]
frame={"01_Титул.DWG":1,"02_Однолинейка.dwg":1,"02_Однолинейка действующая.dwg":1,"05_КЖ.dwg":1,"06_СО.dwg":1,"Развертки.dwg":1,"03_Чертежи ЭОМ_Пост охраны.dwg":0}
plan={}
for s in d['Set']['Sheets']:
  f=os.path.basename(s['File'].replace(chr(92),'/'))
  if f not in frame or s['Title']=='Обложка':continue
  off="Сквозная нумерация СО" if s['Path'].startswith('/6_') else "Сквозная нумерация ВОР" if s['Path'].startswith('/7_') else "Сквозная нумерация"
  plan.setdefault(f,[]).append({"Layout":s['Layout'],"Number":s['Number'],"Offset":off})
with open(out,'w',encoding='utf-8') as w:
  for f,p in plan.items(): w.write(f"{os.path.join(src,f)}\t{os.path.join(dst,f)}\t{frame[f]}\t{json.dumps(p,ensure_ascii=False)}\r\n")
print({k:len(v) for k,v in plan.items()})
