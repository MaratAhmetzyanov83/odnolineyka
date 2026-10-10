import pathlib,sys,json,xml.etree.ElementTree as ET,subprocess,concurrent.futures
sys.path.insert(0,str(pathlib.Path(__file__).parent));import convert as c
P=c.HERE;D=P/'patches';D.mkdir(exist_ok=True)
rows=json.loads((P/'svg_inventory.json').read_text(encoding='utf-8'));fix=[]
for r in rows:
 root=ET.parse(r['file']).getroot()
 if not any(e.tag.endswith('}pattern') for e in root.iter()):continue
 original=r['name'];source=c.ezdxf.new('R2018');source.units=4;source.styles.new('SVG_ARIAL',dxfattribs={'font':'arial.ttf'})
 new=c.convert(pathlib.Path(r['file']),source)
 if not any('Малая растровая' in n for n in new['notes']):continue
 b=source.blocks[new['name']];out=c.ezdxf.new('R2018');out.units=4;out.styles.new('SVG_ARIAL',dxfattribs={'font':'arial.ttf'})
 for e in b:out.modelspace().add_foreign_entity(e,copy=True)
 dest=D/f'p{r["id"]:04}.dxf';out.saveas(dest)
 script=dest.with_suffix('.scr');dwg=dest.with_suffix('.dwg')
 script.write_text('FILEDIA\n0\n_SAVEAS\n2018\n'+dwg.as_posix()+'\n_QUIT\n_Y\n',encoding='utf-8')
 fix.append((new,dest,script,dwg));print('TEXTURE',r['id'],len(b),flush=True)
def run(item):
 row,dxf,scr,dwg=item
 result=subprocess.run([r'C:\Program Files\Autodesk\AutoCAD 2027\accoreconsole.exe','/i',str(dxf),'/s',str(scr)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
 assert result.returncode==0 and dwg.exists(),(row['id'],result.returncode)
 return row
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
 for r in pool.map(run,fix):print('NATIVE PATCH',r['id'],flush=True)
mapping={r['name']:r for r,_,_,_ in fix}
for fname in ['svg_inventory.json','catalog_all.json']:
 p=P/fname;data=json.loads(p.read_text(encoding='utf-8'))
 for old in data:
  if old['name'] in mapping:old.update({k:mapping[old['name']][k] for k in ['notes','entities']})
 p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
def quote(s):return '"'+s.replace('\\','/').replace('"','\\"')+'"'
lines=['(setvar "CMDECHO" 0)','(setvar "FILEDIA" 0)','(setvar "TILEMODE" 1)','(setvar "ATTREQ" 0)','(setvar "OSMODE" 0)']
for row,_,_,dwg in fix:
 spec=row['name']+'='+dwg.as_posix()
 lines+=['(command "_.-INSERT" '+quote(spec)+')',
 '(setq patprompt (getvar "LASTPROMPT"))',
 '(if (or (wcmatch (strcase patprompt) "*REDEFINE*") (wcmatch (strcase patprompt) "*ПЕРЕОПРЕДЕЛ*")) (command "_Y"))',
 '(command "_non" (list 0.0 0.0 0.0) 1.0 1.0 0.0)',
 '(if (= (cdr (assoc 0 (entget (entlast)))) "INSERT") (entdel (entlast)))']
lines+=['(command "_.AUDIT" "_Y")','(command "_.QSAVE")','(princ "PATTERNS_COMPLETE") (princ)']
(P/'patch_patterns.lsp').write_text('\n'.join(lines),encoding='utf-8')
(P/'patch_patterns.scr').write_text('(load "'+(P/'patch_patterns.lsp').as_posix()+'")\n_QUIT\n_Y\n',encoding='utf-8')
print('PATCHES',len(fix),flush=True)
