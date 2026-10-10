import pathlib,json,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent
a=w.GetActiveObject('AutoCAD.Application')
out=[]
for d in a.Documents:
    names=[]
    for b in d.Blocks:
        name=b.Name
        if not name.startswith('*'):
            names.append({'name':name,'count':b.Count,'xref':b.IsXRef})
    out.append({'name':d.Name,'path':d.FullName,'blocks':names})
    print(d.Name,len(names),flush=True)
(p/'block_inventory.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
for d in out:
    print('\nDRAWING',d['name'])
    print('\n'.join(b['name'] for b in d['blocks'] if not b['xref']))
