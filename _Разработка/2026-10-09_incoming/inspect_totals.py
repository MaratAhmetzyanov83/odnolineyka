import pathlib,json,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent
u=json.loads((p/'used_sheet.json').read_text(encoding='utf-8'))
for i,r in enumerate(u['values']):
    for j,v in enumerate(r):
        if isinstance(v,str) and any(t in v.lower() for t in ['руст','l1р','l1p','cos(','кc=','кс=']): print('SAVED',i+1,j+1,v)
x=w.GetActiveObject('Excel.Application')
for b in x.Workbooks:
    print('BOOK',b.FullName,'ACTIVE',b.ActiveSheet.Name)
    for s in b.Worksheets:
        if s.Name in ['Однолинейка','В Акад','Нагрузка_щитов','Расчёт_нагрузок']:
            vals=s.UsedRange.Value; found=[]
            for i,r in enumerate(vals):
                for j,v in enumerate(r):
                    if isinstance(v,str) and any(t in v.lower() for t in ['руст','l1p','l1р','cos(ф','кс=','кc=']): found.append((i+s.UsedRange.Row,j+s.UsedRange.Column,v))
            print('SHEET',s.Name,s.UsedRange.Address,'HITS',found[:30])
            if found:
                out={'path':b.FullName,'sheet':s.Name,'row':s.UsedRange.Row,'col':s.UsedRange.Column,'values':vals,'formulas':s.UsedRange.Formula}
                (p/('totals_'+str(len(found))+'_'+s.Name.replace(' ','_')+'.json')).write_text(json.dumps(out,ensure_ascii=False,default=str),encoding='utf-8')
es=json.loads((p/'parsed_entities.json').read_text(encoding='utf-8'))
for e in es:
    if e['p'] and 55<=e['p'][0]<=180 and 1285<=e['p'][1]<1345.99: print('CAD',json.dumps(e,ensure_ascii=False))
