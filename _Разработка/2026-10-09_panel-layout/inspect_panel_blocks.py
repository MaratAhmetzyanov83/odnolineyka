import pathlib,json,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8');p=pathlib.Path(__file__).parent
a=w.GetActiveObject('AutoCAD.Application');d=next(q for q in a.Documents if q.Name.startswith('Однолинейка '))
names=['S201','S203','DS201','DS203','ABB S201','Дин','3+1','Шина N','VPR-2.5-3L-GY','VPR-6-GY','VPR-6-BU','D-VPR-6','БП DIN','ESB40','Кмка','3 ур. кламмы','F204','R5ISP']
result=[]
for name in names:
    b=d.Blocks.Item(name); rows=[];lo=[];hi=[]
    for e in b:
        r={'type':e.ObjectName}
        if e.ObjectName in ['AcDbAttributeDefinition','AcDbText']:
            r['text']=e.TextString
            if e.ObjectName=='AcDbAttributeDefinition':r['tag']=e.TagString
        try:
            mn,mx=e.GetBoundingBox();lo.append(mn);hi.append(mx)
        except Exception:pass
        rows.append(r)
    item={'name':name,'origin':b.Origin,'objects':rows}
    if lo:item['bbox']=[[min(x[i] for x in lo) for i in range(3)],[max(x[i] for x in hi) for i in range(3)]]
    result.append(item)
    print(json.dumps(item,ensure_ascii=False),flush=True)
(p/'panel_blocks.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print('Source',d.FullName)
