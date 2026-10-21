import json,os,sys
for r in json.load(open(sys.argv[1],encoding='utf-8')):
  n=os.path.basename(r['File'].replace(chr(92),'/'))
  if 'Error' in r: print('ERR',n,r['Error'][:800]);continue
  R=r['Result'] or {'Nested':'-','Frames':[]};print('##',n,r['Version'],'nested',R['Nested'],'frames',len(R['Frames']))
  for f in R['Frames']:
    if f['PropDiff'] or f['AttrDiff']: print('   DIFF',f['Layout'],f['PropDiff'],f['AttrDiff'])
