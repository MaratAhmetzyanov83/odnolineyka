import re,collections,sys
s=open(sys.argv[1],encoding='utf-8').read()
depth=0;i=0;n=len(s);bad=[];line=1;tops=[]
while i<n:
    c=s[i]
    if c=='\n': line+=1
    if c==';':
        while i<n and s[i]!='\n': i+=1
        continue
    if c=='"':
        j=i+1
        while s[j]!='"':
            if s[j]=='\\': j+=1
            if s[j]=='\n': line+=1
            j+=1
        if any(ord(ch)>127 for ch in s[i:j]): bad.append((line,s[i:j][:40]))
        i=j+1;continue
    if c=='(':
        if depth==0: tops.append(line)
        depth+=1
    elif c==')':
        depth-=1
        if depth<0: print('NEG',line); depth=0
    i+=1
s2=re.sub(r';[^\n]*','',s); s2=re.sub(r'"(\\.|[^"\\])*"','""',s2)
d=collections.defaultdict(set)
for x in set(re.findall(r'[A-Za-z*][A-Za-z0-9:*_\-]*',s2)): d[x.lower()].add(x)
print('depth',depth,'nonascii',bad,'collide',[v for v in d.values() if len(v)>1 and not ({x.lower() for x in v} & {'t','open','close','name','h'})])
# defun top-level check: each top form should start with (defun/(setq/(if/(princ/(vl-load
import itertools
for ln in tops:
    pass
# undefined sx: functions
defs=set(m.lower() for m in re.findall(r'\(defun\s+([^\s(]+)',s2))
calls=set(m.lower() for m in re.findall(r"\((sx:[A-Za-z0-9:\-_]+)",s2))
print('undefined calls',sorted(calls-defs))
