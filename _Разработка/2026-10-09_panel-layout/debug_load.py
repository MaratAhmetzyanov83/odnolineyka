import pathlib,time,win32com.client as w
p=pathlib.Path(__file__).parent;root=p.parents[1]
s=(root/'2_AutoCAD/Компоновка/SX_Panel.lsp').read_text(encoding='utf-8')
forms=[];depth=0;string=False;comment=False;i=0;start=0;line=1;startline=0
while i<len(s):
    c=s[i]
    if c=='\n':line+=1
    if comment:
        if c=='\n':comment=False
    elif string:
        if c=='\\':i+=1
        elif c=='"':string=False
    elif c==';':comment=True
    elif c=='"':string=True
    elif c=='(':
        if depth==0:start=i;startline=line
        depth+=1
    elif c==')':
        depth-=1
        if depth==0:forms.append((startline,s[start:i+1]))
    i+=1
def ls(x):return '"'+x.replace('\\','\\\\').replace('"','\\"').replace('\n','\\n')+'"'
code='(setq zf (open '+ls(str(p/'load_debug.txt').replace('\\','/'))+' "w"))\n'
for ln,f in forms:
    code+='(setq zr (vl-catch-all-apply \'(lambda () (eval (read '+ls(f)+'))) nil))\n'
    code+='(prin1 (list '+str(ln)+' (if (vl-catch-all-error-p zr) (vl-catch-all-error-message zr) "OK")) zf) (write-line "" zf)\n'
code+='(close zf) (princ)\n'
(p/'load_debug.lsp').write_text(code,encoding='utf-8')
a=w.GetActiveObject('AutoCAD.Application');d=a.ActiveDocument;assert d.Name=='Компоновка_проверка_из_базы.dwg'
d.SendCommand('(load "'+str(p/'load_debug.lsp').replace('\\','/')+'")\n')
print('Diagnostic requested',len(forms))
