import subprocess, sys, pathlib
dwg, scr = sys.argv[1], sys.argv[2]
to = int(sys.argv[3]) if len(sys.argv)>3 else 300
cmd=[r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe","/i",dwg,"/s",scr,"/l","en-US"]
try:
    r=subprocess.run(cmd,cwd=str(pathlib.Path(dwg).parent),stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,timeout=to)
    out=r.stdout
except subprocess.TimeoutExpired as e:
    out=e.stdout or b""; print("TIMEOUT")
sys.stdout.reconfigure(encoding="utf-8"); print(out.decode("utf-16le",errors="replace")[-2500:])
