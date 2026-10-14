$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$acadPath = 'C:\Program Files\Autodesk\AutoCAD 2025'
$assembly = [Reflection.Assembly]::LoadFrom("$acadPath\AcSmComponents.Interop.dll")
Add-Type -ReferencedAssemblies "$acadPath\AcSmComponents.Interop.dll" -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
public static class SsmTask {
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode)] static extern bool SetDllDirectory(string path);
 static void Walk(IAcSmSubset subset,string path) {
  var en=subset.GetSheetEnumerator(); IAcSmComponent item;
  while((item=en.Next())!=null) {
   if(item.GetTypeName()=="AcSmSheet") {
    var sheet=(IAcSmSheet)item; var layout=sheet.GetLayout();
    Console.WriteLine("SHEET\t"+path+"\t"+sheet.GetNumber()+"\t"+sheet.GetTitle()+"\t"+layout.GetFileName()+"\t"+layout.GetName());
   } else if(item.GetTypeName()=="AcSmSubset") {
    Console.WriteLine("SUBSET\t"+path+"/"+item.GetName()); Walk((IAcSmSubset)item,path+"/"+item.GetName());
   } else Console.WriteLine("OTHER\t"+item.GetTypeName());
  }
 }
 public static void Inspect(string filename) {
  SetDllDirectory(@"C:\Program Files\Autodesk\AutoCAD 2025");
  Console.WriteLine("STAGE create");
  IAcSmSheetSetMgr mgr=new AcSmSheetSetMgrClass();
  Console.WriteLine("STAGE enumerate");
  var en=mgr.GetDatabaseEnumerator(); AcSmDatabase db;
  while((db=en.Next())!=null) Console.WriteLine("OPEN\t"+db.GetFileName());
  Console.WriteLine("STAGE find "+filename);
  try { db=mgr.FindOpenDatabase(filename); } catch(COMException) { db=null; }
  Console.WriteLine("STAGE open");
  if(db==null) db=mgr.OpenDatabase(filename,false);
  Console.WriteLine("DATABASE\t"+db.GetFileName()+"\t"+db.GetLockStatus());
  Console.WriteLine("TITLE\t"+db.GetSheetSet().GetName());
  Walk((IAcSmSubset)db.GetSheetSet(),"");
 }
}
'@
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('AutoCAD.Application')
[SsmTask]::Inspect($app.ActiveDocument.GetVariable('SSFOUND'))
