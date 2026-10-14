using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class SsmRenumberCommands {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\ssm_task";
 public record Row(string Section,string Number,string Title,string File,string Layout,string NewNumber);
 public record Plan(string File,Row[] Sheets);
 static void Walk(IAcSmSubset subset,string section,List<Row> rows,List<IAcSmSheet> sheets) {
  var en=subset.GetSheetEnumerator(); IAcSmComponent item;
  while((item=en.Next())!=null) {
   if(item.GetTypeName()=="AcSmSheet") {
    var s=(IAcSmSheet)item; var l=s.GetLayout(); sheets.Add(s);
    rows.Add(new Row(section,s.GetNumber(),s.GetTitle(),l.GetFileName(),l.GetName(),""));
   } else if(item.GetTypeName()=="AcSmSubset") Walk((IAcSmSubset)item,section+"/"+item.GetName(),rows,sheets);
  }
 }
 static bool Same(Row a,Row b) => a.Section==b.Section && a.Number==b.Number && a.Title==b.Title && a.File==b.File && a.Layout==b.Layout;
 [CommandMethod("CODEXSSMRENUMBER",CommandFlags.Session)]
 public void Renumber() {
  AcSmDatabase db=null; bool locked=false; bool committed=false;
  try {
   var plan=JsonSerializer.Deserialize<Plan>(File.ReadAllText(Path.Combine(Folder,"plan.json")));
   IAcSmSheetSetMgr mgr=new AcSmSheetSetMgrClass(); var en=mgr.GetDatabaseEnumerator();
   while((db=en.Next())!=null) if(string.Equals(db.GetFileName(),plan.File,StringComparison.OrdinalIgnoreCase)) break;
   if(db==null) throw new System.Exception("Target sheet set is not open");
   if(db.GetLockStatus()!=AcSmLockStatus.AcSmLockStatus_UnLocked) throw new System.Exception("Sheet set is locked");
   db.LockDb(db); locked=true;
   var before=new List<Row>(); var sheets=new List<IAcSmSheet>(); Walk((IAcSmSubset)db.GetSheetSet(),"",before,sheets);
   if(before.Count!=plan.Sheets.Length || before.Where((r,i)=>!Same(r,plan.Sheets[i])).Any()) throw new System.Exception("Sheet order or properties changed since inspection");
   var backup=Path.Combine(Folder,"backup_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".dst");
   System.IO.File.Copy(db.GetFileName(),backup,false);
   int changed=0;
   for(int i=0;i<sheets.Count;i++) if(sheets[i].GetNumber()!=plan.Sheets[i].NewNumber) {sheets[i].SetNumber(plan.Sheets[i].NewNumber);changed++;}
   var after=new List<Row>(); var afterSheets=new List<IAcSmSheet>(); Walk((IAcSmSubset)db.GetSheetSet(),"",after,afterSheets);
   for(int i=0;i<after.Count;i++) {
    var expected=plan.Sheets[i] with { Number=plan.Sheets[i].NewNumber };
    if(!Same(after[i],expected)) throw new System.Exception("Post-change verification failed at "+i);
   }
   db.UnlockDb(db,true); locked=false; committed=true;
   File.WriteAllText(Path.Combine(Folder,"result.json"),JsonSerializer.Serialize(new {Committed=true,File=db.GetFileName(),Backup=backup,Changed=changed,Sheets=after},new JsonSerializerOptions{WriteIndented=true}));
   AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nCODEX: renumbered and saved; "+changed+" sheet numbers updated.\n");
  } catch(System.Exception ex) {
   if(locked) {try {db.UnlockDb(db,false);} catch(System.Exception rollback) {ex=new AggregateException(ex,rollback);}}
   File.WriteAllText(Path.Combine(Folder,"renumber_error.txt"),"Committed="+committed+"\n"+ex);
  }
 }
}
