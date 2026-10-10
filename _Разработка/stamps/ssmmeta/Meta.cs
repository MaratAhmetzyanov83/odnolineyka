using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class StampSsmMeta {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\ssm_task";
 record SheetInfo(string Section,string Number,string Title,string File,string Layout,string Id);
 static void Walk(IAcSmSubset subset,string section,List<SheetInfo> rows,List<IAcSmSheet> sheets) {
  var en=subset.GetSheetEnumerator(); IAcSmComponent item;
  while((item=en.Next())!=null) {
   if(item.GetTypeName()=="AcSmSheet") {
    var s=(IAcSmSheet)item; var l=s.GetLayout(); sheets.Add(s);
    rows.Add(new SheetInfo(section,s.GetNumber(),s.GetTitle(),l.GetFileName(),l.GetName(),s.GetObjectId().GetHandle()));
   } else if(item.GetTypeName()=="AcSmSubset") Walk((IAcSmSubset)item,section+"/"+item.GetName(),rows,sheets);
  }
 }
 [CommandMethod("CODEXSSMSTAMPMETA",CommandFlags.Session)]
 public void Inspect() {
  try {
   IAcSmSheetSetMgr mgr=new AcSmSheetSetMgrClass(); var en=mgr.GetDatabaseEnumerator(); AcSmDatabase db;
   var sets=new List<object>();
   while((db=en.Next())!=null) {
    var rows=new List<SheetInfo>(); var sheets=new List<IAcSmSheet>();
    Walk((IAcSmSubset)db.GetSheetSet(),"",rows,sheets);
    sets.Add(new {File=db.GetFileName(),Title=db.GetSheetSet().GetName(),Id=db.GetSheetSet().GetObjectId().GetHandle(),Lock=db.GetLockStatus().ToString(),Sheets=rows});
   }
   File.WriteAllText(Path.Combine(Folder,"stamp_meta.json"),JsonSerializer.Serialize(sets,new JsonSerializerOptions{WriteIndented=true}));
   AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nCODEX: sheet sets inspected: "+sets.Count+"\n");
  } catch(System.Exception ex) { File.WriteAllText(Path.Combine(Folder,"error.txt"),ex.ToString()); }
 }
}
