using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
public class PostTexts {
 static string Dir=Environment.GetEnvironmentVariable("POSTINSP_DIR");
 static void Collect(BlockTableRecord b,Transaction tr,List<string> o,int depth){if(depth>4)return;
  foreach(ObjectId id in b){var e=tr.GetObject(id,OpenMode.ForRead);
   if(e is MText m)o.Add(m.Text.Replace("\n"," | "));else if(e is DBText t)o.Add(t.TextString);
   else if(e is BlockReference r){var c=(BlockTableRecord)tr.GetObject(r.BlockTableRecord,OpenMode.ForRead);o.Add("[блок "+(r.IsDynamicBlock?((BlockTableRecord)tr.GetObject(r.DynamicBlockTableRecord,OpenMode.ForRead)).Name:c.Name)+(c.IsFromExternalReference?" XREF "+c.PathName:"")+"]");if(!c.IsFromExternalReference)Collect(c,tr,o,depth+1);}}}
 [CommandMethod("POSTTEXTS")]public void Run(){var outp=new Dictionary<string,object>();
  foreach(var f in File.ReadAllLines(Path.Combine(Dir,"files.txt")).Where(x=>x.Trim()!=""))try{
   using var db=new Database(false,true);db.ReadDwgFile(f,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
   using var tr=db.TransactionManager.StartTransaction();var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);var per=new Dictionary<string,object>();
   var info=db.SummaryInfo;var cp=new Dictionary<string,string>();var en=info.CustomProperties;while(en.MoveNext())cp[(string)en.Key]=(string)en.Value;per["_CustomDP"]=cp;
   foreach(DBDictionaryEntry de in ld){var l=(Layout)tr.GetObject(de.Value,OpenMode.ForRead);if(l.ModelType)continue;var btr=(BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead);
    var frames=new List<object>();
    foreach(ObjectId id in btr){if(!(tr.GetObject(id,OpenMode.ForRead) is BlockReference br))continue;var root=(BlockTableRecord)tr.GetObject(br.IsDynamicBlock?br.DynamicBlockTableRecord:br.BlockTableRecord,OpenMode.ForRead);
     if(root.Name!="Формат"&&!root.Name.StartsWith("Рамка"))continue;var o=new List<string>();Collect((BlockTableRecord)tr.GetObject(br.BlockTableRecord,OpenMode.ForRead),tr,o,0);frames.Add(new{root.Name,Texts=o});}
    per[l.LayoutName]=frames;}
   outp[Path.GetFileName(f)]=per;}catch(System.Exception e){outp[f]=e.Message;}
  File.WriteAllText(Path.Combine(Dir,"texts.json"),JsonSerializer.Serialize(outp,new JsonSerializerOptions{WriteIndented=true,Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping}));}
}
