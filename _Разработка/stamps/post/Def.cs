using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
public class PostDef {
 static string Dir=Environment.GetEnvironmentVariable("POSTINSP_DIR");
 [CommandMethod("POSTDEF")]public void Run(){var outp=new Dictionary<string,object>();
  foreach(var f in Directory.GetFiles(Dir,"*.dwg"))try{
   using var db=new Database(false,true);db.ReadDwgFile(f,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
   using var tr=db.TransactionManager.StartTransaction();var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
   if(!bt.Has("Рамка со штампом")){outp[Path.GetFileName(f)]="нет блока";continue;}
   var root=(BlockTableRecord)tr.GetObject(bt["Рамка со штампом"],OpenMode.ForRead);var ents=new List<string>();
   foreach(ObjectId id in root){var e=(Entity)tr.GetObject(id,OpenMode.ForRead);string s=e.GetType().Name+" L="+e.Layer;
    if(e is MText m)s+=$" @({m.Location.X:0.#},{m.Location.Y:0.#}) '{m.Contents}'";else if(e is DBText t)s+=$" @({t.Position.X:0.#},{t.Position.Y:0.#}) '{t.TextString}'";
    else if(e is AttributeDefinition ad)s+=$" TAG={ad.Tag} '{ad.TextString}'";else if(e is BlockReference r)s+=" ["+((BlockTableRecord)tr.GetObject(r.BlockTableRecord,OpenMode.ForRead)).Name+"]";
    if(!(e is Line||e is Polyline))ents.Add(s);}
   outp[Path.GetFileName(f)]=new{Count=root.Cast<ObjectId>().Count(),Anon=root.GetAnonymousBlockIds().Count,Refs=root.GetBlockReferenceIds(true,false).Count,Ents=ents};
  }catch(System.Exception e){outp[f]=e.Message;}
  File.WriteAllText(Path.Combine(Dir,"def.json"),JsonSerializer.Serialize(outp,new JsonSerializerOptions{WriteIndented=true,Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping}));}
}
