using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
// Возвращает внешнюю ссылку «Штамп» внутрь блока «Штамп_шифр_ПО» (FixCode подменил ею xref и потерял графику штампа).
// POSTAPPLY_LIST — строки «вход\tвыход».
public class PostXref {
 [CommandMethod("POSTXREF")]public void Run(){var list=Environment.GetEnvironmentVariable("POSTAPPLY_LIST");var res=new List<object>();
  foreach(var line in File.ReadAllLines(list).Where(x=>x.Contains('\t'))){var io=line.Split('\t');var log=new List<string>();
   try{using var db=new Database(false,true);db.ReadDwgFile(io[0],FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);var ver=db.OriginalFileVersion;
    using(var tr=db.TransactionManager.StartTransaction()){var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
     var x=(BlockTableRecord)tr.GetObject(bt["Штамп"],OpenMode.ForRead);log.Add("Штамп: xref="+x.IsFromExternalReference+" path="+x.PathName+" origin="+x.Origin);
     var want=Environment.GetEnvironmentVariable("POSTXREF_PATH");if(!string.IsNullOrEmpty(want)&&x.IsFromExternalReference&&!string.Equals(x.PathName,want,StringComparison.OrdinalIgnoreCase)){x.UpgradeOpen();x.PathName=want;log.Add("путь xref → "+want);}
     foreach(ObjectId id in bt){var b=(BlockTableRecord)tr.GetObject(id,OpenMode.ForRead);if(!b.Name.StartsWith("Штамп_шифр_ПО"))continue;
      if(b.Cast<ObjectId>().Any(e=>tr.GetObject(e,OpenMode.ForRead) is BlockReference)){log.Add(b.Name+": xref уже есть");continue;}
      if(b.Origin!=x.Origin)throw new System.Exception("Разные базы: "+b.Origin+" / "+x.Origin);
      b.UpgradeOpen();var r=new BlockReference(b.Origin,x.ObjectId);r.SetDatabaseDefaults(db);r.Layer="0";b.AppendEntity(r);tr.AddNewlyCreatedDBObject(r,true);
      log.Add(b.Name+": xref добавлен, ссылок на блок "+b.GetBlockReferenceIds(true,false).Count);}
     tr.Commit();}
    db.SaveAs(io[1],ver);res.Add(new{File=io[1],Log=log});}
   catch(System.Exception e){res.Add(new{File=io[0],Log=log,Error=e.ToString()});}
   File.WriteAllText(list+".xref.json",JsonSerializer.Serialize(res,new JsonSerializerOptions{WriteIndented=true,Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping}));}}
}
