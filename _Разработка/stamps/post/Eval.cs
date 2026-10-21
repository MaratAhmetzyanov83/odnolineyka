using System;using System.IO;using System.Linq;using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
public class PostEval {
 [CommandMethod("POSTEVAL")]public void Run(){var f=Environment.GetEnvironmentVariable("POSTEVAL_FILE");var o=new List<string>();
  using var db=new Database(false,true);db.ReadDwgFile(f,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
  var prior=HostApplicationServices.WorkingDatabase;HostApplicationServices.WorkingDatabase=db;
  try{db.EvaluateFields();}catch(System.Exception e){o.Add("EVAL "+e.Message);}
  using(var tr=db.TransactionManager.StartTransaction()){var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);
   foreach(DBDictionaryEntry de in ld){var l=(Layout)tr.GetObject(de.Value,OpenMode.ForRead);if(l.ModelType)continue;
    foreach(ObjectId id in (BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead))if(tr.GetObject(id,OpenMode.ForRead) is BlockReference br)foreach(ObjectId a in br.AttributeCollection){var at=(AttributeReference)tr.GetObject(a,OpenMode.ForRead);if(at.Tag=="ЛИСТ_№")o.Add(l.LayoutName+" "+at.TextString);}}}
  HostApplicationServices.WorkingDatabase=prior;File.WriteAllLines(f+".eval.txt",o);}
}
