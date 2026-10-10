using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;using Autodesk.AutoCAD.Geometry;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class InspectStamps {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\stamps";
 static double[] Pt(Point3d p)=>new[]{p.X,p.Y,p.Z};
 static string Code(DBObject o,Transaction tr){if(!o.HasFields)return null;var f=(Field)tr.GetObject(o.GetField(),OpenMode.ForRead);return f.GetFieldCode(FieldCodeFlags.FieldCode|FieldCodeFlags.AddMarkers);}
 static object Dump(Database db) {
  using(var tr=db.TransactionManager.StartTransaction()){
   var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);var rows=new List<object>();
   foreach(DBDictionaryEntry de in ld){var l=(Layout)tr.GetObject(de.Value,OpenMode.ForRead);if(l.ModelType)continue;var btr=(BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead);
    foreach(ObjectId id in btr){if(!(tr.GetObject(id,OpenMode.ForRead) is BlockReference br))continue;
     var root=(BlockTableRecord)tr.GetObject(br.IsDynamicBlock?br.DynamicBlockTableRecord:br.BlockTableRecord,OpenMode.ForRead);
     if(!root.Name.StartsWith("Рамка со штампом"))continue;
     var props=new Dictionary<string,object>();if(br.IsDynamicBlock)foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(!p.ReadOnly)props[p.PropertyName]=p.Value;
     var attrs=new List<object>();foreach(ObjectId aid in br.AttributeCollection){var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForRead);attrs.Add(new{Handle=a.Handle.ToString(),Tag=a.Tag,Text=a.TextString,Field=Code(a,tr),Height=a.Height,Position=Pt(a.Position)});}
     var constants=new List<object>();var eval=(BlockTableRecord)tr.GetObject(br.BlockTableRecord,OpenMode.ForRead);
     foreach(ObjectId eid in eval){var e=tr.GetObject(eid,OpenMode.ForRead);if(e is MText m)constants.Add(new{Handle=m.Handle.ToString(),Text=m.Text,Contents=m.Contents,Position=Pt(m.Location),Height=m.TextHeight,Width=m.Width});}
     rows.Add(new{Layout=l.LayoutName,Handle=br.Handle.ToString(),Root=root.Name,RootHandle=root.Handle.ToString(),Evaluated=eval.Name,Position=Pt(br.Position),Scale=new[]{br.ScaleFactors.X,br.ScaleFactors.Y,br.ScaleFactors.Z},Properties=props,Attributes=attrs,Constants=constants});
    }
   }return new{File=db.Filename,Frames=rows};
  }
 }
 [CommandMethod("CODEXSTAMPINSPECT",CommandFlags.Session)]public void Inspect(){
  try{
   var files=JsonSerializer.Deserialize<string[]>(File.ReadAllText(Path.Combine(Folder,"files.json")));var results=new List<object>();
   var docs=AcApp.DocumentManager.Cast<Autodesk.AutoCAD.ApplicationServices.Document>().ToArray();
   foreach(var file in files){var doc=docs.FirstOrDefault(d=>string.Equals(d.Name,file,StringComparison.OrdinalIgnoreCase));
    if(doc!=null)results.Add(Dump(doc.Database));else using(var db=new Database(false,true)){db.ReadDwgFile(file,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);results.Add(Dump(db));}}
   File.WriteAllText(Path.Combine(Folder,"inventory.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
  }catch(System.Exception e){File.WriteAllText(Path.Combine(Folder,"error.txt"),e.ToString());}
 }
}
