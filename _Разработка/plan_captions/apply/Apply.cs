using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class ApplyPlanCaptions {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\plan_captions";
 const string RefFile=@"\\Server\DropBox\Работа server\4. Техдепартамент\4.1. Проекты\098 Ташкент-КТ\5 ПИР\5.5_Том 4 Пост охраны\01_ELF-D1-03-ЭОМ\03_Чертежи ЭОМ_Пост охраны.dwg";
 public record PlanItem(string Layout,string Viewport,string Title,double Scale);
 [CommandMethod("CODEXAPPLYCAPTIONS",CommandFlags.Session)]
 public void Apply() {
  try {
   var items=JsonSerializer.Deserialize<PlanItem[]>(File.ReadAllText(Path.Combine(Folder,"plan.json")));
   var docs=AcApp.DocumentManager.Cast<Autodesk.AutoCAD.ApplicationServices.Document>().Where(d=>Path.GetFileName(d.Name).Equals("03_Чертежи ЭОМ_ХБК.dwg",StringComparison.OrdinalIgnoreCase)).ToArray();
   if(docs.Length!=1)throw new System.Exception("Expected one target document"); var doc=docs[0];var db=doc.Database;
   using(doc.LockDocument()) {
    string backup=Path.Combine(Folder,"ХБК_до_подписей_"+DateTime.Now.ToString("yyyyMMdd_HHmmss")+".dwg");
    using(var copy=db.Wblock())copy.SaveAs(backup,DwgVersion.Current);
    using(var src=new Database(false,true)) {
     src.ReadDwgFile(RefFile,FileOpenMode.OpenForReadAndAllShare,false,null);src.CloseInput(true);
     var sourceId=src.GetObjectId(false,new Handle(Convert.ToInt64("1DC21E",16)),0);
     var results=new List<object>();
     using(var tr=db.TransactionManager.StartTransaction()) {
      var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);
      foreach(var item in items) {
       var layout=(Layout)tr.GetObject(ld.GetAt(item.Layout),OpenMode.ForRead); var btr=(BlockTableRecord)tr.GetObject(layout.BlockTableRecordId,OpenMode.ForWrite);
       var vid=db.GetObjectId(false,new Handle(Convert.ToInt64(item.Viewport,16)),0); var v=(Viewport)tr.GetObject(vid,OpenMode.ForRead);
       if(v.OwnerId!=btr.ObjectId || Math.Abs(v.CustomScale-item.Scale)>1e-8)throw new System.Exception("Viewport changed: "+item.Layout);
       var existing=btr.Cast<ObjectId>().Select(id=>tr.GetObject(id,OpenMode.ForRead)).OfType<MText>().Where(m=>m.Text.StartsWith("План ") && m.Location.Y>v.CenterPoint.Y+v.Height/2-40).ToArray();
       if(existing.Length>0)throw new System.Exception("Heading already exists: "+item.Layout);
       var ids=new ObjectIdCollection(new[]{sourceId});var map=new IdMapping();src.WblockCloneObjects(ids,btr.ObjectId,map,DuplicateRecordCloning.Ignore,false);
       var caption=(MText)tr.GetObject(map[sourceId].Value,OpenMode.ForWrite);
       var denominator=1.0/v.CustomScale;var scale=Math.Round(denominator).ToString(System.Globalization.CultureInfo.InvariantCulture);
       caption.Contents="\\pxqc;"+item.Title+" (1:"+scale+")";
       caption.Attachment=AttachmentPoint.TopCenter;caption.TextHeight=5;caption.Width=v.Width-40;caption.Rotation=0;
       caption.Location=new Point3d(v.CenterPoint.X,v.CenterPoint.Y+v.Height/2-18.65884484580073,0);
       caption.Layer="0";caption.ColorIndex=256;
       var ext=caption.GeometricExtents;
       if(ext.MinPoint.X<v.CenterPoint.X-v.Width/2 || ext.MaxPoint.X>v.CenterPoint.X+v.Width/2 || ext.MaxPoint.Y>v.CenterPoint.Y+v.Height/2)throw new System.Exception("Caption exceeds viewport: "+item.Layout);
       results.Add(new{Layout=item.Layout,Handle=caption.Handle.ToString(),Text=caption.Text,Scale=v.CustomScale,Height=caption.TextHeight,Style=((TextStyleTableRecord)tr.GetObject(caption.TextStyleId,OpenMode.ForRead)).Name,Position=new[]{caption.Location.X,caption.Location.Y}});
      }
      tr.Commit();
     }
     doc.Editor.Regen();
     File.WriteAllText(Path.Combine(Folder,"applied.json"),JsonSerializer.Serialize(new{Applied=true,Backup=backup,File=doc.Name,Captions=results},new JsonSerializerOptions{WriteIndented=true}));
    }
   }
  }catch(System.Exception ex){File.WriteAllText(Path.Combine(Folder,"apply_error.txt"),ex.ToString());}
 }
}
