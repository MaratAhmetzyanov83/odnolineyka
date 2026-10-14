using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class CaptionInspection {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\plan_captions";
 const string RefFile=@"\\Server\DropBox\Работа server\4. Техдепартамент\4.1. Проекты\098 Ташкент-КТ\5 ПИР\5.5_Том 4 Пост охраны\01_ELF-D1-03-ЭОМ\03_Чертежи ЭОМ_Пост охраны.dwg";
 static double[] Pt(Point3d p)=>new[]{p.X,p.Y,p.Z};
 static double[] Pt(Vector3d p)=>new[]{p.X,p.Y,p.Z};
 static object Bounds(Entity e) {try {var b=e.GeometricExtents; return new {Min=Pt(b.MinPoint),Max=Pt(b.MaxPoint)};}catch{return null;}}
 static object Dump(Database db) {
  using(var tr=db.TransactionManager.StartTransaction()) {
   var layouts=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead); var list=new List<object>();
   foreach(DBDictionaryEntry entry in layouts) {
    var l=(Layout)tr.GetObject(entry.Value,OpenMode.ForRead); var btr=(BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead); var ents=new List<object>();
    foreach(ObjectId id in btr) {
     var e=tr.GetObject(id,OpenMode.ForRead) as Entity; if(e==null)continue;
     object detail=null;
     if(e is DBText t) detail=new {Text=t.TextString,Position=Pt(t.Position),Alignment=Pt(t.AlignmentPoint),Height=t.Height,WidthFactor=t.WidthFactor,Rotation=t.Rotation,Style=((TextStyleTableRecord)tr.GetObject(t.TextStyleId,OpenMode.ForRead)).Name,Justify=t.Justify.ToString(),StyleHandle=t.TextStyleId.Handle.ToString()};
     if(e is MText m) detail=new {Text=m.Contents,Plain=m.Text,Position=Pt(m.Location),Height=m.TextHeight,Width=m.Width,Rotation=m.Rotation,Style=((TextStyleTableRecord)tr.GetObject(m.TextStyleId,OpenMode.ForRead)).Name,Attachment=m.Attachment.ToString(),StyleHandle=m.TextStyleId.Handle.ToString()};
     if(e is Viewport v) detail=new {Number=v.Number,Center=Pt(v.CenterPoint),Width=v.Width,Height=v.Height,CustomScale=v.CustomScale,ViewCenter=new[]{v.ViewCenter.X,v.ViewCenter.Y},ViewHeight=v.ViewHeight,ViewTarget=Pt(v.ViewTarget),ViewDirection=Pt(v.ViewDirection),Twist=v.TwistAngle,Locked=v.Locked,On=v.On};
     if(e is BlockReference br) {
      var attrs=new List<object>(); foreach(ObjectId aid in br.AttributeCollection) {var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForRead);attrs.Add(new{Handle=a.Handle.ToString(),Tag=a.Tag,Text=a.TextString,Position=Pt(a.Position),Height=a.Height,Field=a.HasFields});}
      if(l.ModelType && attrs.Count==0)continue;
      detail=new {Name=((BlockTableRecord)tr.GetObject(br.BlockTableRecord,OpenMode.ForRead)).Name,Position=Pt(br.Position),Scale=new[]{br.ScaleFactors.X,br.ScaleFactors.Y,br.ScaleFactors.Z},Rotation=br.Rotation,Attributes=attrs};
     }
     if(detail!=null) ents.Add(new {Handle=e.Handle.ToString(),Type=e.GetType().Name,Layer=e.Layer,Color=e.ColorIndex,LineWeight=e.LineWeight.ToString(),Field=e.HasFields,Bounds=Bounds(e),Detail=detail});
    }
    list.Add(new{Name=l.LayoutName,Model=l.ModelType,TabOrder=l.TabOrder,Entities=ents});
   }
   return new {File=db.Filename,Layouts=list};
  }
 }
 [CommandMethod("CODEXCAPTIONINSPECT",CommandFlags.Session)]
 public void Inspect() {
  try {
   var docs=AcApp.DocumentManager.Cast<Autodesk.AutoCAD.ApplicationServices.Document>().Where(d=>Path.GetFileName(d.Name).Equals("03_Чертежи ЭОМ_ХБК.dwg",StringComparison.OrdinalIgnoreCase)).ToArray();
   if(docs.Length!=1)throw new System.Exception("Expected one target drawing");
   var target=Dump(docs[0].Database);
   using(var db=new Database(false,true)) {db.ReadDwgFile(RefFile,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);var reference=Dump(db);File.WriteAllText(Path.Combine(Folder,"inspection.json"),JsonSerializer.Serialize(new {Target=target,Reference=reference},new JsonSerializerOptions{WriteIndented=true}));}
  }catch(System.Exception ex) {File.WriteAllText(Path.Combine(Folder,"error.txt"),ex.ToString());}
 }
}
