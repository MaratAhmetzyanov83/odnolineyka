using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using AcApp=Autodesk.AutoCAD.ApplicationServices.Core.Application;
public class StampApply {
 const string Folder=@"C:\Users\user\Desktop\02_Справочники_CAD_Revit\odnolineyka\_Разработка\stamps";
 const string RefDir=@"\\Server\DropBox\Работа server\4. Техдепартамент\4.1. Проекты\098 Ташкент-КТ\5 ПИР\5.5_Том 4 Пост охраны\01_ELF-D1-03-ЭОМ";
 public record Item(string Layout,string Handle,string Variant,string Title);
 public record FilePlan(string File,Item[] Items);
 public record MetaSheet(string File,string Layout,string Number,string Title,string Id);
 public record MetaSet(string File,string Title,string Id,MetaSheet[] Sheets);
 static MetaSet Metadata=JsonSerializer.Deserialize<MetaSet[]>(File.ReadAllText(Path.Combine(Path.GetDirectoryName(Folder),"ssm_task","stamp_meta.json"))).First(x=>x.Title.Contains("02-ЭОМ"));
 static string SheetCode(MetaSheet sheet,string property)=>"%<\\AcSm Database(\""+Metadata.File+"\").SheetSet(\""+Metadata.Id+"\").Component(\""+sheet.Id+"\")."+property+">%";
 static string FieldCode(DBObject obj,Transaction tr){if(!obj.HasFields)return null;var f=(Field)tr.GetObject(obj.GetField(),OpenMode.ForRead);return f.GetFieldCode(FieldCodeFlags.FieldCode|FieldCodeFlags.AddMarkers);}
 static Database Source(string variant){
  var src=new Database(false,true);src.ReadDwgFile(Path.Combine(RefDir,variant=="Схемы"?"02_Однолинейка действующая.dwg":"03_Чертежи ЭОМ_Пост охраны.dwg"),FileOpenMode.OpenForReadAndAllShare,false,null);src.CloseInput(true);
  ObjectId rootId;
  using(var tr=src.TransactionManager.StartTransaction()){
   var bt=(BlockTable)tr.GetObject(src.BlockTableId,OpenMode.ForRead);rootId=bt["Рамка со штампом"];
   var root=(BlockTableRecord)tr.GetObject(rootId,OpenMode.ForWrite);root.Name="Рамка_ХБК_образец_"+variant;
   var childIds=root.Cast<ObjectId>().Select(id=>tr.GetObject(id,OpenMode.ForRead)).OfType<BlockReference>().Select(b=>b.BlockTableRecord).Distinct().ToArray();
   foreach(var cid in childIds){var child=(BlockTableRecord)tr.GetObject(cid,OpenMode.ForWrite);if(!child.IsAnonymous && !child.IsLayout)child.Name=child.Name+"_ХБК_образец_"+variant;}
   foreach(ObjectId id in root){if(!(tr.GetObject(id,OpenMode.ForRead) is MText m))continue;
    if(m.Text.Contains("ELF-D1-")){m.UpgradeOpen();m.Contents="ELF-D1-02-ЭОМ";}
    else if(m.Text.Trim()=="D100 КПП-1" || m.Text.Trim()=="D100"){m.UpgradeOpen();m.Contents="D100 ХБК";}
   }tr.Commit();
  }
  using(var tr=src.TransactionManager.StartTransaction()){((BlockTableRecord)tr.GetObject(rootId,OpenMode.ForWrite)).UpdateAnonymousBlocks();tr.Commit();}
  return src;
 }
 static List<object> Apply(Database db,FilePlan plan,Dictionary<string,Database> sources){
  var priorWorking=HostApplicationServices.WorkingDatabase;HostApplicationServices.WorkingDatabase=db;try{
  var imported=new Dictionary<string,ObjectId>();
  foreach(string variant in plan.Items.Select(i=>i.Variant).Distinct()){
   var src=sources[variant];ObjectId sid;
   using(var tr=src.TransactionManager.StartTransaction()){var bt=(BlockTable)tr.GetObject(src.BlockTableId,OpenMode.ForRead);sid=bt["Рамка_ХБК_образец_"+variant];}
   var map=new IdMapping();src.WblockCloneObjects(new ObjectIdCollection(new[]{sid}),db.BlockTableId,map,DuplicateRecordCloning.MangleName,false);imported[variant]=map[sid].Value;
  }
  var results=new List<object>();
  using(var tr=db.TransactionManager.StartTransaction()){
   foreach(var item in plan.Items){var br=(BlockReference)tr.GetObject(db.GetObjectId(false,new Handle(Convert.ToInt64(item.Handle,16)),0),OpenMode.ForWrite);
    var root=(BlockTableRecord)tr.GetObject(br.IsDynamicBlock?br.DynamicBlockTableRecord:br.BlockTableRecord,OpenMode.ForRead);
    if(!root.Name.StartsWith("Рамка со штампом"))throw new System.Exception("Unexpected original frame: "+item.Layout);
    var previous=br;
    var oldProperties=new Dictionary<string,object>();foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(!p.ReadOnly)oldProperties[p.PropertyName]=p.Value;
    var oldAttrs=br.AttributeCollection.Cast<ObjectId>().Select(id=>(AttributeReference)tr.GetObject(id,OpenMode.ForRead)).ToDictionary(a=>a.Tag,a=>new{Text=a.TextString,Code=FieldCode(a,tr)});
    var originalPosition=br.Position;var originalScale=br.ScaleFactors;var originalRotation=br.Rotation;
    var owner=(BlockTableRecord)tr.GetObject(br.OwnerId,OpenMode.ForWrite);
    br=new BlockReference(originalPosition,imported[item.Variant]);br.SetDatabaseDefaults(db);br.ScaleFactors=originalScale;br.Rotation=originalRotation;br.LayerId=previous.LayerId;
    owner.AppendEntity(br);tr.AddNewlyCreatedDBObject(br,true);
    foreach(ObjectId aid in previous.AttributeCollection){var old=(AttributeReference)tr.GetObject(aid,OpenMode.ForRead);var a=new AttributeReference();a.SetDatabaseDefaults(db);a.Tag=old.Tag;a.TextString=old.TextString;br.AttributeCollection.AppendAttribute(a);tr.AddNewlyCreatedDBObject(a,true);}
    previous.Erase();
    if(!br.IsDynamicBlock)throw new System.Exception("Replacement lost dynamic behavior: "+item.Layout);
    foreach(string name in new[]{"Выбор1","Расстояние Х","Расстояние У","Видимость1"}){
     foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection){if(p.PropertyName!=name||p.ReadOnly)continue;if(name=="Видимость1")p.Value="Форма 3";else if(oldProperties.TryGetValue(name,out var value))p.Value=value;}
    }
    var evaluated=(BlockTableRecord)tr.GetObject(br.BlockTableRecord,OpenMode.ForRead);
    var defs=evaluated.Cast<ObjectId>().Select(id=>tr.GetObject(id,OpenMode.ForRead)).OfType<AttributeDefinition>().ToDictionary(a=>a.Tag);
    var sheet=Metadata.Sheets.FirstOrDefault(x=>string.Equals(x.File,plan.File,StringComparison.OrdinalIgnoreCase)&&x.Layout==item.Layout);
    if(sheet==null && item.Layout=="ВК цоколь")sheet=Metadata.Sheets.First(x=>x.Layout=="ВК цоколь");
    var codes=new Dictionary<string,string>();
    foreach(ObjectId aid in br.AttributeCollection){var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForWrite);
     var tag=a.Tag;var old=oldAttrs[tag];
     if(defs.TryGetValue(tag,out var ad))a.SetAttributeFromBlock(ad,br.BlockTransform);
     var code=old.Code;
     if(sheet!=null && (tag=="ЛИСТ_№" || tag=="ЛИСТ№" || tag=="СКВОЗНАЯ_НУМЕРЦИЯ"))code=SheetCode(sheet,"Number");
     if(sheet!=null && tag=="ОПИСАНИЕ_ЛИСТА")code=SheetCode(sheet,"Title");
     if(tag=="ФОРМАТ_БУМАГИ" && code!=null)code=code.Replace(previous.ObjectId.OldIdPtr.ToInt64().ToString(),br.ObjectId.OldIdPtr.ToInt64().ToString());
     if(code!=null){var nf=new Field(code);a.SetField(nf);tr.AddNewlyCreatedDBObject(nf,true);nf.Evaluate((int)FieldEvaluationContext.Demand,db);codes[tag]=FieldCode(a,tr);}
     else a.TextString=old.Text;
     if(code!=null && FieldCode(a,tr)!=code)throw new System.Exception("Field was not retained: "+item.Layout+" "+tag);
    }
    if(br.Position!=originalPosition || br.ScaleFactors!=originalScale || br.Rotation!=originalRotation)throw new System.Exception("Frame transform changed");
    var constants=evaluated.Cast<ObjectId>().Select(id=>tr.GetObject(id,OpenMode.ForRead)).OfType<MText>().Select(m=>m.Text).ToArray();
    var newProps=new Dictionary<string,object>();foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(!p.ReadOnly)newProps[p.PropertyName]=p.Value;
    if((string)newProps["Видимость1"]!="Форма 3" || !Equals(newProps["Выбор1"],oldProperties["Выбор1"]))throw new System.Exception("Wrong form/format");
    if(!constants.Any(t=>t.Contains("ELF-D1-02-ЭОМ")))throw new System.Exception("Wrong project code: "+item.Layout);
    results.Add(new{Layout=item.Layout,Handle=br.Handle.ToString(),OldHandle=item.Handle,Variant=item.Variant,Properties=newProps,Attributes=oldAttrs,FieldCodes=codes,SheetNumber=sheet?.Number,SheetId=sheet?.Id,Constants=constants});
   }tr.Commit();
  }return results;}finally{HostApplicationServices.WorkingDatabase=priorWorking;}
 }
 [CommandMethod("CODEXSTAMPPREVIEW13",CommandFlags.Session)]public void Preview(){
  try{
   var plans=JsonSerializer.Deserialize<FilePlan[]>(File.ReadAllText(Path.Combine(Folder,"plan.json")));
   var selected=plans.Where(p=>Path.GetFileName(p.File).Equals("03_Чертежи ЭОМ_ХБК.dwg",StringComparison.OrdinalIgnoreCase)||Path.GetFileName(p.File).Equals("Однолинейка ЩР-1.dwg",StringComparison.OrdinalIgnoreCase)).ToArray();
   var sources=new Dictionary<string,Database>{{"Схемы",Source("Схемы")},{"Общий",Source("Общий")}};var results=new List<object>();
   try{foreach(var plan in selected){using(var db=new Database(false,true)){
    db.ReadDwgFile(plan.File,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
    var subset=plan with{Items=new[]{plan.Items.First(i=>i.Variant==(Path.GetFileName(plan.File).StartsWith("Однолинейка")?"Схемы":"Общий"))}};
    var rows=Apply(db,subset,sources);var preview=Path.Combine(Folder,"PREVIEW_"+Path.GetFileName(plan.File));db.SaveAs(preview,DwgVersion.Current);results.Add(new{File=preview,Rows=rows});
   }}File.WriteAllText(Path.Combine(Folder,"preview13.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));}
   finally{foreach(var db in sources.Values)db.Dispose();}
  }catch(System.Exception ex){File.WriteAllText(Path.Combine(Folder,"preview_error13.txt"),ex.ToString());}
 }
 [CommandMethod("CODEXSTAMPAPPLY13",CommandFlags.Session)]public void Execute(){
  var completed=new List<object>();string runDir=Path.Combine(Folder,"backup_"+DateTime.Now.ToString("yyyyMMdd_HHmmss"));Directory.CreateDirectory(runDir);
  try{
   var plans=JsonSerializer.Deserialize<FilePlan[]>(File.ReadAllText(Path.Combine(Folder,"plan.json")));
   var sources=new Dictionary<string,Database>{{"Схемы",Source("Схемы")},{"Общий",Source("Общий")}};
   try{int n=0;foreach(var plan in plans){n++;var backup=Path.Combine(runDir,n.ToString("00")+"_"+Path.GetFileName(plan.File));
    var doc=AcApp.DocumentManager.Cast<Autodesk.AutoCAD.ApplicationServices.Document>().FirstOrDefault(d=>string.Equals(d.Name,plan.File,StringComparison.OrdinalIgnoreCase));
    List<object> rows;
    if(doc!=null){using(doc.LockDocument()){
     using(var copy=doc.Database.Wblock())copy.SaveAs(backup,DwgVersion.Current);
     rows=Apply(doc.Database,plan,sources);((dynamic)doc.GetAcadDocument()).Save();doc.Editor.Regen();
    }}else{
     var before=File.GetLastWriteTimeUtc(plan.File);
     using(var db=new Database(false,true)){
      db.ReadDwgFile(plan.File,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);File.Copy(plan.File,backup,false);
      rows=Apply(db,plan,sources);
      if(File.GetLastWriteTimeUtc(plan.File)!=before)throw new System.Exception("File changed while being processed: "+plan.File);
      db.SaveAs(plan.File,DwgVersion.Current);
     }
    }
    completed.Add(new{File=plan.File,Backup=backup,Rows=rows,Saved=true});
    File.WriteAllText(Path.Combine(Folder,"result.json"),JsonSerializer.Serialize(new{Complete=false,Completed=completed},new JsonSerializerOptions{WriteIndented=true}));
   }
   File.WriteAllText(Path.Combine(Folder,"result.json"),JsonSerializer.Serialize(new{Complete=true,Completed=completed},new JsonSerializerOptions{WriteIndented=true}));}
   finally{foreach(var db in sources.Values)db.Dispose();}
  }catch(System.Exception ex){File.WriteAllText(Path.Combine(Folder,"apply_error13.txt"),ex.ToString());}
 }
}
