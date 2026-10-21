using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
// Ставит определение «Рамка со штампом» из образца во все файлы списка, сохраняя у каждой рамки
// динамические свойства, атрибуты и поля. POSTAPPLY_SRC — образец, POSTAPPLY_LIST — файлы (вход\tвыход).
public class PostApply {
 const string Name="Рамка со штампом",Tmp="Рамка со штампом__из_образца";
 static string FieldCode(DBObject o,Transaction tr){if(!o.HasFields)return null;var f=(Field)tr.GetObject(o.GetField(),OpenMode.ForRead);return f.GetFieldCode(FieldCodeFlags.FieldCode|FieldCodeFlags.AddMarkers);}
 static Dictionary<string,object> Props(BlockReference br){var d=new Dictionary<string,object>();foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(!p.ReadOnly)d[p.PropertyName]=p.Value;return d;}
 static int Count(Database db,string name){using var tr=db.TransactionManager.StartTransaction();var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);return bt.Has(name)?((BlockTableRecord)tr.GetObject(bt[name],OpenMode.ForRead)).Cast<ObjectId>().Count():-1;}
 static object Apply(Database db,Database src){
  var prior=HostApplicationServices.WorkingDatabase;HostApplicationServices.WorkingDatabase=db;try{
  ObjectId sid;var nested=new List<object>();
  using(var tr=src.TransactionManager.StartTransaction()){var bt=(BlockTable)tr.GetObject(src.BlockTableId,OpenMode.ForRead);sid=bt[Tmp];
   foreach(var b in ((BlockTableRecord)tr.GetObject(sid,OpenMode.ForRead)).Cast<ObjectId>().Select(i=>tr.GetObject(i,OpenMode.ForRead)).OfType<BlockReference>().Select(r=>((BlockTableRecord)tr.GetObject(r.BlockTableRecord,OpenMode.ForRead)).Name).Distinct())
    nested.Add(new{Block=b,Src=Count(src,b),Dst=Count(db,b)});}
  var map=new IdMapping();src.WblockCloneObjects(new ObjectIdCollection(new[]{sid}),db.BlockTableId,map,DuplicateRecordCloning.Ignore,false);var newId=map[sid].Value;
  var rows=new List<object>();
  using(var tr=db.TransactionManager.StartTransaction()){
   var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);var oldRoot=(BlockTableRecord)tr.GetObject(bt[Name],OpenMode.ForRead);
   var refs=new List<ObjectId>();foreach(ObjectId id in oldRoot.GetBlockReferenceIds(true,false))refs.Add(id);
   foreach(ObjectId aid in oldRoot.GetAnonymousBlockIds())foreach(ObjectId id in ((BlockTableRecord)tr.GetObject(aid,OpenMode.ForRead)).GetBlockReferenceIds(true,false))refs.Add(id);
   foreach(var id in refs.Distinct()){
    var old=(BlockReference)tr.GetObject(id,OpenMode.ForWrite);var owner=(BlockTableRecord)tr.GetObject(old.OwnerId,OpenMode.ForWrite);
    string layout=((Layout)tr.GetObject(owner.LayoutId,OpenMode.ForRead)).LayoutName;
    var oldProps=Props(old);
    var oldAttrs=old.AttributeCollection.Cast<ObjectId>().Select(a=>(AttributeReference)tr.GetObject(a,OpenMode.ForRead)).ToDictionary(a=>a.Tag,a=>new{Text=a.TextString,Code=FieldCode(a,tr)});
    var br=new BlockReference(old.Position,newId);br.SetDatabaseDefaults(db);br.SetPropertiesFrom(old);br.ScaleFactors=old.ScaleFactors;br.Rotation=old.Rotation;br.Normal=old.Normal;
    owner.AppendEntity(br);tr.AddNewlyCreatedDBObject(br,true);
    var def=(BlockTableRecord)tr.GetObject(newId,OpenMode.ForRead);
    foreach(var ad in def.Cast<ObjectId>().Select(i=>tr.GetObject(i,OpenMode.ForRead)).OfType<AttributeDefinition>().Where(a=>!a.Constant)){var a=new AttributeReference();a.SetAttributeFromBlock(ad,br.BlockTransform);br.AttributeCollection.AppendAttribute(a);tr.AddNewlyCreatedDBObject(a,true);}
    foreach(string n in new[]{"Выбор1","Расстояние Х","Расстояние У","Видимость1"})foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(p.PropertyName==n&&!p.ReadOnly&&oldProps.TryGetValue(n,out var v))p.Value=v;
    var eval=(BlockTableRecord)tr.GetObject(br.BlockTableRecord,OpenMode.ForRead);
    var defs=eval.Cast<ObjectId>().Select(i=>tr.GetObject(i,OpenMode.ForRead)).OfType<AttributeDefinition>().GroupBy(a=>a.Tag).ToDictionary(g=>g.Key,g=>g.First());
    var lost=new List<string>();
    foreach(ObjectId aid in br.AttributeCollection){var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForWrite);if(defs.TryGetValue(a.Tag,out var ad))a.SetAttributeFromBlock(ad,br.BlockTransform);
     if(!oldAttrs.TryGetValue(a.Tag,out var o)){lost.Add("+"+a.Tag);continue;}
     a.TextString=o.Text;
     if(o.Code!=null){var code=o.Code.Replace("_ObjId "+old.ObjectId.OldIdPtr.ToInt64(),"_ObjId "+br.ObjectId.OldIdPtr.ToInt64());var f=new Field(code);a.SetField(f);tr.AddNewlyCreatedDBObject(f,true);
      if(code.Contains("AcObjProp"))f.Evaluate((int)FieldEvaluationContext.Demand,db);}}
    foreach(var t in oldAttrs.Keys)if(!br.AttributeCollection.Cast<ObjectId>().Any(x=>((AttributeReference)tr.GetObject(x,OpenMode.ForRead)).Tag==t))lost.Add("-"+t);
    var dot=(DrawOrderTable)tr.GetObject(owner.DrawOrderTableId,OpenMode.ForWrite);dot.SwapOrder(old.ObjectId,br.ObjectId);
    old.Erase();
    var np=Props(br);var diff=oldProps.Where(kv=>!np.TryGetValue(kv.Key,out var v)||!Equals(Convert.ToString(v),Convert.ToString(kv.Value))).Select(kv=>kv.Key).ToList();
    rows.Add(new{Layout=layout,Old=old.Handle.ToString(),New=br.Handle.ToString(),Props=np.ToDictionary(k=>k.Key,k=>Convert.ToString(k.Value)),PropDiff=diff,AttrDiff=lost});
   }
   var anon=oldRoot.GetAnonymousBlockIds().Cast<ObjectId>().ToList();oldRoot.UpgradeOpen();oldRoot.Erase();
   foreach(var a in anon){var b=(BlockTableRecord)tr.GetObject(a,OpenMode.ForWrite);if(b.GetBlockReferenceIds(true,false).Count==0)b.Erase();}
   var nr=(BlockTableRecord)tr.GetObject(newId,OpenMode.ForWrite);nr.Name=Name;
   tr.Commit();}
  return new{Nested=nested,Frames=rows};}finally{HostApplicationServices.WorkingDatabase=prior;}}
 public record SheetPlan(string Layout,string Number,string Offset);
 const string Num="%<\\AcSm Sheet.Number>%",Title="%<\\AcSm Sheet.Title>%";
 // Поля нумерации: ЛИСТ = номер листа, ОПИСАНИЕ = название, СКВОЗНАЯ = номер + сдвиг (свойство подшивки).
 static object Fields(Database db,SheetPlan[] plan){
  var prior=HostApplicationServices.WorkingDatabase;HostApplicationServices.WorkingDatabase=db;try{var rows=new List<object>();
  using(var tr=db.TransactionManager.StartTransaction()){var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);
   foreach(var sp in plan){if(!ld.Contains(sp.Layout)){rows.Add(new{sp.Layout,Error="нет листа"});continue;}
    var l=(Layout)tr.GetObject(ld.GetAt(sp.Layout),OpenMode.ForRead);int n=0;
    foreach(ObjectId id in (BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead)){
     if(!(tr.GetObject(id,OpenMode.ForRead) is BlockReference br))continue;var root=(BlockTableRecord)tr.GetObject(br.IsDynamicBlock?br.DynamicBlockTableRecord:br.BlockTableRecord,OpenMode.ForRead);if(root.Name!=Name)continue;n++;
     var set=new Dictionary<string,string>();
     foreach(ObjectId aid in br.AttributeCollection){var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForWrite);string code=null;
      if(a.Tag=="ЛИСТ_№"||a.Tag=="ЛИСТ№")code=Num;else if(a.Tag=="ОПИСАНИЕ_ЛИСТА")code=Title;
      else if(a.Tag=="СКВОЗНАЯ_НУМЕРЦИЯ")code=sp.Number==""?"%<\\AcSm SheetSet."+sp.Offset+">%":"%<\\AcExpr (%<\\AcSm Sheet.Number \\f \"%tc4\">%+%<\\AcSm SheetSet."+sp.Offset+" \\f \"%tc4\">%)>%";
      if(code==null)continue;if(a.HasFields)a.RemoveField();var f=new Field(code);a.SetField(f);tr.AddNewlyCreatedDBObject(f,true);set[a.Tag]=FieldCode(a,tr);}
     rows.Add(new{sp.Layout,Frame=br.Handle.ToString(),Codes=set});}
    if(n==0)rows.Add(new{sp.Layout,Error="нет рамки"});}
   tr.Commit();}
  return rows;}finally{HostApplicationServices.WorkingDatabase=prior;}}
 [CommandMethod("POSTAPPLY")]public void Run(){
  string srcFile=Environment.GetEnvironmentVariable("POSTAPPLY_SRC"),list=Environment.GetEnvironmentVariable("POSTAPPLY_LIST");var res=new List<object>();
  using var src=new Database(false,true);src.ReadDwgFile(srcFile,FileOpenMode.OpenForReadAndAllShare,false,null);src.CloseInput(true);
  using(var tr=src.TransactionManager.StartTransaction()){var bt=(BlockTable)tr.GetObject(src.BlockTableId,OpenMode.ForRead);((BlockTableRecord)tr.GetObject(bt[Name],OpenMode.ForWrite)).Name=Tmp;tr.Commit();}
  foreach(var line in File.ReadAllLines(list).Where(x=>x.Contains('\t'))){var io=line.Split('\t');
   try{using var db=new Database(false,true);db.ReadDwgFile(io[0],FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);var ver=db.OriginalFileVersion;
    object r=io[2]=="1"?Apply(db,src):null;object fl=io.Length>3?Fields(db,JsonSerializer.Deserialize<SheetPlan[]>(io[3])):null;db.SaveAs(io[1],ver);res.Add(new{File=io[0],Out=io[1],Version=ver.ToString(),Result=r,Fields=fl});}
   catch(System.Exception e){res.Add(new{File=io[0],Error=e.ToString()});}
   File.WriteAllText(list+".result.json",JsonSerializer.Serialize(res,new JsonSerializerOptions{WriteIndented=true,Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping}));}
 }
}
