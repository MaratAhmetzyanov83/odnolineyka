using System;using System.IO;using System.Linq;using System.Collections.Generic;using System.Text.Json;
using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;using Autodesk.AutoCAD.Interop;
public class PostInsp {
 static string Dir=Environment.GetEnvironmentVariable("POSTINSP_DIR");
 static string Code(DBObject o,Transaction tr){if(!o.HasFields)return null;try{var f=(Field)tr.GetObject(o.GetField(),OpenMode.ForRead);return f.GetFieldCode(FieldCodeFlags.FieldCode|FieldCodeFlags.AddMarkers);}catch{return "?";}}
 static void Walk(IAcSmSubset s,string path,List<object> rows){var en=s.GetSheetEnumerator();IAcSmComponent c;
  while((c=en.Next())!=null){
   if(c.GetTypeName()=="AcSmSheet"){var sh=(IAcSmSheet)c;var l=sh.GetLayout();rows.Add(new{Path=path,Number=sh.GetNumber(),Title=sh.GetTitle(),File=l.GetFileName(),Layout=l.GetName(),Id=sh.GetObjectId().GetHandle()});}
   else if(c.GetTypeName()=="AcSmSubset")Walk((IAcSmSubset)c,path+"/"+c.GetName(),rows);}}
 static object Dump(string file){using var db=new Database(false,true);db.ReadDwgFile(file,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
  using var tr=db.TransactionManager.StartTransaction();var ld=(DBDictionary)tr.GetObject(db.LayoutDictionaryId,OpenMode.ForRead);var rows=new List<object>();
  foreach(DBDictionaryEntry de in ld){var l=(Layout)tr.GetObject(de.Value,OpenMode.ForRead);if(l.ModelType)continue;var btr=(BlockTableRecord)tr.GetObject(l.BlockTableRecordId,OpenMode.ForRead);
   var blocks=new List<object>();
   foreach(ObjectId id in btr){if(!(tr.GetObject(id,OpenMode.ForRead) is BlockReference br))continue;
    var root=(BlockTableRecord)tr.GetObject(br.IsDynamicBlock?br.DynamicBlockTableRecord:br.BlockTableRecord,OpenMode.ForRead);
    var props=new Dictionary<string,string>();if(br.IsDynamicBlock)foreach(DynamicBlockReferenceProperty p in br.DynamicBlockReferencePropertyCollection)if(!p.ReadOnly)props[p.PropertyName]=Convert.ToString(p.Value);
    var attrs=new List<object>();foreach(ObjectId aid in br.AttributeCollection){var a=(AttributeReference)tr.GetObject(aid,OpenMode.ForRead);attrs.Add(new{a.Tag,Text=a.TextString,Field=Code(a,tr)});}
    if(attrs.Count==0 && props.Count==0 && !root.IsFromExternalReference)continue;
    blocks.Add(new{Name=root.Name,Xref=root.IsFromExternalReference,Path=root.IsFromExternalReference?root.PathName:null,Handle=br.Handle.ToString(),Pos=new[]{br.Position.X,br.Position.Y},Scale=br.ScaleFactors.X,Props=props,Attrs=attrs});}
   rows.Add(new{Layout=l.LayoutName,Tab=l.TabOrder,Blocks=blocks});}
  return new{File=file,Layouts=rows};}
 [CommandMethod("POSTINSP")]public void Run(){
  var log=new List<string>();var outp=new Dictionary<string,object>();
  try{var mgr=new AcSmSheetSetMgrClass();var sdb=mgr.OpenDatabase(Path.Combine(Dir,"set.dst"),false);var ss=sdb.GetSheetSet();var rows=new List<object>();Walk((IAcSmSubset)ss,"",rows);
   outp["Set"]=new{Name=ss.GetName(),Id=ss.GetObjectId().GetHandle(),Sheets=rows};mgr.Close(sdb);}catch(System.Exception e){log.Add("SSM "+e);}
  foreach(var f in File.ReadAllLines(Path.Combine(Dir,"files.txt")).Where(x=>x.Trim()!=""))try{outp[f]=Dump(f);}catch(System.Exception e){log.Add(f+" "+e.Message);}
  outp["Log"]=log;File.WriteAllText(Path.Combine(Dir,"insp.json"),JsonSerializer.Serialize(outp,new JsonSerializerOptions{WriteIndented=true,Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping}));}
}
