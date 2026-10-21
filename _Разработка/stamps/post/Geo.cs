using System;using System.IO;using System.Linq;using System.Collections.Generic;using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
public class PostGeo{[CommandMethod("POSTGEO")]public void Run(){var f=Environment.GetEnvironmentVariable("POSTEVAL_FILE");var o=new List<string>();
 using var db=new Database(false,true);db.ReadDwgFile(f,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
 using var tr=db.TransactionManager.StartTransaction();var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
 var root=(BlockTableRecord)tr.GetObject(bt["Рамка со штампом"],OpenMode.ForRead);
 void dump(BlockTableRecord b,string tag){foreach(ObjectId id in b){var e=(Entity)tr.GetObject(id,OpenMode.ForRead);
  if(e is Line l&&Math.Abs(l.StartPoint.Y-l.EndPoint.Y)<0.01&&l.StartPoint.Y>40&&l.StartPoint.Y<120)o.Add($"{tag} HLINE y={l.StartPoint.Y:0.#} x={l.StartPoint.X:0.#}..{l.EndPoint.X:0.#} vis={e.Visible}");
  if(e is MText m&&m.Text.Contains("D1-03"))o.Add($"{tag} MTEXT ({m.Location.X:0.#},{m.Location.Y:0.#}) h={m.TextHeight} vis={e.Visible} '{m.Text}' {m.Handle}");
  if(e is BlockReference r)o.Add($"{tag} BREF {((BlockTableRecord)tr.GetObject(r.BlockTableRecord,OpenMode.ForRead)).Name} ({r.Position.X:0.#},{r.Position.Y:0.#}) vis={e.Visible}");}}
 dump(root,"ROOT");
 foreach(ObjectId rid in root.GetAnonymousBlockIds()){var a=(BlockTableRecord)tr.GetObject(rid,OpenMode.ForRead);var refs=a.GetBlockReferenceIds(true,false);if(refs.Count==0)continue;
  var br=(BlockReference)tr.GetObject(refs[0],OpenMode.ForRead);string p=string.Join(",",br.DynamicBlockReferencePropertyCollection.Cast<DynamicBlockReferenceProperty>().Where(x=>!x.ReadOnly&&!x.PropertyName.StartsWith("Origin")).Select(x=>x.PropertyName+"="+x.Value));
  dump(a,"ANON["+p+"]");}
 File.WriteAllLines(f+".geo.txt",o);}}
