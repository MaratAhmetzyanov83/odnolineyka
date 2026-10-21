using System;using System.IO;using System.Linq;using System.Collections.Generic;using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.DatabaseServices;
public class PostTags{[CommandMethod("POSTTAGS")]public void Run(){var f=Environment.GetEnvironmentVariable("POSTEVAL_FILE");var o=new List<string>();
 using var db=new Database(false,true);db.ReadDwgFile(f,FileOpenMode.OpenForReadAndAllShare,false,null);db.CloseInput(true);
 using var tr=db.TransactionManager.StartTransaction();var bt=(BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
 foreach(ObjectId id in (BlockTableRecord)tr.GetObject(bt["Рамка со штампом"],OpenMode.ForRead))if(tr.GetObject(id,OpenMode.ForRead) is AttributeDefinition a)o.Add($"{a.Tag}\t({a.Position.X:0.#},{a.Position.Y:0.#})\talign({a.AlignmentPoint.X:0.#},{a.AlignmentPoint.Y:0.#})\tinv={a.Invisible}\t'{a.TextString}'");
 File.WriteAllLines(f+".tags.txt",o);}}
