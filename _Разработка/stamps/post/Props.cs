using System;using System.IO;using System.Collections.Generic;using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.Interop;
public class PostProps {
 static void P(IAcSmComponent c,string pre,List<string> o){try{var en=c.GetCustomPropertyBag().GetPropertyEnumerator();string n;AcSmCustomPropertyValue v;
  while(true){en.Next(out n,out v);if(n==null)break;o.Add(pre+"\t"+n+"\t"+v.GetValue()+"\t"+v.GetFlags());}}catch(System.Exception e){o.Add(pre+"\tERR "+e.Message);}}
 static void W(IAcSmSubset s,string path,List<string> o){var en=s.GetSheetEnumerator();IAcSmComponent c;while((c=en.Next())!=null){
  if(c.GetTypeName()=="AcSmSheet"){var sh=(IAcSmSheet)c;P(c,"SHEET "+sh.GetNumber()+" "+sh.GetTitle(),o);}else if(c.GetTypeName()=="AcSmSubset"){P(c,"SUB "+c.GetName(),o);W((IAcSmSubset)c,path+"/"+c.GetName(),o);}}}
 [CommandMethod("POSTPROPS")]public void Run(){var o=new List<string>();var dir=Environment.GetEnvironmentVariable("POSTINSP_DIR");
  try{var m=new AcSmSheetSetMgrClass();var db=m.OpenDatabase(Path.Combine(dir,"set.dst"),false);var ss=db.GetSheetSet();P(ss,"SET",o);
   W((IAcSmSubset)ss,"",o);m.Close(db);}catch(System.Exception e){o.Add("ERR "+e);}
  File.WriteAllLines(Path.Combine(dir,"props.txt"),o);}
}
