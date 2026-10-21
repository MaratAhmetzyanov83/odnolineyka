using System;using System.IO;using System.Collections.Generic;using Autodesk.AutoCAD.Runtime;using Autodesk.AutoCAD.Interop;
// Пишет свойства сдвига сквозной нумерации в подшивку POSTSET_FILE: строки «имя=значение» из POSTSET_PROPS.
public class PostSetProps {
 [CommandMethod("POSTSETPROPS")]public void Run(){var f=Environment.GetEnvironmentVariable("POSTSET_FILE");var o=new List<string>();AcSmDatabase db=null;bool locked=false;
  try{var m=new AcSmSheetSetMgrClass();db=m.OpenDatabase(f,false);db.LockDb(db);locked=true;var ss=db.GetSheetSet();var bag=ss.GetCustomPropertyBag();
   foreach(var line in Environment.GetEnvironmentVariable("POSTSET_PROPS").Split('|')){var kv=line.Split('=');
    var v=new AcSmCustomPropertyValueClass();v.InitNew(bag);v.SetFlags(PropertyFlags.CUSTOM_SHEETSET_PROP);v.SetValue(kv[1]);bag.SetProperty(kv[0],v);o.Add("SET "+kv[0]+"="+kv[1]);}
   db.UnlockDb(db,true);locked=false;m.Close(db);o.Add("OK");}
  catch(System.Exception e){if(locked)try{db.UnlockDb(db,false);}catch{}o.Add("ERR "+e);}
  File.WriteAllLines(f+".setprops.txt",o);}
}
