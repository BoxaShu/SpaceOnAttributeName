using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Reflection;

using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;

[assembly: Rtm.CommandClass(typeof(SpaceOnAttributeName.Commands))]

namespace SpaceOnAttributeName
{
    public class Commands : Rtm.IExtensionApplication
    {

        /// <summary>
        /// Загрузка библиотеки
        /// http://through-the-interface.typepad.com/through_the_interface/2007/03/getting_the_lis.html
        /// </summary>
        #region 
        public void Initialize()
        {
            String assemblyFileFullName = GetType().Assembly.Location;
            String assemblyName = System.IO.Path.GetFileName(
                                                      GetType().Assembly.Location);
            Assembly assembly = GetType().Assembly;


            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            //Db.Database acCurDb = acDoc.Database;
            Ed.Editor ed = acDoc.Editor;


            // Сообщаю о том, что произведена загрузка сборки 
            //и указываю полное имя файла,
            // дабы было видно, откуда он загружен
            ed.WriteMessage(string.Format("\n{0} {1} {2}.\n{3}: {4}\n{5}\n",
                      "Assembly", assemblyName, "Loaded",
                      "Assembly File:", assemblyFileFullName,
                       "Copyright © Владимир Шульжицкий, 2017"));


            //Вывожу список комманд определенных в библиотеке
            ed.WriteMessage("\nStart list of commands: \n");

            // Just get the commands for this assembly
            App.DocumentCollection dm = App.Application.DocumentManager;
            Assembly asm = Assembly.GetExecutingAssembly();

            string[] cmds = GetCommands(asm, false);
            foreach (string cmd in cmds)
                ed.WriteMessage(cmd + "\n");

            ed.WriteMessage("\nEnd list of commands.\n");


        }

        public void Terminate()
        {
            Console.WriteLine("finish!");
        }

        /// <summary>
        /// Получение списка комманд определенных в сборке
        /// </summary>
        /// <param name="asm"></param>
        /// <param name="markedOnly"></param>
        /// <returns></returns>
        private static string[] GetCommands(Assembly asm, bool markedOnly)
        {
            StringCollection sc = new StringCollection();
            object[] objs =
              asm.GetCustomAttributes(
                typeof(Rtm.CommandClassAttribute),
                true
              );
            Type[] tps;
            int numTypes = objs.Length;
            if (numTypes > 0)
            {
                tps = new Type[numTypes];
                for (int i = 0; i < numTypes; i++)
                {
                    Rtm.CommandClassAttribute cca =
                      objs[i] as Rtm.CommandClassAttribute;
                    if (cca != null)
                    {
                        tps[i] = cca.Type;
                    }
                }
            }
            else
            {
                // If we're only looking for specifically
                // marked CommandClasses, then use an
                // empty list
                if (markedOnly)
                    tps = new Type[0];
                else
                    tps = asm.GetExportedTypes();
            }
            foreach (Type tp in tps)
            {
                MethodInfo[] meths = tp.GetMethods();
                foreach (MethodInfo meth in meths)
                {
                    objs =
                      meth.GetCustomAttributes(
                        typeof(Rtm.CommandMethodAttribute),
                        true
                      );
                    foreach (object obj in objs)
                    {
                        Rtm.CommandMethodAttribute attb =
                          (Rtm.CommandMethodAttribute)obj;
                        sc.Add(attb.GlobalName);
                    }
                }
            }
            string[] ret = new string[sc.Count];
            sc.CopyTo(ret, 0);
            return ret;
        }
        #endregion




        [Rtm.CommandMethod("SOAN")]
        static public void SOAN()
        {
            SpaceOnAttributeName();
        }
        

        [Rtm.CommandMethod("SpaceOnAttributeName")]
        static public void SpaceOnAttributeName()
        {
            // Получение текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;

            // старт транзакции
            using (Db.Transaction acTrans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                Db.TypedValue[] acTypValAr = new Db.TypedValue[1];
                acTypValAr.SetValue(new Db.TypedValue((int)Db.DxfCode.Start, "INSERT"), 0);
                Ed.SelectionFilter acSelFtr = new Ed.SelectionFilter(acTypValAr);

                Ed.PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection(acSelFtr);
                if (acSSPrompt.Status == Ed.PromptStatus.OK)
                {
                    Ed.SelectionSet acSSet = acSSPrompt.Value;
                    foreach (Ed.SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            if (acSSObj.ObjectId.ObjectClass.IsDerivedFrom(Rtm.RXClass.GetClass(typeof(Db.BlockReference))))
                            {

                                Db.BlockReference acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                    Db.OpenMode.ForRead) as Db.BlockReference;

                                Db.BlockTableRecord blr = acTrans.GetObject(acEnt.BlockTableRecord,
                                                               Db.OpenMode.ForRead) as Db.BlockTableRecord;
                                if (acEnt.IsDynamicBlock)
                                {
                                    blr = acTrans.GetObject(acEnt.DynamicBlockTableRecord,
                                                               Db.OpenMode.ForRead) as Db.BlockTableRecord;
                                }

                                if (blr.HasAttributeDefinitions)
                                {
                                    foreach (Db.ObjectId id in blr)
                                    {
                                        if (id.ObjectClass.IsDerivedFrom(Rtm.RXClass.GetClass(typeof(Db.AttributeDefinition))))
                                        {
                                            Db.AttributeDefinition acAttrRef = acTrans.GetObject(id,
                                                            Db.OpenMode.ForWrite) as Db.AttributeDefinition;

                                            if (acAttrRef != null)
                                            {
                                                acAttrRef.Tag = acAttrRef.Tag.Replace('_', ' ');
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                acTrans.Commit();
            }
        }
    }
}
