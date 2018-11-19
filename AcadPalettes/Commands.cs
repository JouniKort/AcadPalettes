using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration; //Reference windowsformsintegration
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using System.Data.OleDb;

using Autodesk.AutoCAD.Runtime; //Reference acdbmgd, acmgd
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.IO;

[assembly: CommandClass(typeof(AcadPalettes.Commands))]
namespace AcadPalettes
{

    class Commands
    {
        private Dictionary<string, string> IdentifierPairs = new Dictionary<string, string>();
        private OleDbConnection DbConnection = null;

        static PaletteSet pLibrary = null;
        static PaletteSet pReplace = null;
        static PaletteSet pSpreadsheet = null;
        static PaletteSet pReplaceBlock = null;
        static PaletteSet pBlockAttributes = null;

        [CommandMethod("StartLibrary")]
        public void StartLibrary()
        {
            if(pLibrary == null)
            {
                pLibrary = new PaletteSet("Library");
                pLibrary.DockEnabled = (DockSides)((int)DockSides.Left + (int)DockSides.Right);

                Library Lib = new Library();
                ElementHost host = new ElementHost();
                host.AutoSize = true;
                host.Dock = DockStyle.Fill;
                host.Child = Lib;
                pLibrary.Add("Library", host);
            }

            pLibrary.KeepFocus = true;
            pLibrary.Visible = true;
        }

        /// <summary>
        /// Replace selected block values from multiple dwgs
        /// </summary>
        [CommandMethod("Replace")]
        public void StartReplace()
        {
            if (pReplace == null)
            {
                pReplace = new PaletteSet("Replace");
                pReplace.DockEnabled = (DockSides)((int)DockSides.Left + (int)DockSides.Right);

                Replace Rep = new Replace();
                ElementHost host = new ElementHost();
                host.AutoSize = true;
                host.Dock = DockStyle.Fill;
                host.Child = Rep;
                pReplace.Add("Replace", host);
            }
            pReplace.Visible = true;
        }

        /// <summary>
        /// get/set selected block values from multiple dwgs. Values are strored in a .csv file
        /// </summary>
        [CommandMethod("ReplaceSpreadsheet")]
        public void StartSpreadsheet()
        {
            if (pSpreadsheet == null)
            {
                pSpreadsheet = new PaletteSet("ReplaceSpreadsheet");
                pSpreadsheet.DockEnabled = (DockSides)((int)DockSides.Bottom);

                ReplaceSpreadsheet RepSpread = new ReplaceSpreadsheet();
                ElementHost host = new ElementHost();
                host.AutoSize = true;
                host.Dock = DockStyle.Fill;
                host.Child = RepSpread;
                pSpreadsheet.Add("ReplaceSpreadsheet", host);
            }
            pSpreadsheet.Visible = true;
        }

        /// <summary>
        /// Create an array with desired distance between numbers at a desired angle and count
        /// </summary>
        [CommandMethod("ArrayIncrement")]
        public void Array()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; //Valitaan avoinna oleva dokumentti
            Database acCurDb = acDoc.Database;

            using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
            {
                PromptSelectionResult selRes = acDoc.Editor.GetSelection();

                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet acSet = selRes.Value;

                    if (acSet[0] != null)
                    {
                        Entity acEnt = trans.GetObject(acSet[0].ObjectId, OpenMode.ForWrite) as Entity;
                        if (acEnt != null)
                        {
                            if (acEnt.GetType() != typeof(DBText)) { return; }
                            DBText text = acEnt as DBText;
                            int number = int.Parse(text.TextString);

                            PromptDistanceOptions distOpt = new PromptDistanceOptions("\nOffset");
                            PromptDoubleResult doubleRes = acDoc.Editor.GetDistance(distOpt);

                            if (doubleRes.Status == PromptStatus.OK)
                            {
                                double offset = doubleRes.Value;
                                PromptAngleOptions angOpt = new PromptAngleOptions("\nAngle");
                                angOpt.UseDashedLine = true;
                                angOpt.BasePoint = text.Position;
                                doubleRes = acDoc.Editor.GetAngle(angOpt);

                                if (doubleRes.Status == PromptStatus.OK)
                                {
                                    double angle = doubleRes.Value;
                                    PromptIntegerResult intRes = acDoc.Editor.GetInteger(new PromptIntegerOptions("\nCount"));
                                    if (intRes.Status == PromptStatus.OK)
                                    {
                                        int count = intRes.Value;

                                        BlockTable blkTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                                        BlockTableRecord blkRec;
                                        blkRec = trans.GetObject(blkTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                                        for (int i = 1; i <= count; i++)
                                        {
                                            DBText newText = new DBText();
                                            int increment = number + i;

                                            newText.Color = text.Color;
                                            newText.Height = text.Height;
                                            newText.Layer = text.Layer;
                                            newText.TextStyleId = text.TextStyleId;
                                            newText.Rotation = text.Rotation;

                                            newText.TextString = increment.ToString();
                                            newText.Position = new Point3d(text.Position.X + Math.Cos(angle) * offset * i, text.Position.Y + Math.Sin(angle) * offset * i, text.Position.Z);

                                            blkRec.AppendEntity(newText);
                                            trans.AddNewlyCreatedDBObject(newText, true);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

        /// <summary>
        /// The AutoCAD's find and replace doesn't work on some old dwgs, 
        /// this function implements simple find and replace for selection and the whole drawing
        /// </summary>
        [CommandMethod("FindAndReplace")]
        public void FindAndReplace()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; //Valitaan avoinna oleva dokumentti
            Database acCurDb = acDoc.Database;

            PromptKeywordOptions promptKey = new PromptKeywordOptions("\nScope:");
            promptKey.Keywords.Add("Drawing");
            promptKey.Keywords.Add("Selection");

            PromptResult promptRes = acDoc.Editor.GetKeywords(promptKey);

            PromptStringOptions promptString = new PromptStringOptions("\nReplace:");
            PromptResult promptStringResult = acDoc.Editor.GetString(promptString);

            string replace = promptStringResult.StringResult;

            promptString = new PromptStringOptions("\nWith:");
            promptStringResult = acDoc.Editor.GetString(promptString);

            string replaceWith = promptStringResult.StringResult;

            switch (promptRes.StringResult)
            {
                case "Drawing":
                    using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                        BlockTableRecordEnumerator recEnum = ms.GetEnumerator();

                        while (recEnum.MoveNext())
                        {
                            Entity ent = (Entity)trans.GetObject(recEnum.Current, OpenMode.ForRead);
                            if (ent.GetType() == typeof(BlockReference))
                            {
                                BlockReference blockRef = ent as BlockReference;
                                AttributeCollection attCol = blockRef.AttributeCollection;

                                foreach (ObjectId att in attCol)
                                {
                                    using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForWrite) as AttributeReference)
                                    {
                                        if (attRef.TextString.Contains(replace))
                                        {
                                            attRef.TextString = attRef.TextString.Replace(replace, replaceWith);
                                        }
                                    }
                                }
                            }
                        }
                        trans.Commit();
                    }
                    break;
                case "Selection":
                    PromptSelectionResult selRes = acDoc.Editor.GetSelection();
                    if(selRes.Status == PromptStatus.OK)
                    {
                        using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                        {
                            SelectionSet selSet = selRes.Value;
                            if (selSet[0] != null)
                            {
                                for (int i = 0; i < selSet.Count; i++)
                                {
                                    Entity acEnt = trans.GetObject(selSet[i].ObjectId, OpenMode.ForRead) as Entity;
                                    if (acEnt != null)
                                    {
                                        if (acEnt.GetType() != typeof(BlockReference)) { return; }

                                        BlockReference selectedBlock = acEnt as BlockReference;
                                        AttributeCollection attrCol = selectedBlock.AttributeCollection;

                                        BlockTable blkTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                                        BlockTableRecord blkRec;
                                        blkRec = trans.GetObject(blkTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;


                                        if (attrCol.Count == 0) { return; }
                                        foreach (ObjectId attRefId in attrCol)
                                        {
                                            using (AttributeReference attRef = trans.GetObject(attRefId, OpenMode.ForWrite) as AttributeReference)
                                            {
                                                if (attRef.TextString.Contains(replace))
                                                {
                                                    attRef.TextString = attRef.TextString.Replace(replace, replaceWith);
                                                }
                                            }
                                        }
                                        trans.Commit();
                                    }
                                }
                            }
                        }
                    }
                    break;

                default:
                    acDoc.Editor.WriteMessage("\nUnknown keyword\n");
                    break;
            }
        }

        /// <summary>
        /// Replace selected block with another block in multiple dwgs
        /// </summary>
        [CommandMethod("ReplaceBlock")]
        public void ReplaceBlock()
        {
            if (pLibrary == null)
            {
                pReplaceBlock = new PaletteSet("ReplaceBlock");
                pReplaceBlock.DockEnabled = (DockSides)((int)DockSides.Left + (int)DockSides.Right);

                ReplaceBlock Rep = new ReplaceBlock();
                ElementHost host = new ElementHost();
                host.AutoSize = true;
                host.Dock = DockStyle.Fill;
                host.Child = Rep;
                pReplaceBlock.Add("ReplaceBlock", host);
            }

            pReplaceBlock.KeepFocus = true;
            pReplaceBlock.Visible = true;
        }

        /// <summary>
        /// Adds event to update data in a database when block value is edited
        /// </summary>
        [CommandMethod("AddEvent")]
        public void AddEvent()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; //Valitaan avoinna oleva dokumentti
            Database acCurDb = acDoc.Database;

            using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
            {
                PromptSelectionResult selRes = acDoc.Editor.GetSelection();

                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet acSet = selRes.Value;

                    if (acSet[0] != null)
                    {
                        Entity acEnt = trans.GetObject(acSet[0].ObjectId, OpenMode.ForWrite) as Entity;
                        if (acEnt != null)
                        {
                            if (acEnt.GetType() != typeof(BlockReference)) { return; }

                            BlockReference selectedBlock = acEnt as BlockReference;
                            AttributeCollection attrCol = selectedBlock.AttributeCollection;

                            BlockTable blkTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                            BlockTableRecord blkRec;
                            blkRec = trans.GetObject(blkTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                            if (attrCol.Count == 0) { return; }
                            foreach (ObjectId attRefId in attrCol)
                            {
                                using (AttributeReference attRef = trans.GetObject(attRefId, OpenMode.ForWrite) as AttributeReference)
                                {
                                    attRef.Modified += AttRef_Modified;
                                }
                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

        private void AttRef_Modified(object sender, EventArgs e)
        {
            AttributeReference attRef = sender as AttributeReference;
            string tag = attRef.Tag;
            string sql = QueryBuilder(tag.Substring(2, tag.LastIndexOf(".") - 2), tag.Substring(tag.LastIndexOf(".") + 1), attRef.TextString);
            OleDbCommand cmd = new OleDbCommand(sql, DbConnection);
            cmd.ExecuteNonQuery();
        }

        private string QueryBuilder(string variable, string column, string value)
        {
            int AL = 2;
            string sql = "";

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; //Valitaan avoinna oleva dokumentti
            Database acCurDb = acDoc.Database;
            DatabaseSummaryInfo dbSumInfo = acCurDb.SummaryInfo;
            System.Collections.IDictionaryEnumerator info = dbSumInfo.CustomProperties;

            acDoc.Editor.WriteMessage("\nBuilding query for " + variable + "\n");

            while (info.MoveNext())
            {
                if (info.Key.Equals("IPIIRI")) { AL = 0; break; }
                else if (info.Key.Equals("SPIIRI")) { AL = 1; break; }
                else if (info.Key.Equals("DB") && DbConnection == null)
                {
                    string prop = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info = False; ", info.Value.ToString());
                    DbConnection = new OleDbConnection(prop);
                    DbConnection.Open();
                }
            }

            acDoc.Editor.WriteMessage(AL + "\n");

            info.Reset();
            while (info.MoveNext())
            {
                if (!info.Key.ToString().Contains("@"))
                {
                    IdentifierPairs.Add(
                        info.Key.ToString(),
                        info.Value.ToString());
                    acDoc.Editor.WriteMessage(info.Key.ToString() + " was added\n");
                }
                if (info.Key.ToString().Contains(variable) && !info.Key.ToString().Contains("@")) { sql = "ID = " + info.Value; }
                else if (info.Key.ToString().Contains(variable))
                {
                    switch (info.Value.ToString().Substring(0, 1))
                    {
                        case "L":
                            break;
                        case "E":
                            break;
                        case "R":
                            sql = "UPDATE IKYTKENTA SET " + column + "='" + value + "' WHERE RIVITUNN = '" + info.Value.ToString().Substring(2, info.Value.ToString().Length - 3) + "' ";
                            string val;
                            IdentifierPairs.TryGetValue(info.Key.ToString().Substring(info.Key.ToString().IndexOf("@") + 1), out val);
                            sql = sql + "AND LAITE_KIR = '" + val.Substring(2, val.ToString().Length - 3) + "' AND ";
                            break;
                        case "T":
                            break;
                        case "P":
                            break;
                        case "V":
                            break;
                        case "A":
                            break;
                        case "S":
                            break;
                        case "K":
                            break;
                        case "O":
                            break;
                        case "F":
                            break;
                        case "M":
                            break;
                        case "C":
                            break;
                    }
                    break;
                }
            }

            IdentifierPairs.Clear();

            if (AL == 0) { sql = sql + "POSITIO = (SELECT POSITIO FROM IPIIRI WHERE TMALLI = '" + Path.GetFileNameWithoutExtension(acDoc.Name) + "')"; }
            else if (AL == 1) { }

            acDoc.Editor.WriteMessage(sql + "\n");

            return sql;
        }

        //Registery
        #region
        [CommandMethod("Register")]
        public void Register()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            string sAppName = "AcadPalettes";

            Microsoft.Win32.RegistryKey regAcadProdKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sProdKey);
            Microsoft.Win32.RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Check to see if the "MyApp" key exists
            string[] subKeys = regAcadAppKey.GetSubKeyNames();
            foreach (string subKey in subKeys)
            {
                // If the application is already registered, exit
                if (subKey.Equals(sAppName))
                {
                    regAcadAppKey.Close();
                    return;
                }
            }

            // Get the location of this module
            string sAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // Register the application
            Microsoft.Win32.RegistryKey regAppAddInKey = regAcadAppKey.CreateSubKey(sAppName);
            regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
            regAppAddInKey.SetValue("LOADCTRLS", 14, RegistryValueKind.DWord);
            regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String);
            regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord);

            regAcadAppKey.Close();
        }

        [CommandMethod("Unregister")]
        public void UnregisterMyApp()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            string sAppName = "AcadPalettes";

            Microsoft.Win32.RegistryKey regAcadProdKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sProdKey);
            Microsoft.Win32.RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Delete the key for the application
            regAcadAppKey.DeleteSubKeyTree(sAppName);
            regAcadAppKey.Close();
        }
        #endregion
    }
}
