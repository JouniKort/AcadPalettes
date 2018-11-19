using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Autodesk.AutoCAD.Runtime; //Reference acdbmgd, acmgd
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.IO;

namespace AcadPalettes
{
    /// <summary>
    /// Interaction logic for ReplaceSpreadsheet.xaml
    /// </summary>
    public partial class ReplaceSpreadsheet : UserControl
    {

        private System.Data.DataTable DT = new System.Data.DataTable();
        private string pathExists = "";

        public ReplaceSpreadsheet()
        {
            InitializeComponent();
        }

        private void ButtonSelect_Click(object sender, RoutedEventArgs e)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            PromptSelectionResult selRes;
            using (EditorUserInteraction preventFlikkering = acDoc.Editor.StartUserInteraction(Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.Handle))
            {
                selRes = acDoc.Editor.GetSelection();
                preventFlikkering.End();
            }

            using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
            {
                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet acSet = selRes.Value;
                    if (acSet[0] != null)
                    {
                        if (trans.GetObject(acSet[0].ObjectId, OpenMode.ForRead) is Entity acEnt)
                        {
                            if (acEnt.GetType() == typeof(BlockReference))
                            {
                                BlockReference selectedBlock = acEnt as BlockReference;

                                //hBlock = selectedBlock.Handle;

                                AttributeCollection attCol = selectedBlock.AttributeCollection;
                                if (attCol.Count != 0)
                                {
                                    foreach (ObjectId att in attCol)
                                    {
                                        using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForRead) as AttributeReference)
                                        {
                                            acDoc.Editor.WriteMessage(selectedBlock.Handle.Value.ToString());
                                            System.Data.DataColumn col = new System.Data.DataColumn(selectedBlock.Name + "-" + selectedBlock.Handle.Value + "-" + attRef.Tag, typeof(string));
                                            col.DefaultValue = "";
                                            DT.Columns.Add(col);
                                        }
                                    }
                                    Spreadsheet.ItemsSource = null;
                                    Spreadsheet.ItemsSource = DT.DefaultView;
                                    }
                                else { acDoc.Editor.WriteMessage("Selected block did not contain any attributes.\n"); return; }
                            }
                            else { acDoc.Editor.WriteMessage("Selected entity was not a block.\n"); return; }
                        }
                        else { acDoc.Editor.WriteMessage("Selected entity was not an entity.\n"); return; }
                    }
                    else { acDoc.Editor.WriteMessage("Null selection.\n"); return; }
                }
                else { acDoc.Editor.WriteMessage("Selection cancelled.\n"); return; }
            }
        }

        private void ButtonRead_Click(object sender, RoutedEventArgs e)
        {
            foreach (DataRow DR in DT.Rows)
            {
                int index = 1;
                Database acDb = new Database(false, true);
                using (acDb)
                {
                    acDb.ReadDwgFile(DR["Filename"].ToString(), FileShare.Read, false, "");
                    using (Transaction trans = acDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(acDb.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        BlockTableRecordEnumerator recEnum = ms.GetEnumerator();

                        while (index < DT.Columns.Count - 1)
                        {
                            recEnum.Reset();
                            Handle bHandle = new Handle(long.Parse(DT.Columns[index].ColumnName.Split('-')[1].ToString()));
                            string blockName = DT.Columns[index].ColumnName.Split('-')[0].ToString();
                            while (recEnum.MoveNext())
                            {
                                Entity ent = (Entity)trans.GetObject(recEnum.Current, OpenMode.ForRead);
                                if (ent.GetType() == typeof(BlockReference))
                                {
                                    BlockReference blockRef = ent as BlockReference;
                                    if (blockRef.Handle.Value.CompareTo(bHandle.Value) == 0)
                                    {
                                        AttributeCollection attCol = blockRef.AttributeCollection;

                                        foreach (ObjectId att in attCol)
                                        {
                                            using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForRead) as AttributeReference)
                                            {
                                                if (DT.Columns.Contains(blockName + "-" + bHandle.Value + "-" + attRef.Tag))
                                                {
                                                    DR[blockName + "-" + bHandle.Value + "-" + attRef.Tag] = attRef.TextString;
                                                    index++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Spreadsheet.ItemsSource = null;
            Spreadsheet.ItemsSource = DT.DefaultView;
        }

        private void ButtonWrite_Click(object sender, RoutedEventArgs e)
        {
            if (Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Count > 0)
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.CurrentDocument.Editor.WriteMessage("Close the open documents:\n");
                foreach (Document doc in Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager)
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.CurrentDocument.Editor.WriteMessage(doc.Name + "\n");
                }
                return;
            }

            foreach (DataRow DR in DT.Rows)
            {
                int index = 1;
                Database acDb = new Database(false, true);
                using (acDb)
                {
                    if (File.Exists(DR["Filename"].ToString()))
                    {
                        acDb.ReadDwgFile(DR["Filename"].ToString(), FileShare.ReadWrite, false, "");
                        pathExists = DR["Filename"].ToString();
                    }
                    else
                        acDb.ReadDwgFile(pathExists, FileShare.ReadWrite, false, "");

                    using (Transaction trans = acDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(acDb.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        BlockTableRecordEnumerator recEnum = ms.GetEnumerator();

                        while(index < DT.Columns.Count - 1)
                        {
                            recEnum.Reset();
                            Handle bHandle = new Handle(long.Parse(DT.Columns[index].ColumnName.Split('-')[1].ToString()));
                            string blockName = DT.Columns[index].ColumnName.Split('-')[0].ToString();

                            while (recEnum.MoveNext())
                            {
                                Entity ent = (Entity)trans.GetObject(recEnum.Current, OpenMode.ForRead);
                                if (ent.GetType() == typeof(BlockReference))
                                {
                                    BlockReference blockRef = ent as BlockReference;
                                    if (blockRef.Handle.Value.CompareTo(bHandle.Value) == 0)
                                    {
                                        AttributeCollection attCol = blockRef.AttributeCollection;

                                        foreach (ObjectId att in attCol)
                                        {
                                            using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForWrite) as AttributeReference)
                                            {
                                                if (DT.Columns.Contains(blockName + "-" + bHandle.Value + "-" + attRef.Tag))
                                                {
                                                    attRef.TextString = DR[blockName + "-" + bHandle.Value + "-" + attRef.Tag].ToString();
                                                    index++;
                                                }
                                            }
                                        }

                                        break;
                                    }
                                }
                            }
                        }
                        trans.Commit();
                        acDb.SaveAs(DR["Filename"].ToString(), DwgVersion.Current);
                    }
                }
            }
        }

        private void ButtonBrowse_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog diag = new System.Windows.Forms.OpenFileDialog();
            if(diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DT = new System.Data.DataTable();
                string path = diag.FileName;
                string line;
                using (StreamReader sr = new StreamReader(path))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (DT.Columns.Count == 0)
                        {
                            foreach (string val in line.Split(';'))
                            {
                                DT.Columns.Add(val);
                            }
                        }
                        else
                        {
                            string[] vals = line.Split(';');
                            DataRow DR = DT.NewRow();
                            for (int i = 0; i < DT.Columns.Count; i++)
                            {
                                DR[i] = vals[i];                           
                            }
                            DT.Rows.Add(DR);
                        }
                    }
                }
                Spreadsheet.ItemsSource = DT.DefaultView;
            }
        }

        private void ButtonExport_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog diag = new System.Windows.Forms.SaveFileDialog();
            if(diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (System.IO.Path.GetExtension(diag.FileName).Equals(".csv"))
                {
                    using (StreamWriter sw = new StreamWriter(diag.FileName, false))
                    {
                        string line = "";
                        foreach(System.Data.DataColumn DC in DT.Columns)
                        {
                            line += DC.ColumnName + ";";
                        }
                        line = line.Substring(0, line.Length - 1);
                        sw.WriteLine(line);
                        foreach(DataRow DR in DT.Rows)
                        {
                            line = "";
                            for(int i = 0; i < DT.Columns.Count; i++)
                            {
                                line += DR[i].ToString() + ";";
                            }
                            line = line.Substring(0, line.Length - 1);
                            sw.WriteLine(line);
                        }
                    }
                }
            }
        }
    }
}
