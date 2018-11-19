using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;
using System.Collections.ObjectModel;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace AcadPalettes
{
    /// <summary>
    /// Interaction logic for Fetch.xaml
    /// </summary>
    public partial class Replace : UserControl
    {
        private ObservableCollection<DWG> DWGs = new ObservableCollection<DWG>();
        private Handle hBlock;

        public Replace()
        {
            InitializeComponent();
            LB_Replace.ItemsSource = DWGs;
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

                                LabelBlock.Content = selectedBlock.Name;
                                hBlock = selectedBlock.Handle;

                                AttributeCollection attCol = selectedBlock.AttributeCollection;
                                if (attCol.Count != 0)
                                {
                                    List<string> attList = new List<string>();
                                    foreach (ObjectId att in attCol)
                                    {
                                        using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForRead) as AttributeReference)
                                        {
                                            attList.Add(attRef.Tag);
                                        }
                                    }
                                    ComboAttribute.ItemsSource = attList;
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

        private void ButtonBrowse_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog diag = new System.Windows.Forms.OpenFileDialog();
            diag.Multiselect = true;
            diag.DefaultExt = "dwg files (*.dwg)|*.dwg";

            if(diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DWGs.Clear();
                foreach (string path in diag.FileNames)
                {
                    DWGs.Add(new DWG() { Path = path, Check = false });
                }
                //LB_Replace.ItemsSource = DWGs;
            }
        }

        private void RuttonReplace_Click(object sender, RoutedEventArgs e)
        {
            if(Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Count > 0)
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.CurrentDocument.Editor.WriteMessage("Close the open documents:\n");
                foreach(Document doc in Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager)
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.CurrentDocument.Editor.WriteMessage(doc.Name + "\n");
                }
                return;
            }

            foreach (DWG item in DWGs)
            {
                Database acDb = new Database(false, true);
                using (acDb)
                {
                    acDb.ReadDwgFile(item.Path, FileShare.ReadWrite, false, "");
                    using (Transaction trans = acDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(acDb.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        BlockTableRecordEnumerator recEnum = ms.GetEnumerator();

                        while (recEnum.MoveNext())
                        {
                            Entity ent = (Entity)trans.GetObject(recEnum.Current, OpenMode.ForRead);
                            if (ent.GetType() == typeof(BlockReference))
                            {
                                BlockReference blockRef = ent as BlockReference;
                                if (blockRef.Handle == hBlock)
                                {
                                    AttributeCollection attCol = blockRef.AttributeCollection;

                                    foreach (ObjectId att in attCol)
                                    {
                                        using (AttributeReference attRef = trans.GetObject(att, OpenMode.ForWrite) as AttributeReference)
                                        {
                                            if (attRef.Tag == ComboAttribute.SelectedValue.ToString())
                                            {
                                                attRef.TextString = TBReplacement.Text;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        trans.Commit();
                        acDb.SaveAs(item.Path, DwgVersion.Current);
                        item.Check = true;
                    }
                }
            }
        }

        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Label lb = sender as Label;
            string path = lb.Content.ToString();
            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(path,false);
        }
    }
}
