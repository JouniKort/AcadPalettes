using System;
using System.Collections.Generic;
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

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Collections.ObjectModel;
using Autodesk.AutoCAD.Runtime;

namespace AcadPalettes
{
    /// <summary>
    /// Interaction logic for ReplaceBlock.xaml
    /// </summary>
    public partial class ReplaceBlock : UserControl
    {
        private ObservableCollection<DWG> DWGs = new ObservableCollection<DWG>();
        private string path;
        private Handle hBlockReplace;
        private string BlockReplacement;

        public ReplaceBlock()
        {
            InitializeComponent();
            LB_Replace.ItemsSource = DWGs;
        }

        private void ButtonSelect_Click(object sender, RoutedEventArgs e)
        {
            Button b = sender as Button;
            if (b.Name.Equals("ButtonSelect1")) { SelectBlock(1); }
            else { SelectBlock(2); }
        }

        private void SelectBlock(int button)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            path = acDoc.Name;

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

                                if (button == 1)
                                {
                                    hBlockReplace = selectedBlock.Handle;
                                    LabelBlock1.Content = selectedBlock.Name;
                                }
                                else
                                {
                                    BlockReplacement = selectedBlock.Name;
                                    LabelBlock2.Content = selectedBlock.Name;
                                }
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
            diag.Filter = "dwg files (*.dwg)|*.dwg";

            if (diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
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
            AddToDB();
            ReplaceBlocks();
        }

        private void AddToDB()
        {
            using (Database DB = new Database(false, true))
            {
                DB.ReadDwgFile(path, FileShare.Read, false, "");
                ObjectIdCollection ids = new ObjectIdCollection();

                using (Transaction tr = DB.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(DB.BlockTableId, OpenMode.ForRead);

                    if (bt.Has(BlockReplacement))
                    {
                        ids.Add(bt[BlockReplacement]);
                    }
                    tr.Commit();
                }

                if (ids.Count != 0)
                {
                    foreach(DWG dwg in DWGs)
                    {
                        Database destdb = new Database(false, true);
                        using (destdb)
                        {
                            destdb.ReadDwgFile(dwg.Path, FileShare.ReadWrite, false, "");

                            IdMapping iMap = new IdMapping();
                            destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Ignore, false);
                            destdb.SaveAs(dwg.Path, DwgVersion.Current);
                        }
                    }
                }
            }
        }

        private void ReplaceBlocks()
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
                            Entity ent = (Entity)trans.GetObject(recEnum.Current, OpenMode.ForWrite);
                            if (ent.GetType() == typeof(BlockReference))
                            {
                                BlockReference blockRef = ent as BlockReference;
                                if (blockRef.Handle.Value.CompareTo(hBlockReplace.Value) == 0)
                                {
                                    using (BlockReference acBlkRef = new BlockReference(blockRef.Position, bt[BlockReplacement]))
                                    {
                                        BlockTableRecord acCurSpaceBlkTblRec;
                                        acCurSpaceBlkTblRec = trans.GetObject(acDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                                        trans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        AddAttributes(acBlkRef);
                                    }
                                    item.Check = true;
                                    blockRef.Erase();
                                    break;
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

        private void AddAttributes(BlockReference bref)
        {
            Transaction trans = bref.Database.TransactionManager.TopTransaction;

            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bref.BlockTableRecord, OpenMode.ForRead);
            RXClass attDefClass = RXClass.GetClass(typeof(AttributeDefinition));

            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass != attDefClass)
                    continue;
                AttributeDefinition attDef = (AttributeDefinition)trans.GetObject(id, OpenMode.ForRead);
                AttributeReference attRef = new AttributeReference();
                attRef.SetAttributeFromBlock(attDef, bref.BlockTransform);
                bref.AttributeCollection.AppendAttribute(attRef);
                trans.AddNewlyCreatedDBObject(attRef, true);
            }
        }
    }
}
