using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Data.SqlClient;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Runtime.Serialization.Formatters.Binary;

namespace AcadPalettes
{
    /// <summary>
    /// Interaction logic for Library.xaml
    /// </summary>
    public partial class Library : UserControl
    {

        public Library()
        {
            InitializeComponent();
        }

        private void ButtonCopy_Click(object sender, RoutedEventArgs e)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; //Valitaan avoinna oleva dokumentti
            Database acCurDb = acDoc.Database;

            using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
            {
                acDoc.Window.Focus();
                PromptSelectionResult selRes = acDoc.Editor.GetSelection();

                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet acSet = selRes.Value;

                    if (acSet[0] != null)
                    {
                        acDoc.Editor.WriteMessage("\nSelection found\n");

                        for (int i = 0; i < acSet.Count; i++)
                        {
                            acDoc.Editor.WriteMessage("Processing entity " + acSet[i].ObjectId.ToString() + "\n");

                            if (trans.GetObject(acSet[i].ObjectId, OpenMode.ForRead) is Entity acEnt)
                            {
                                if (acEnt.GetType() != typeof(BlockReference)) { acDoc.Editor.WriteMessage("\nYour selection contained a non-block"); return; }

                                BlockReference selectedBlock = acEnt as BlockReference;

                                MemoryStream memStream = new MemoryStream();
                                StreamWriter sw = new StreamWriter(memStream);
                                sw.Write(selectedBlock);
                                sw.Flush();

                                SqlConnection sqlConnection = new SqlConnection("Data Source=localhost\\SQLEXPRESS;Initial Catalog=AutoCAD;Integrated Security=SSPI");
                                sqlConnection.Open();

                                //SqlCommand sqlCmd = new SqlCommand("INSERT INTO Library(Memory) VALUES (@VarBinary)", sqlConnection);
                                //sqlCmd.Parameters.Add("@VarBinary", System.Data.SqlDbType.VarBinary, Int32.MaxValue);
                                //sqlCmd.Parameters["@VarBinary"].Value = memStream.GetBuffer();

                                //sqlCmd.ExecuteNonQuery();

                                SqlCommand sqlCmd = new SqlCommand("SELECT Memory FROM Library WHERE ID=4");
                                using (var reader = sqlCmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        BinaryFormatter bf = new BinaryFormatter();
                                        BlockReference br = bf.Deserialize(reader.GetStream(0)) as BlockReference;
                                    }
                                }

                                    sqlConnection.Close();
                            }
                        }
                    }
                }
            }
        }
    }
}
