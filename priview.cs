using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.IO;

namespace DocumentsModule
{
    public partial class priview : Form
    {
        public priview()
        {
            InitializeComponent();
        }

        private void priview_Load(object sender, EventArgs e)
        {
            SqlConnection conn1 = new SqlConnection(Global.ConStr);
            SqlCommand comm1 = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn1);
            comm1.Parameters.AddWithValue("@id", 533329);
            conn1.Open();
            SqlDataReader sdr1 = comm1.ExecuteReader();
            sdr1.Read();
            string mainFile_extention = sdr1.GetString(1);

            //isOriginalNull = sdr1.IsDBNull(0);
            byte[] fileData = sdr1.GetSqlBytes(0).Buffer;
            string filePath = Program.folderPath + "\\" + "533329" + "." + mainFile_extention;
            File.WriteAllBytes(filePath, fileData);
            if (!File.Exists(filePath))
            {
                File.WriteAllBytes(filePath, fileData);
            }
            string htmlPath = Path.ChangeExtension(filePath, ".html");
            Word.Application wapp = new Word.Application();
            Word.Document docu = wapp.Documents.Open(filePath);
            docu.SaveAs2(htmlPath, Word.WdSaveFormat.wdFormatHTML);
            docu.Close();
            webBrowser1.Navigate(htmlPath);
            webBrowser1.Visible = true;
            webBrowser1.BringToFront();
            conn1.Close();
        }
    }
}
