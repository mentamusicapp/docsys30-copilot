using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Threading;

namespace DocumentsModule
{
    public partial class SaveTxt : Form
    {
        string conStr = Global.ConStr;
        static List<int> documents = new List<int>();
        bool okID = false, //האם לסנן לפי שוטף
            okDate = false; //האם לסנן לפי תאריך
        private List<int> searchResults = new List<int>();
        string dirPath = "C:\\Temp";

        public SaveTxt()
        {
            InitializeComponent();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                okDate = false;
                dateTimePicker1.ResetText();
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                dateTimePicker2.Value = DateTime.Today.AddMonths(-1);
            }
            else
            {
                okDate = true;
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int res1, res2;
                bool ok1 = int.TryParse(textBox1.Text, out res1), ok2 = int.TryParse(textBox2.Text, out res2);
                
                if (ok1)
                {
                    textBox2.Text = textBox1.Text;
                }
            }

            catch { }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value < dateTimePicker2.Value)
            {
                MessageBox.Show("לא ניתן לבחור תאריך סיום הקודם לתאריך ההתחלה", "בחירת תאריך שגויה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "") okID = false; else okID = true;

            DialogResult res = System.Windows.Forms.DialogResult.Yes;
            if (res == DialogResult.Yes)
            {
                Cursor = Cursors.WaitCursor;
                int numOfDocuments = getDocuments("shotef_mismach", false);
                if (numOfDocuments == 0)
                    MessageBox.Show("לא נמצאו מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                      MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                else if (numOfDocuments != -1)
                {
                    MessageBox.Show("טקסט של המסמכים עודכן.", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    progressBar1.Value = 1;
                    progressBar1.Visible = false;
                }
                Cursor = Cursors.Default;
            }
        }
        private int getDocuments(string sortColumn, bool order)
        {
            SqlConnection conn = new SqlConnection(conStr);
            SqlDataReader sdr;

            SqlCommand comm = new SqlCommand("select shotef_mismach Shotef, file_extension Ext,datalength(file_data) FileDataLen"
            + " from Documents"
            + " where (shotef_mismach between @idMin and @idMax or @idMin+@idMax = 0)"
            + " and datalength(file_data)>8000"
            + " and (tarich_hamichtav between @dateMin and @dateMax or(@dateMin = '00000000' and @dateMax = '00000000'))", conn);
            //+ " order by Shotef, Nispah"

            int res1 = 0, res2 = 0;
            if (int.TryParse(textBox1.Text, out res1))
                comm.Parameters.AddWithValue("@idMin", res1);
            else if (textBox1.Text == "")
                comm.Parameters.AddWithValue("@idMin", res1);
            else
                res1 = -1;

            if (int.TryParse(textBox2.Text, out res2))
                comm.Parameters.AddWithValue("@idMax", res2);
            else if (textBox2.Text == "")
                comm.Parameters.AddWithValue("@idMax", res2);
            else
                res2 = -1;

            if (res1 != -1 && res2 != -1)
            {
                if (res1 > res2)
                {
                    MessageBox.Show("לא ניתן להזין שוטף סיום קטן יותר משוטף התחלה", "חיפוש שוטפים שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
            }
            else
            {
                MessageBox.Show("שוטף יכול להכיל רק מספרים", "חיפוש שוטפים שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return -1;
            }

            if (okDate)
            {
                DateTime dateMin = dateTimePicker2.Value, dateMax = dateTimePicker1.Value;
                string dMin = dateMin.Year + "" + dateMin.Month.ToString().PadLeft(2, '0') + dateMin.Day.ToString().PadLeft(2, '0');
                string dMax = dateMax.Year + "" + dateMax.Month.ToString().PadLeft(2, '0') + dateMax.Day.ToString().PadLeft(2, '0');
                comm.Parameters.AddWithValue("@dateMin", dMin);
                comm.Parameters.AddWithValue("@dateMax", dMax);
            }
            else
            {
                comm.Parameters.AddWithValue("@dateMin", "00000000");
                comm.Parameters.AddWithValue("@dateMax", "00000000");
            }

            comm.CommandTimeout = 0;
            comm.Connection = conn;
            conn.Open();
            DataTable dataTable = new DataTable();
            while (true)
            {
                try
                {
                    sdr = comm.ExecuteReader();
                    
                    dataTable.Load(sdr);
                    conn.Close();
                    dataGridView1.DataSource = dataTable;
                   // dataGridView1.Columns["FileDataLen"].Visible = false;
                    break;
                }
                catch (Exception e)
                { //MessageBox.Show(e.Message);

                    Log(dirPath + "\\DocExp\\", 0, "TextUpdate_Exceptions", e.Message);
                }
            }

            try
            {
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                string DocPath = dirPath + "\\DocUpd";
                if (!Directory.Exists(DocPath))
                    Directory.CreateDirectory(DocPath);
                
               // int numOfThreads = 10;
               // Thread[] threads = new Thread[numOfThreads];
                SaveTxtOfTheDocuments(0, dataTable, DocPath, comm, conn, conStr);
               // for (int i = 0; i<numOfThreads; i++)
                //{
                   // threads[i] = new Thread(() => SaveTxtOfTheDocuments( i, dataTable, DocPath, comm, conn, conStr));
                   // threads[i].Start();
                    
                //}
                /*for (int i = 0; i < numOfThreads; i++)
                {
                    threads[i].Join();
                }*/
            }
            catch (Exception e)
            {

            }
            Cursor = Cursors.Default;
            return documents.Count;
        }
        static void SaveTxtOfTheDocuments(int lastNum, DataTable dt,string DocPath, SqlCommand comm,SqlConnection conn, string conStr)
        {
            //   if (lastNum == 10)
            //       lastNum = 0;
            //   var filteredRows = dt.AsEnumerable().Where(row => row.Field<int>("Shotef") % 10 == lastNum);
            //   dt = filteredRows.Any() ? filteredRows.CopyToDataTable() : dt.Clone();
            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                int Shotef = (int)row["Shotef"];
                string Ext = row["Ext"].ToString().Replace(".", "");

                documents.Add(Shotef);
                if ((Ext.ToLower() == "doc" || Ext.ToLower() == "docx"))
                {
                    long FileLength;
                   // byte[] fileData;
                    try
                    {
                        FileLength = (long)row["FileDataLen"];
                      //  fileData = (byte[])row["FileData"];
                    }
                    catch
                    {
                        FileLength = 0;
                        //fileData = null;
                    }

                    string DocFile = DocPath + "\\" + Shotef+"."+Ext;
                    if (FileLength > 0)
                    {
                        try
                        {
                            // Emily Lutvak - 30.10.2024
                            // loop all the documents by shotef number and select each file-data and write to docs and read from the document
                            //------------------------------
                            SqlConnection conn2 = new SqlConnection(conStr);
                            SqlCommand comm2 = new SqlCommand("select file_data from Documents where (shotef_mismach = @Shotef)", conn2);
                            comm2.Parameters.AddWithValue("@Shotef", Shotef);

                            comm2.CommandTimeout = 0;
                            comm2.Connection = conn2;
                            conn2.Open();
                            DataTable dataTable = new DataTable();

                            while (true)
                            {
                                try
                                {
                                    SqlDataReader sdr = comm2.ExecuteReader();

                                    dataTable.Load(sdr);
                                    conn2.Close();
                                    break;
                                }
                                catch (Exception ex)
                                { //MessageBox.Show(e.Message); 
                                    Log(DocPath + "\\" + lastNum, Shotef, "TextUpdate_Exceptions", ex.Message);
                                }
                            }


                            if (dataTable.Rows.Count > 0)
                            {
                                byte[] fileData = (byte[])dataTable.Rows[0][0];

                                File.WriteAllBytes(DocFile, fileData);
                                Word.Document doc = null;// wapp.Documents.Open(DocFile);
                                string text = PublicFuncsNvars.docToTxt(doc, DocFile);
                                //PublicFuncsNvars.saveDocToDB(ref fileData, Shotef, DocFile, ref comm, ref conn, text);
                                comm = new SqlCommand("UPDATE dbo.documents SET Txt=@TextString, LastTxtUpdateDate=@dateTime WHERE shotef_mismach=@id", conn);

                                comm.Parameters.AddWithValue("@id", Shotef);
                                comm.Parameters.AddWithValue("@TextString", text);
                                comm.Parameters.AddWithValue("@dateTime", DateTime.Now);

                                //conn.Close();
                                while (true)
                                {
                                    try
                                    {
                                        conn.Open();
                                        try
                                        {
                                            comm.ExecuteNonQuery();
                                        }
                                        catch (Exception e)
                                        {
                                            Log(DocPath + "\\" + lastNum, Shotef, "TextUpdate_Exceptions", e.Message);
                                        }
                                        conn.Close();
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(DocPath + "\\" + lastNum, Shotef, "TextUpdate_Exceptions", ex.Message);
                                    }
                                }
                                /*try
                                {
                                    File.Delete(DocFile);
                                }
                                catch (Exception ex)
                                {
                                    int i = 0;
                                    while (File.Exists(DocFile) && (++i) <= 5)
                                    {
                                        //Thread.Sleep(1000);
                                        File.Delete(DocFile);
                                    }
                                }*/
                                Log(DocPath + "\\" + lastNum, Shotef, "TextUpdate", "The text has been successfuly updated", (i++).ToString());
                            }



                            //------------------------------
                        }
                        catch (Exception ex)
                        {
                               Log(DocPath + "\\" + lastNum, Shotef, "TextUpdate_Exceptions", ex.Message);
                            /*   Thread.Sleep(2000);
                               try
                               {
                                   File.WriteAllBytes(DocFile, fileData);
                                   Word.Document doc = null;// wapp.Documents.Open(DocFile);
                                   string text = PublicFuncsNvars.docToTxt(doc, DocFile);
                                   //PublicFuncsNvars.saveDocToDB(ref fileData, Shotef, DocFile, ref comm, ref conn, text);
                                   comm = new SqlCommand("UPDATE dbo.documents SET Txt=@TextString, LastTxtUpdateDate=@dateTime WHERE shotef_mismach=@id", conn);

                                   comm.Parameters.AddWithValue("@id", Shotef);
                                   comm.Parameters.AddWithValue("@TextString", text);
                                   comm.Parameters.AddWithValue("@dateTime", DateTime.Now);

                                   //conn.Close();
                                   while (true)
                                   {
                                       try
                                       {
                                           conn.Open();
                                           try
                                           {
                                               comm.ExecuteNonQuery();
                                           }
                                           catch (Exception e)
                                           {
                                               MessageBox.Show(e.Message);
                                           }
                                           conn.Close();
                                           break;
                                       }
                                       catch { }
                                   }*/
                        }
                    }
                        
                    
                }
            }
        }
        
        private static void Log(string DocPath, int shotef, string v, string Message, string rowNum = "")
        {
            if (!string.IsNullOrEmpty(rowNum)) rowNum += ". ";
            using (StreamWriter sw = new StreamWriter(@"C:\temp\DocExp\" + v + ".log", true))
            {
                if (shotef > 0)
                    sw.WriteLine(rowNum + shotef + ": " + Message);
                else sw.WriteLine(rowNum + ": " + Message);
            }
        }

        private void btnDeleteFiles_Click(object sender, EventArgs e)
        {
            try
            {
                
                string DocPath = dirPath + "\\DocUpd";
                foreach (string file in Directory.GetFiles(DocPath))
                {
                    File.Delete(file);
                }
                if (Directory.GetFiles(DocPath).Length == 0)
                    MessageBox.Show("הקבצים נמחקו בהצלחה!");
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(dirPath + @"\DocExp\Exceptions.log", true))
                {
                    sw.WriteLine(ex.Message);
                }
            }
        }

        private void SaveTxt_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
        }
    }
}
