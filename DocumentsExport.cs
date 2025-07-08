using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;


namespace DocumentsModule
{
    public partial class DocumentsExport : Form
    {
        string conStr = Global.ConStr;
        List<int> documents = new List<int>();
        bool okID = false, //האם לסנן לפי שוטף
            okDate = false; //האם לסנן לפי תאריך
        private List<int> searchResults = new List<int>();
        public DocumentsExport()
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
                    MessageBox.Show("מסמך CVS נמצא בתיקיה.", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
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
            
            SqlCommand comm = new SqlCommand("; with d as (select shotef_mismach Shotef,0 Nispah, tarich_hamichtav Taarich, hanadon Nadon,simuchin Simuchin, hasAtts HasAtts  ,kod_sholeah Sholeah, file_extension Ext,file_data FileData"
            + " from Documents"
            + " where (shotef_mismach between @idMin and @idMax or @idMin+@idMax = 0)"
            + " and (tarich_hamichtav between @dateMin and @dateMax or(@dateMin = '00000000' and @dateMax = '00000000')))"
            + ", n as (select"
            + " n.shotef_mchtv,shotef_nisph,tarich,shm_kovtz, cast(IdentityKeyCol as varchar(15)) Simuchin,1 HasAtts,0 Sholeah,file_extension,file_data"
            + " from docnisp n"
            + " join d on n.shotef_mchtv = d.Shotef)"
            + " select *, (select count(*) from d) rows from d"
            + " union all"
            + " select *, (select count(*) from n) rows  from n", conn);
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
            while (true)
            {
                try
                {
                    sdr = comm.ExecuteReader();
                    break;
                }
                catch (Exception e)
                {}
            }

            try
            {
                string dirPath = "C:\\Temp";
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                string dateandtime = DateTime.Now.ToString("yyyyMMdd_HHmss");
                string CsvPath = dirPath + "\\DocumentsExp_"+dateandtime+".csv";
                Encoding encoding = Encoding.UTF8;
                using (StreamWriter writer = new StreamWriter(CsvPath, false,encoding))
                {
                    writer.WriteLine("Shotef, Nispah,Taarich,Nadon,Simuchin,HasAts,Sholeah,Ext");
                    progressBar1.Refresh();
                    progressBar1.Minimum = 0;
                    progressBar1.Value = 0;
                    progressBar1.Step = 1;
                    progressBar1.PerformStep();
                    int NumberofLines=0,lineNow = 0;
                    bool wasS= false, wasN = false;
                    int percent;
                    string w;
                    while (sdr.Read())
                    {
                        int Shotef = sdr.GetInt32(0);
                        int Nispah = sdr.GetInt32(1);
                        string Nadon = sdr.GetString(3).Trim().Replace(",",".");
                        //string Simuchin = sdr.GetString(4).Trim();//int HasAts = sdr.GetInt32(5);//int Sholeah = sdr.GetInt32(6);
                        string Ext = sdr.GetString(7).Trim();
                        string datas = sdr.GetString(2).Trim();
                        string line = ""+sdr.GetInt32(0) + "," + sdr.GetInt32(1) + ",\"" + datas + "\",\"" + Nadon + "\",\"" + sdr.GetString(4).Trim() + "\",\"" + sdr.GetInt32(5) + "\",\"" + sdr.GetInt32(6) + "\"," + sdr.GetString(7).Trim()+"";
                        if (!wasS&& Nispah==0)
                        {
                            wasS = true;
                            NumberofLines+= sdr.GetInt32(9);
                            progressBar1.Maximum = NumberofLines;
                        }
                        if (!wasN && Nispah != 0)
                        {
                            wasN = true;
                            NumberofLines += sdr.GetInt32(9);
                            progressBar1.Maximum = NumberofLines;
                        }
                        writer.WriteLine(line);
                        documents.Add(sdr.GetInt32(0));
                        if (checkBox1.Checked)
                        {
                            byte[] fileData = sdr.GetSqlBytes(8).Buffer;
                            string DocPath = dirPath + "\\DocExp";
                            if (!Directory.Exists(DocPath))
                                Directory.CreateDirectory(DocPath);
                            string DocFile = DocPath + "\\"+ Shotef;
                            if (Nispah == 0)
                                DocFile += "." + Ext;
                            else
                                DocFile += "_" + Nispah + "." + Ext.Replace(".","");
                            if (fileData!=null)
                                System.IO.File.WriteAllBytes(DocFile, fileData);
                        }
                        
                        progressBar1.Visible = true;
                        progressBar1.Update();
                        progressBar1.PerformStep();
                        progressBar1.Update();
                        lineNow += 1;
                        percent = (int)((lineNow / (double)progressBar1.Maximum) * 100);
                        w = percent.ToString() + "%" + "  " + lineNow.ToString() + "/" + NumberofLines.ToString();
                        progressBar1.CreateGraphics().DrawString(w, new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                        progressBar1.Update();
                    }
                }
            }
            catch (Exception e)
            {}
            
            conn.Close();
            Cursor = Cursors.Default;
            return documents.Count;
        }
                
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {}

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value < dateTimePicker2.Value)
            {
                MessageBox.Show("לא ניתן לבחור תאריך סיום הקודם לתאריך ההתחלה", "בחירת תאריך שגויה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {}

        private void DocumentsExport_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
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
    }
}
