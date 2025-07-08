using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace DocumentsModule
{
    public partial class DocumenstImport : Form
    {
        String filePathCVS;
        //string dirPath = "C:\\Temp\\DocExp\\";
        string docpath;
        int id, nispah;
        byte[] fileData;
        BackgroundWorker BG = new BackgroundWorker(); 
        //internal static string conStr = Properties.Settings.Default.MantakDBPConnectionString;

        public DocumenstImport()
        {
            InitializeComponent();
            filePath.Text = "C:\\Temp\\";
            folderPath.Text = @"C:\temp";
            BG.DoWork += BG_DoWork;
            BG.ProgressChanged += BG_ProgressChanged;
            BG.WorkerReportsProgress = true;
            BG.RunWorkerCompleted += BG_Complete;
            
        }

        private void BG_Complete(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar2.Value = 1;
            progressBar2.Visible = false;
            textBox1.Visible = false;
            MessageBox.Show("תהליך יבוא הסתיים.", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void BG_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int percent = e.ProgressPercentage;
            progressBar2.PerformStep();
            progressBar2.Update();
            progressBar2.CreateGraphics().DrawString(percent + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));
            //  progressBar2.Value = percent;
        }

        private void BG_DoWork(object sender, DoWorkEventArgs e)
        {
            int i = 0;
            foreach (string file in Directory.GetFiles(folderPath.Text))
            {
                bool isNispah = false;
                string[] splitPath = file.Split('\\');
                string[] splitFileNameExt = splitPath.Last().Split('.');
                string shotef = splitFileNameExt[0];
                string shotefN = string.Empty;
                string ext = string.Empty;
                
                if (splitFileNameExt.Length > 1)
                    ext = splitFileNameExt[1];

                if (shotef.Contains("_"))
                {
                    isNispah = true;
                    string[] splitShotef = shotef.Split('_');
                    shotef = splitShotef[0];
                    shotefN = splitShotef[1];
                }

                if (!int.TryParse(shotef, out id) || isNispah && !int.TryParse(shotefN, out nispah))
                {
                    MessageBox.Show("שם המסמך לא תקין, השוטף או הנספח לא נכונים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    BG.CancelAsync();
                }

                try
                {
                    fileData = File.ReadAllBytes(file);
                }
                catch
                {
                    MessageBox.Show("המסמך  פתוח", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return;
                }

                try
                {
                    if (!isNispah)
                    {
                        SqlConnection conn = new SqlConnection(Global.ConStr);
                        SqlCommand comm = new SqlCommand("UPDATE documents SET file_data=@data WHERE shotef_mismach=@id", conn);
                        comm.Parameters.AddWithValue("@id", id);
                        comm.Parameters.AddWithValue("@data", fileData);
                        conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        Log(i++, splitPath.Last(), isNispah);
                    }
                    else
                    {
                        SqlConnection conn = new SqlConnection(Global.ConStr);
                        SqlCommand comm = new SqlCommand("UPDATE docnisp SET file_data=@data WHERE shotef_mchtv=@id and shotef_nisph=@id_nisph", conn);
                        comm.Parameters.AddWithValue("@id", id);
                        comm.Parameters.AddWithValue("@id_nisph", nispah);
                        comm.Parameters.AddWithValue("@data", fileData);
                        conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        Log(i++, splitPath.Last(), isNispah);
                    }
                }
                catch (Exception ex)
                {
                    Log(i++, splitPath.Last(), isNispah, ex.Message);
                }
                int prc = (int)((i * 100) / Directory.GetFiles(folderPath.Text).Length);
                BG.ReportProgress(prc);
            }
        }

        private void BG_DoWork_CSV(object sender, DoWorkEventArgs e)
        {
            //-----------------------
            string filePathCVS = e.Argument.ToString();
            bool isFileCSVexp;
            using (StreamReader reader = new StreamReader(filePathCVS))
            {
                isFileCSVexp = false;
                try
                {
                    isFileCSVexp = reader.ReadLine().Contains("Shotef");
                }
                catch
                {
                    isFileCSVexp = false;
                    // cancel!
                }

                if (isFileCSVexp)
                {
                    int i = 0;


                    while (!reader.EndOfStream)
                    {

                        //if (lineNow % 1000 == 0) Thread.Sleep(5000);
                        string line = reader.ReadLine();
                        string[] values = line.Split(',');
                        try
                        {
                            id = Convert.ToInt32(values[0]);
                            nispah = Convert.ToInt32(values[1]);
                        }
                        catch
                        {
                            MessageBox.Show("המסך לא תקין, השוטף או הנספח לא נכונים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            return;
                        }
                        if (!Directory.Exists(folderPath.Text))
                        {
                            MessageBox.Show("התקיה  לC:\\Temp\\DocExp\\א קיימת.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            return;
                        }
                        bool isNispah = false;
                        if (nispah == 0)
                            docpath = folderPath.Text + id.ToString() + "." + values[7].Replace(".", "");
                        else
                        {
                            docpath = folderPath.Text + id.ToString() + "_" + nispah.ToString() + "." + values[7].Replace(".", "");
                            isNispah = true;
                        }

                        if (File.Exists(docpath))
                        {
                            try
                            {
                                fileData = File.ReadAllBytes(docpath);
                            }
                            catch
                            {
                                MessageBox.Show("המסמך  פתוח", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                return;
                            }
                            try
                            {
                                if (!isNispah)
                                {
                                    SqlConnection conn = new SqlConnection(Global.ConStr);
                                    SqlCommand comm = new SqlCommand("UPDATE documents SET file_data=@data WHERE shotef_mismach=@id", conn);
                                    comm.Parameters.AddWithValue("@id", id);
                                    comm.Parameters.AddWithValue("@data", fileData);
                                    conn.Open();
                                    comm.ExecuteNonQuery();
                                    conn.Close();
                                    Log(i++, id, isNispah);
                                }
                                else
                                {
                                    SqlConnection conn = new SqlConnection(Global.ConStr);
                                    SqlCommand comm = new SqlCommand("UPDATE docnisp SET file_data=@data WHERE shotef_mchtv=@id and shotef_nisph=@id_nisph", conn);
                                    comm.Parameters.AddWithValue("@id", id);
                                    comm.Parameters.AddWithValue("@id_nisph", nispah);
                                    comm.Parameters.AddWithValue("@data", fileData);
                                    conn.Open();
                                    comm.ExecuteNonQuery();
                                    conn.Close();
                                    Log(i++, id, isNispah);
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(i++, id, isNispah, ex.Message);
                            }
                        }
                        int percentage = (int)((i / (double)progressBar2.Maximum) * 100);
                        BG.ReportProgress(percentage, progressBar2.Maximum);

                    }

                }
                else
                {
                    MessageBox.Show("המסך לא תקין", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }

            //-----------------------
            
                
            
        }

        private void filePath_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void DocumenstImport_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            filePathCVS = filePath.Text;
            if (File.Exists(filePathCVS))
            {
                if (filePathCVS.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    //Encoding encoding = Encoding.UTF8;
                    try
                    {
                        using (StreamReader reader = new StreamReader(filePathCVS))
                        {
                            bool isFileCSVexp = false;
                            try
                            {
                                isFileCSVexp = reader.ReadLine().Contains("Shotef");
                            }
                            catch
                            {
                                isFileCSVexp = false;
                            }

                            if (isFileCSVexp)
                            {
                                progressBar2.Refresh();
                                progressBar2.Visible = true;
                                progressBar2.Minimum = 1;
                                progressBar2.Value = 1;
                                progressBar2.Step = 1;
                                int lengthOfLines = File.ReadAllLines(filePathCVS).Length;
                                lengthOfLines -= 1;
                                progressBar2.Maximum = lengthOfLines;
                                int lineNow = 0;
                                int percent = (int)((lineNow / (double)progressBar2.Maximum) * 100);
                                string w = percent.ToString() + "%" + "  " + progressBar2.Value.ToString() + "/" + lengthOfLines.ToString();
                                progressBar2.CreateGraphics().DrawString(w, new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));

                                while (!reader.EndOfStream)
                                {

                                    string line = reader.ReadLine();
                                    string[] values = line.Split(',');
                                    try
                                    {
                                        id = Convert.ToInt32(values[0]);
                                        nispah = Convert.ToInt32(values[1]);
                                    }
                                    catch
                                    {
                                        MessageBox.Show("המסך לא תקין, השוטף או הנספך לא נכונים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                        return;
                                    }
                                    if (!Directory.Exists(folderPath.Text))
                                    {
                                        MessageBox.Show("התקיה  לC:\\Temp\\DocExp\\א קיימת.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                        return;
                                    }
                                    bool isNispah = false;
                                    if (nispah == 0)
                                        docpath = folderPath.Text + id.ToString() + "." + values.Last().Replace(".", "");
                                    else
                                    {
                                        docpath = folderPath.Text + id.ToString() + "_" + nispah.ToString() + "." + values.Last().Replace(".", "");
                                        isNispah = true;
                                    }

                                    if (File.Exists(docpath))
                                    {
                                        try
                                        {
                                            if (values.Last().Replace(".", "").ToLower() == "doc" || values.Last().Replace(".", "").ToLower() == "docx")
                                            {
                                                fileData = File.ReadAllBytes(docpath);
                                                if (!isNispah)
                                                {
                                                    if (values.Last().Replace(".", "").ToLower() == "doc" || values.Last().Replace(".", "").ToLower() == "docx")
                                                    {
                                                        SqlConnection conn = new SqlConnection(Global.ConStr);
                                                        SqlCommand comm = new SqlCommand("UPDATE top(1) documents SET file_data=@data WHERE shotef_mismach=@id", conn);
                                                        Word.Application wapp;
                                                        try
                                                        {
                                                            wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                                                        }
                                                        catch
                                                        {
                                                            wapp = new Word.Application();
                                                        }
                                                        Word.Document doc = wapp.Documents.Open(docpath);
                                                        string text = PublicFuncsNvars.docToTxt(doc, docpath);
                                                        PublicFuncsNvars.saveDocToDB(ref fileData, id, docpath, ref comm, ref conn, text);

                                                    }

                                                }
                                            }
                                                
                                        }
                                        catch
                                        {
                                            MessageBox.Show("המסמך  פתוח" + docpath, "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                            //return;
                                        }
                                        
                                    }
                                    progressBar2.PerformStep();
                                    progressBar2.Update();
                                    lineNow += 1;
                                    percent = (int)((lineNow / (double)progressBar2.Maximum) * 100);
                                    w = percent.ToString() + "%" + "  " + progressBar2.Value.ToString() + "/" + lengthOfLines.ToString();
                                    progressBar2.CreateGraphics().DrawString(w, new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));


                                }
                                progressBar2.Value = 1;
                                progressBar2.Visible = false;
                                textBox1.Visible = false;
                                MessageBox.Show("שמירת טקסט עבר בהצלחה", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            }
                            else
                            {
                                MessageBox.Show("המסך לא תקין", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("מסמך CSV פתוח.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }

                }
                else
                {
                    MessageBox.Show("המסמך הזה הוא לא מסמך CVS.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
            else
            {
                MessageBox.Show("מיקום של הקובץ CVS לא קיים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog f = new OpenFileDialog())
            {
                f.Filter = "CSV files(*.csv)|*.csv|All files (*.*)|*.*";
                f.Title = "תבחר CVS קובץ.";
                if (f.ShowDialog() == DialogResult.OK)
                {
                    filePath.Text = f.FileName;
                    folderPath.Text = Path.GetDirectoryName(filePath.Text);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog f = new FolderBrowserDialog())
            {
                f.Description = "תבחר תיקיה.";
                f.SelectedPath = @"C:\temp";
                DialogResult result = f.ShowDialog();
                if (result== DialogResult.OK && !string.IsNullOrWhiteSpace(f.SelectedPath))
                    folderPath.Text = f.SelectedPath;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!folderPath.Text.EndsWith("\\"))
            {
                folderPath.Text += "\\";
            }
            if (!Directory.Exists(folderPath.Text))
            {
                MessageBox.Show("התקיה  לC:\\Temp\\DocExp\\א קיימת.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }

            if (!BG.IsBusy)
            {
                progressBar2.Visible = true;
                progressBar2.Minimum = 0;
                progressBar2.Value = 1;
                progressBar2.Step = 1;
                int lengthOfLines = Directory.GetFiles(folderPath.Text).Length - 1;

                progressBar2.Maximum = lengthOfLines;

                BG.RunWorkerAsync();
            }
            
        }

        private void button1_Click_CSV(object sender, EventArgs e)
        {
            if (!folderPath.Text.EndsWith("\\"))
            {
                folderPath.Text += "\\";
            }
            if (!Directory.Exists(folderPath.Text))
            {
                MessageBox.Show("התקיה  לC:\\Temp\\DocExp\\א קיימת.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }
            filePathCVS = filePath.Text;
            if (File.Exists(filePathCVS))
            {
                if (filePathCVS.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    //Encoding encoding = Encoding.UTF8;
                    try
                    {
                        progressBar2.Refresh();
                        progressBar2.Visible = true;
                        progressBar2.Minimum = 0;
                        progressBar2.Value = 1;
                        progressBar2.Step = 1;
                        int lengthOfLines = File.ReadAllLines(filePathCVS).Length;
                        lengthOfLines -= 1;
                        progressBar2.Maximum = lengthOfLines;
                        int lineNow = 0;
                        int percent = (int)((lineNow / (double)progressBar2.Maximum) * 100);
                        string w = percent.ToString() + "%" + "  " + progressBar2.Value.ToString() + "/" + lengthOfLines.ToString();
                        progressBar2.CreateGraphics().DrawString(w, new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));

                        if (!BG.IsBusy)
                            BG.RunWorkerAsync(argument: filePathCVS);
                    }
                    catch
                    {
                        MessageBox.Show("מסמך CSV פתוח.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    
                }
                else
                {
                    MessageBox.Show("המסמך הזה הוא לא מסמך CVS.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
            else
            {
                MessageBox.Show("מיקום של הקובץ CVS לא קיים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void Log(int i, int id, bool nispah, string message = "")
        {
            string DocType = "Document ";
            if (nispah) DocType = "Attatchment ";
            string logFolder = folderPath.Text + @"Log\";
            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);

            using (StreamWriter sw = new StreamWriter(logFolder + "DocumentImport.log", true))
            {
                if (string.IsNullOrEmpty(message))
                    sw.WriteLine(i + ". " + DocType + id + " has succesfuly uploaded");
                else sw.WriteLine(i + ". " + DocType + id + " has failed");
            }
        }

        private void Log(int i, string fileName, bool nispah, string message = "")
        {
            string DocType = "Document ";
            if (nispah) DocType = "Attatchment ";
            string logFolder = folderPath.Text + @"Log\";
            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);

            using (StreamWriter sw = new StreamWriter(logFolder + "DocumentImport.log", true))
            {
                if (string.IsNullOrEmpty(message))
                    sw.WriteLine(i + ". " + DocType + fileName + " has succesfuly uploaded");
                else sw.WriteLine(i + ". " + DocType + fileName + " has failed");
            }
        }
    }
}
//update top(1) documents set file_data=cast(N'333' as varbinary) where shotef_mismach=7787878788
//update top (1) docnisp set file_data=cast(N'111' as varbinary) where shotef_mchtv=str1 and shotef_nisph=str2