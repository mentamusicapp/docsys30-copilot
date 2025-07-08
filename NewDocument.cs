using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordTools = Microsoft.Office.Tools.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace DocumentsModule
{
   
    public partial class NewDocument : Form
    {
        bool topaste = false;
        Microsoft.Office.Interop.Outlook.Application app = null;
        Document curDoc;
        Process scandAll=new Process();
        Word.Application Wapp = null;
        Word.Document doc;
        bool success = false;
        Excel.Application Eapp;
        Excel.Workbook wb;
        Dictionary<int, string> projects;
        Dictionary<int, string> patterns;
        string strTyped = "";
        Boolean isDragged = false;
        int row1 = 0, row2 = 0, row3 = 0, row5 = 0;
        private List<User> users;
        private List<Folder> directories;
        bool okSender = false, okDir = false, okPro = false, okRef = false, okPat = false, wasAdded = true;
        byte[] fileData;
        string extension;
        enum WayIn { word, excel, scan, file, existing, drag };
        WayIn curDocWI = WayIn.word;
        private bool ok=true;
        private string path;
        string mailPath = "";
        private Word.Row r;
        private bool isExcelClosed = false;
        private bool startToLookAtIndex0;
        string TableType;
        string Ujob;
        DocumentHandling dh = null;
        public NewDocument()
        {
            scandAll.EnableRaisingEvents = true;
            scandAll.Exited += scandAll_Exited;

            InitializeComponent();
            string iconpath = System.Windows.Forms.Application.ExecutablePath;
            this.DoubleBuffered = true;
            ChangeDataGrid("users");
            button5.Visible = false;
            textBox24.Visible = false;
            label2.Visible = false;
            dataGridView2.Visible = false;
            DocumentsMenu.PathTemplate(this.button1,55);
            DocumentsMenu.PathTemplate(this.button2, 55);
            DocumentsMenu.PathTemplate(this.button3, 55);
            DocumentsMenu.PathTemplate(this.button15, 55);
            DocumentsMenu.PathTemplate(this.button14, 35);
            DocumentsMenu.PathTemplate(this.button13, 35);
            DocumentsMenu.PathTemplate(this.button12, 30);
            dataGridView2.CellDoubleClick += dataGridViewTable_CellDoubleClick;
            dataGridView2.RowsAdded += DataGridView2_RowsChanged;
            dataGridView2.RowsRemoved += DataGridView2_RowsChanged;

        }

        private void scandAll_Exited(object sender, EventArgs e)
        {
            //open pdf file
        }

        private void button2_Click_MEKORI(object sender, EventArgs e)
        {
            curDocWI = WayIn.scan;
            panel1.Visible = true;
            customLabel2.Visible = false;// מספר שוטף
            textBox17.Visible = false;// מספר שוטף
            button17.Visible = false;// חפש (שוטף)
            changeDataControlsVisiblity(false);
            changeTemplateControlsVisiblity(false);
            try
            {
                if (Directory.Exists("C:\\Button_Data\\PDF"))
                {
                    string[] files = Directory.GetFiles("C:\\Button_Data\\PDF");
                    foreach (string f in files)
                    {
                        File.Delete(f);
                    }
                    Directory.Delete("C:\\Button_Data\\PDF");
                }
                Directory.CreateDirectory("C:\\Button_Data\\PDF");
                Process p = new Process();
                p.StartInfo.FileName = "C:\\Program Files\\Avision\\Button Manager\\ButtonManager.exe";
                p.StartInfo.Arguments = "/gogoscan";
                p.Start();
                while (Directory.GetFiles("C:\\Button_Data\\PDF").Count() == 0) ;
                long length;

                while (true)
                {
                    try
                    {
                        string[] scannedFiles = Directory.GetFiles("C:\\Button_Data\\PDF");
                        path = scannedFiles[0];
                        FileInfo fi = new FileInfo(path);
                        length = fi.Length;
                        Thread.Sleep(5000);
                        while (length != (new FileInfo(path)).Length)
                        {
                            fi = new FileInfo(path);
                            length = fi.Length;
                            Thread.Sleep(5000);
                        }
                        extension = "pdf";
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(1000);
                        continue;
                    }
                }
 
                changeDataControlsVisiblity(true);
                changeInputControlsVisiblity(true);
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = "99999";
                textBox2.TextChanged += textBox2_TextChanged;
            }
            catch (Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show(ex.Message);
                MessageBox.Show("בבקשה סגרו את כל קבצי ה-PDF הפתוחים מתקיית ה-button manager.");
            }
        }

        private void button2_Click(object sender, EventArgs e)// סריקה ל pdf //לא בשימוש
        {
            dataGridView2.Visible = false;
            button5.Visible = false;
            textBox24.Visible = false;
            label2.Visible = false;
            pnlWayInFiles.Visible = true;
            panel_dropDown.Height = 61;
            this.TopMost = false;
        }

        private void CreateFromScan()//לא בשימוש
        {
            string C_Button_Data_PDF = @"C:\Button_Data\PDF";
            string C_USER_Desktop_Button_Data = @"C:\Users\" + Environment.UserName + @"\Desktop\Button Data";
            curDocWI = WayIn.scan;
            panel1.Visible = true;
            customLabel2.Visible = false;// מספר שוטף
            textBox17.Visible = false;// מספר שוטף
            button17.Visible = false;// חפש (שוטף)
            changeDataControlsVisiblity(false);
            changeTemplateControlsVisiblity(false);
            changeInputControlsVisiblity(false);
            try
            {
                if (Directory.Exists(C_Button_Data_PDF))
                {
                    string[] files = Directory.GetFiles(C_Button_Data_PDF);
                    bool isex = false;
                    foreach (string f in files)
                    {
                        try
                        {
                            File.Delete(f);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(@"נכשל במחיקת קבצים מתיקיה " + C_Button_Data_PDF + "\nיש למחוק תיקיה זו ידנית", "מערכת ניהול מסמכים", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            isex = true;
                        }
                    }
                    try
                    {
                        Directory.Delete(@"C:\Button_Data\PDF");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(@"נכשל במחיקת תיקיה " + C_Button_Data_PDF + "\nיש למחוק תיקיה זו ידנית", "מערכת ניהול מסמכים", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        isex = true;
                    }
                    if (isex) return;
                }
                if (Directory.Exists(C_USER_Desktop_Button_Data))
                {
                    string[] files = Directory.GetFiles(C_USER_Desktop_Button_Data);
                    foreach (string f in files)
                    {
                        File.Delete(f);
                    }
                    Directory.Delete(C_USER_Desktop_Button_Data);
                }
                Directory.CreateDirectory(C_Button_Data_PDF);
                Directory.CreateDirectory(C_USER_Desktop_Button_Data);

                Process p = new Process();
                p.StartInfo.FileName = "C:\\Program Files\\Avision\\Button Manager\\ButtonManager.exe";
                p.StartInfo.Arguments = "/gogoscan";
                p.Start();
                DialogResult dr = MessageBox.Show("סרוק את המסמכים כעת\nלסיום לחץ על אישור", 
                    "קליטת מסמך", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, 
                    MessageBoxDefaultButton.Button1, 
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                if (dr != System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }
                long length;
                string myDir = C_Button_Data_PDF;
                while (true)
                {
                    try
                    {
                        string[] scannedFiles = Directory.GetFiles(C_Button_Data_PDF);
                        if (scannedFiles.Length != 0)
                        {
                            path = scannedFiles[0];
                            FileInfo fi = new FileInfo(path);
                            length = fi.Length;
                            Thread.Sleep(5000);
                            while (length != (new FileInfo(path)).Length)
                            {
                                fi = new FileInfo(path);
                                length = fi.Length;
                                Thread.Sleep(5000);
                            }
                            extension = "pdf";
                            break;
                        }
                        else
                        {
                            MessageBox.Show("תיקיה '" + myDir + "' ריקה", "קליטת מסמך", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            return;
                        }

                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message + "\n" + ex.StackTrace, ex.GetType().ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                changeDataControlsVisiblity(true);
                changeInputControlsVisiblity(true);
                changeWriterControlsVisibility(false);
                changeSelectedFileControlsVisibility(true);

                textBox19.Text = "0";
                textBox18.Text = string.Empty;
                textBox21.Text = string.Empty;
                textBox20.Text = string.Empty;
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = "99999";
                textBox2.TextChanged += textBox2_TextChanged;
                textBox13.Text = path;
            }

            catch (Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show(ex.Message + " \nStackTrace: " + ex.StackTrace);

                MessageBox.Show("בבקשה סגרו את כל קבצי ה-PDF הפתוחים מתקיית ה-button manager.");
            }
            
        }

        private void button1_Click(object sender, EventArgs e)// מסמך Word חדש
        {
            dataGridView2.Visible = false;
            button5.Visible = false;
            label2.Visible = false;
            textBox24.Visible = false;
            curDocWI = WayIn.word;
          
            panel1.Visible = true;
            pnlWayInFiles.Visible = false;
            changeShotefContrlVisibility(false);
            changeDataControlsVisiblity(true);
            changeTemplateControlsVisiblity(true);
            changeInputControlsVisiblity(false);
            changeSelectedFileControlsVisibility(false);
            panel_dropDown.Height = 61;
            this.TopMost = false;
        }

        private void changeShotefContrlVisibility(bool b)
        {
            customLabel2.Visible = b; // מספר שוטף
            textBox17.Visible = b;// מספר שוטף
            button17.Visible = b; // חפש (מספר שוטף)
        }

        private void changeTemplateControlsVisiblity(bool b)
        {
            label13.Visible = b;// תבנית מסמך
            textBox14.Visible = b;// קוד תבנית
            textBox15.Visible = b;// שם תבנית
        }


        private void changeNadonControlsVisiblity(bool b)
        {
            lblNadon.Visible = b; // הנדון
            textBox1.Visible = b;// נדון
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ChangeDataGrid("users");
            dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
            if (!textBox2.Text.Equals("קוד") && textBox2.Text.Equals(""))
            {
                foreach (User u in users) // יערה 19.12.22 - שליפת מפקד המשתמש הפעיל במקום החותם המקורי
                {
                    if (u.userCode.ToString() == PublicFuncsNvars.curUser.userCode.ToString())
                    {
                        textBox2.Text = u.commanderCode.ToString();
                        textBox3.Text = u.firstName + " " + u.lastName + " - " + u.job;
                        Ujob = u.job;
                        //PublicFuncsNvars.nameNjobByCode(ref textBox2, ref textBox4, ref textBox16, ref textBox3);//העברתי פנימה
                        break;
                    }
                }
            }

            if (!textBox2.Text.Equals("קוד") && !textBox2.Text.Equals(""))
            {
                int res;
                if (int.TryParse(textBox2.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            textBox3.Text = u.firstName + " " + u.lastName + " - " + u.job;
                            Ujob = u.job;
                        }
                    }
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox2.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox2.Text))
                {
                    MessageBox.Show("אין משתמש עם יוזר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox2.Text = "";
                }
                //dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                /*int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox2.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox2.Text))
                {
                    MessageBox.Show("אין משתמש עם יוזר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox7.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox16.Text = "";

                }
                else if (index!=0 && dataGridView2.Rows[index].Cells[0].Value.ToString().StartsWith(textBox2.Text))
                {
                    textBox3.Text=dataGridView2.Rows[index].Cells[3].Value.ToString();
                    textBox4.Text = dataGridView2.Rows[index].Cells[2].Value.ToString();
                    textBox16.Text = dataGridView2.Rows[index].Cells[1].Value.ToString();
                }*/
                /*if(textBox2.Text=="1")
                {
                    textBox11.Text = "מפ - 1";
                }
                else if(textBox2.Text=="2")
                {
                    textBox11.Text = "לש - 320";
                }*/
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            ChangeDataGrid("folders");
            textBox11.TextChanged -= textBox11_TextChanged;
            PublicFuncsNvars.directoryByCode(ref textBox12, ref textBox11, ref textBox5, ref textBox6, "קוד", "שם מקוצר",
                "SELECT shm_mshimh, shm_mkotzr FROM dbo.tm_mesimot WHERE ms_mshimh=@id AND shm_mkotzr<>''", "@id", typeof(int));
            textBox11.TextChanged += textBox11_TextChanged;
            textBox7.Text = textBox11.Text + " - " + textBox6.Text;
            if (!textBox12.Text.Equals("קוד") && !textBox12.Text.Equals(""))
            {

                //textBox24.Text = textBox12.Text;
                //dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                List<Folder> FilteredDirectories = directories.Where(word => word.id.ToString().StartsWith(textBox12.Text)).ToList();
                dataGridView2.Columns[0].DataPropertyName = "id";
                dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                dataGridView2.Columns[2].DataPropertyName = "description";
                dataGridView2.DataSource = FilteredDirectories.Select(item => new
                {
                    item.id,
                    item.shortDescription,
                    item.description
                }).ToList();
                dataGridView2.Refresh();
                int numOfRows = dataGridView2.Rows.Cast<DataGridViewRow>().Count(row => row.Visible);
                int index = 0;
                /*foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null  && row.Cells[0].Value.ToString().StartsWith(textBox12.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }*/
                if (numOfRows > 0)
                {
                    dataGridView2.FirstDisplayedScrollingRowIndex = index;
                }
                //dataGridView2.FirstDisplayedScrollingRowIndex = index;
                else
                {
                    MessageBox.Show("אין תיק עם מספר זה במערכת התואם לחתך התיקים שנבחר.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox12.Text = "";
                    dataGridView2.DataSource = directories.Select(item => new
                    {
                        item.id,
                        item.shortDescription,
                        item.description
                    }).ToList();
                    dataGridView2.Refresh();
                }
                /*if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox12.Text))
                {
                    MessageBox.Show("אין תיק עם מספר זה במערכת התואם לחתך התיקים שנבחר.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox12.Text = "";
                }*/
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            ChangeDataGrid("folders");
            textBox12.TextChanged -= textBox12_TextChanged;
            PublicFuncsNvars.directoryByCode(ref textBox11, ref textBox12, ref textBox5, ref textBox6, "שם מקוצר", "קוד",
                "SELECT shm_mshimh, ms_mshimh FROM dbo.tm_mesimot WHERE shm_mkotzr=@shortName", "@shortName", typeof(string));
            textBox12.TextChanged += textBox12_TextChanged;
            textBox7.Text = textBox11.Text + " - " + textBox6.Text;
            if (!textBox11.Text.Equals("שם מקוצר") && !textBox11.Text.Equals(""))
            {
                List<Folder> FilteredDirectories = directories.Where(word => word.shortDescription.ToString().StartsWith(textBox11.Text)).ToList();
                //dataGridView2.Columns[0].DataPropertyName = "id";
                //dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                //dataGridView2.Columns[2].DataPropertyName = "description";
                dataGridView2.DataSource = FilteredDirectories.Select(item => new
                {
                    item.id,
                    item.shortDescription,
                    item.description
                }).ToList();
                dataGridView2.Refresh();
                int numOfRows = dataGridView2.Rows.Cast<DataGridViewRow>().Count(row => row.Visible);
                //int index = 0;
                if (numOfRows > 0)
                {
                    dataGridView2.FirstDisplayedScrollingRowIndex = 0;
                }
                else
                {
                    MessageBox.Show("אין תיק עם שם מקוצר זה במערכת התואם לחתך התיקים שנבחר.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox12.Text = "";
                    dataGridView2.DataSource = directories.Select(item => new
                    {
                        item.id,
                        item.shortDescription,
                        item.description
                    }).ToList();
                    dataGridView2.Refresh();
                }




                /*dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);
                int index = dataGridView2.FirstDisplayedScrollingRowIndex;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[1].Value != null && row.Visible && row.Cells[1].Value.ToString().StartsWith(textBox11.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[1].Value.ToString().StartsWith(textBox11.Text))
                {
                    MessageBox.Show("אין תיק עם שם מקוצר זה במערכת התואם לחתך התיקים שנבחר.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox11.Text = "";
                }*/
            }
        }

        private void textBox11_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox11.Text.Equals("שם מקוצר"))
            {
                textBox11.Text = "";
            }
        }

        private void textBox11_Leave(object sender, EventArgs e)
        {
            if (textBox11.Text.Equals(""))
            {
                textBox11.Text = "שם מקוצר";
            }
        }

        private void textBox12_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox12.Text, out res))
            {
                textBox12.Text = "";
            }
        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            if (textBox12.Text.Equals(""))
                textBox12.Text = "קוד";
        }

        private void NewDocument_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
            /*users = PublicFuncsNvars.users.Where(x=>x.isActive).ToList();
            foreach (User u in users)
                dataGridViewUsers.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
                */
            int folders_count = PublicFuncsNvars.folders.Count;
            object isactive = PublicFuncsNvars.folders.Where(x => x.isActive);

            /*directories = PublicFuncsNvars.folders.Where(x => x.isActive && (x.branch == PublicFuncsNvars.curUser.branch ||
                PublicFuncsNvars.curUser.roleType == RoleType.computers || x.shortDescription == "מפ - 1" || x.shortDescription == "לש - 320")).ToList();
            foreach (Folder d in directories)
                dataGridViewFolders.Rows.Add(d.id, d.shortDescription, d.description);
                
            projects = PublicFuncsNvars.projects;
            foreach (KeyValuePair<int, string> p in projects)
                dataGridViewProjects.Rows.Add(p.Key.ToString(), p.Value);
*/
            patterns = new Dictionary<int, string>();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT onum_word_template, dscr_template,nam_word_template FROM dbo.tm_templ_bhi", conn); //WHERE file_data<>0x00000000
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                patterns.Add(sdr.GetInt16(0), sdr.GetString(1).Trim());
                //dataGridViewPatterns.Rows.Add(sdr.GetInt16(0), sdr.GetString(1).Trim(), sdr.GetString(2).Trim());
            }
            conn.Close();
            comboBox1.DataSource = PublicFuncsNvars.sivug_by_reshet();
            comboBox1.SelectedIndex = 0;
            textBox14.Text = "1";
            textBox19.Text = PublicFuncsNvars.curUser.userCode.ToString();
            textBox20.Text = $"{PublicFuncsNvars.curUser.firstName} {PublicFuncsNvars.curUser.lastName} - {PublicFuncsNvars.curUser.job}";
            //textBox21.Text = PublicFuncsNvars.curUser.firstName;
            //textBox18.Text = PublicFuncsNvars.curUser.lastName;
            if (PublicFuncsNvars.curUser.branch != Branch.development)
                textBox9.ReadOnly = true;// קוד פרויקט



        }

        private void button3_Click(object sender, EventArgs e)//לא בשימוש// חוברת Excel חדשה
        {
            dataGridView2.Visible = false;
            button5.Visible = false;
            textBox24.Visible = false;
            label2.Visible = false;
            curDocWI = WayIn.excel;
            panel1.Visible = true;
            customLabel2.Visible = false;// מספר שוטף
            textBox17.Visible = false;// מספר שוטף
            button17.Visible = false;// חפש (שוטף)
            changeDataControlsVisiblity(true);
            changeTemplateControlsVisiblity(false);
            changeInputControlsVisiblity(false);
            changeSelectedFileControlsVisibility(false);
            panel_dropDown.Height = 61;
            this.TopMost = false;
        }

        private void CreateFromBrowse()//לא בשימוש
        {
            curDocWI = WayIn.file;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            DialogResult res = ofd.ShowDialog();
            if (res == DialogResult.OK)
            {
                path = ofd.FileName;
                textBox13.Text = path;
                panel1.Visible = true;
                customLabel2.Visible = false;// מספר שוטף
                textBox17.Visible = false;// מספר שוטף
                button17.Visible = false;// חפש (שוטף)
                changeDataControlsVisiblity(true);
                changeInputControlsVisiblity(true);
                changeWriterControlsVisibility(false);
                changeSelectedFileControlsVisibility(true);
                string[] pathArr = path.Split('\\');
                string fileName = pathArr[pathArr.Length - 1];
                textBox1.Text = fileName.Substring(0, fileName.LastIndexOf('.'));
                extension = fileName.Substring(fileName.LastIndexOf('.') + 1);
                fileData = File.ReadAllBytes(path);
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = "99999";
                textBox2.TextChanged += textBox2_TextChanged;
            }
        }

        private void button5_Click(object sender, EventArgs e)// יצירת שוטף מקובץ קיים
        {
            curDocWI = WayIn.file;
            panel1.Visible = true;
            customLabel2.Visible = false;// מספר שוטף
            textBox17.Visible = false;// מספר שוטף
            button17.Visible = false;// חפש (שוטף)
            changeDataControlsVisiblity(true);
            changeTemplateControlsVisiblity(false);
            changeInputControlsVisiblity(false);
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            DialogResult res = ofd.ShowDialog();
            if(res==DialogResult.OK)
            {
                path = ofd.FileName;
                textBox13.Text = path;
                changeSelectedFileControlsVisibility(true);
                string[] pathArr = path.Split('\\');
                string fileName = pathArr[pathArr.Length - 1];
                textBox1.Text = fileName.Substring(0, fileName.LastIndexOf('.'));
                extension = fileName.Substring(fileName.LastIndexOf('.') + 1);
                fileData = File.ReadAllBytes(path);
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = "99999";
                textBox2.TextChanged += textBox2_TextChanged;
            }
        }

        private void changeSelectedFileControlsVisibility(bool b)
        {
            label11.Visible = b;// הקובץ הנבחר
            textBox13.Visible = b;// הקובץ הנבחר
        }

        private void NewDocument_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int res;
                if (int.TryParse(textBox9.Text, out res))
                    textBox10.Text = projects[res];
                else
                    textBox10.Text = "שם פרויקט";
            }
            catch
            {
                textBox9.Text = "";
            }
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox9.Text, out res))
            {
                textBox9.Text = "";
            }
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
                textBox9.Text = "קוד";
        }

        private void selectingSender()
        {
            //label24.Visible = true;
            //dataGridViewUsers.Visible = true;
            ChangeDataGrid("users");
            //button6.Visible = true;
            //button8.Visible = true;
            //makeDirectoriesTableInVisible();
            //makeProjectsTableInVisible();
            //makePatternsTableInVisible();
            okSender = true;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox12_Click(object sender, EventArgs e)
        {
            ChangeDataGrid("folders");
            //label25.Visible = true;
            //textBox22.Visible = true;
            //button16.Visible = true;
            //dataGridViewFolders.Visible = true;
            //button7.Visible = true;
            //button9.Visible = true;
            //changeUsersTableVisiblity(false);
            //makeProjectsTableInVisible();
            //makePatternsTableInVisible();
            okDir = true;
        }

        private void makeProjectsTableInVisible()
        {
            Control[] controls = { label9, dataGridViewProjects, button10, button11 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox2.Text, out res))
            {
                textBox2.Text = "";
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text.Equals(""))
                textBox2.Text = "קוד";
        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            if (PublicFuncsNvars.curUser.branch == Branch.development)
            {
                ChangeDataGrid("projects");
                //label9.Visible = true;
                //dataGridViewProjects.Visible = true;
                //button10.Visible = true;
                //button11.Visible = true;
                //changeUsersTableVisiblity(false);
                //makeDirectoriesTableInVisible();
                //makePatternsTableInVisible();
                okPro = true;
            }
        }

        private void makeDirectoriesTableInVisible()
        {
            Control[] controls={ label25, dataGridViewFolders, button7, button9, textBox22, button16};
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void changeUsersTableVisiblity(bool b)
        {
            Control[] controls = { label24, dataGridViewUsers, button6, button8 };
            PublicFuncsNvars.changeControlsVisiblity(b, controls.ToList());
        }

        private void dataGridView3_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridViewUsers.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewUsers.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
                    break;
                }
            }
        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridViewFolders.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewFolders.Rows)
            {
                if (row.Cells[col].Value != null && row.Visible && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
                    break;
                }
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView3_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridView2_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridView2_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row2 = e.RowIndex;
                dataGridViewFolders.Rows[row2].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridView3_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button==MouseButtons.Right)
            {
                row3 = e.RowIndex;
                dataGridViewUsers.Rows[row3].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (okSender)
            {
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox2.TextChanged += textBox2_TextChanged;
                
                textBox3.TextChanged -= textBox3_TextChanged;
                textBox3.Text = $"{dataGridViewUsers.SelectedCells[0].OwningRow.Cells[1].Value.ToString()} {dataGridViewUsers.SelectedCells[0].OwningRow.Cells[2].Value.ToString()} - {dataGridViewUsers.SelectedCells[0].OwningRow.Cells[3].Value.ToString()}";
                textBox3.TextChanged += textBox3_TextChanged;
                Ujob = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                /*textBox4.TextChanged -= textBox4_TextChanged;
                textBox4.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox4.TextChanged += textBox4_TextChanged;

                textBox16.TextChanged -= textBox16_TextChanged;
                textBox16.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                textBox16.TextChanged += textBox16_TextChanged;*/
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (okDir)
            {
                textBox12.TextChanged -= textBox12_TextChanged;
                textBox12.Text = dataGridViewFolders.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox12.TextChanged += textBox12_TextChanged;
                textBox11.TextChanged -= textBox11_TextChanged;
                textBox11.Text = dataGridViewFolders.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox11.TextChanged += textBox11_TextChanged;
                textBox5.Text = dataGridViewFolders.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                    "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@id", conn);
                comm.Parameters.AddWithValue("@id", int.Parse(textBox12.Text));
                conn.Open();
                textBox6.Text = (int.Parse(comm.ExecuteScalar().ToString()) + 1).ToString();
                conn.Close();
                textBox7.Text = textBox11.Text + " - " + textBox6.Text;
            }
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            ChangeDataGrid("folders");
            //label25.Visible = true;
            //textBox22.Visible = true;
            //button16.Visible = true;
            //dataGridViewFolders.Visible = true;
            //button2.Visible = true;
            //button7.Visible = true;
            //changeUsersTableVisiblity(false);
            //makeProjectsTableInVisible();
            okDir = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            changeUsersTableVisiblity(false);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (okSender)
                textBox2.Text = PublicFuncsNvars.curUser.userCode.ToString();
            changeUsersTableVisiblity(false);

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (okPro)
                textBox9.Text = dataGridViewProjects.Rows[dataGridViewProjects.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
        }

        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridViewProjects.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewProjects.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
                    break;
                }
            }
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                row1 = e.RowIndex;
                dataGridViewProjects.Rows[row1].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void button12_Click(object sender, EventArgs e)// אישור
        {
            ControlBox = false;
            int res2 = -1, res9, res14;
            int existingId = -1;
            bool docCreatedd = false;

            // WAS    if (WayIn.existing != curDocWI&&(textBox5.Text.Equals("שם תיק") || textBox5.Text.Equals("")))
            if ((textBox5.Text.Equals("שם תיק") || textBox5.Text.Equals("")))
                MessageBox.Show("יש לבחור תיק קיים לתיוק המסמך", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            //WAS  if (WayIn.existing != curDocWI&&
            else if (((!int.TryParse(textBox2.Text, out res2) &&  textBox16.Text == "") || (PublicFuncsNvars.getUserNameByUserCode(res2) == null && res2 != 99999)))
                MessageBox.Show("יש לבחור משתמש חותם. במידה ונבחר מספר משתמש לא קיים נא למחוק את המספר", "שגיאה", MessageBoxButtons.OK,
                    MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            //WAS  if (WayIn.existing != curDocWI&&
            else if ((!comboBox1.Items.Contains(comboBox1.Text)))
                MessageBox.Show("יש לבחור סיווג למסמך", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            //WAS  if (WayIn.existing != curDocWI&&
            else if ((int.TryParse(textBox9.Text, out res9) && !projects.ContainsKey(res9)))
                MessageBox.Show("אין פרויקט עם מספר מזהה זה", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            //WAS  if (WayIn.existing != curDocWI&&...
            else if ((int.TryParse(textBox14.Text, out res14) && !patterns.ContainsKey(res14)))
                MessageBox.Show("אין תבנית עם מספר מזהה זה", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
          
            else if (WayIn.existing != curDocWI && (textBox1.Text.Equals("")))
                MessageBox.Show("חובה להזין נדון למסמך", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            else if (WayIn.existing == curDocWI&&!int.TryParse(textBox17.Text, out existingId))
                MessageBox.Show("שוטף יכול להכיל רק מספרים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            else if (WayIn.existing == curDocWI&&!PublicFuncsNvars.isNormalDoc(existingId))
                MessageBox.Show("לא ניתן לשכפל שוטף שנוצר מהנחיה/בנ\"מ/השאלה/הקצאה במסך זה.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            else if(WayIn.existing==curDocWI&&!PublicFuncsNvars.docExists(existingId))
                MessageBox.Show("המספר שהוכנס אינו שוטף קיים.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            else
            {
                //makeDirectoriesTableInVisible();
                //makeProjectsTableInVisible();
                dataGridView2.Visible = false;
                button5.Visible = false;
                textBox24.Visible = false;
                label2.Visible = false;
                changeUsersTableVisiblity(false);
                //makePatternsTableInVisible();
                DialogResult result = MessageBox.Show("האם ליצור מסמך זה?", "אישור מסמך חדש", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("INSERT INTO dbo.documents(shotef_mismach, mismach_or_kovetz, hanadon, is_nichnas,"
                                            + " shotef_nichnas_yotze, tarich_hamichtav, tarich_hazana, zman_hazana, kod_sholeah, teur_tafkid_sholeah, simuchin,"
                                            + " simuchin_metzumtzam, kod_sivug_bitchoni, kod_meabed_tamlilim, msd_template, is_kayam, is_hufatz, tarich_hafatza,"
                                            + " is_pail, makor, asmachta_makor, hearot, msd_proiect, is_hasum, is_rapat, is_sodi, user_metaiek, Txt,"
                                            + " LastTxtUpdateDate, file_data, file_extension, docType)"
                                            + Environment.NewLine + "output inserted.shotef_mismach" + Environment.NewLine
                                            + "     VALUES((SELECT MAX(shotef_mismach) FROM dbo.documents WHERE shotef_mismach<90000000) + 1, 0, @subject, 0, 0, @creationDate,"
                                            + " @insertionDate, @insertionTime, @senderCode, @senderRole, @references, @shortRefs, @classificationCode, 2, 99, 1, 0, @publicationDate,"
                                            + " @isActive, '', '', @notes, @projectID, 0, 0, 0, @user, '', '', 0X0, '', 0)",
                                            conn);
                    int id = 0;
                    int userId=0;
                    short classCode= (short)Classification.confidetial;
                    short proId;
                    int directoryId = 0;
                    if (WayIn.existing != curDocWI)
                    {
                        directoryId = int.Parse(textBox12.Text);
                        comm.Parameters.AddWithValue("@subject", textBox1.Text);
                        comm.Parameters.AddWithValue("@creationDate", dateTimePicker1.Value.ToString("yyyyMMdd"));
                        comm.Parameters.AddWithValue("@insertionDate", dateTimePicker2.Value.ToString("yyyyMMdd"));
                        comm.Parameters.AddWithValue("@insertionTime", dateTimePicker2.Value.ToShortTimeString().Replace(":", ""));
                        userId = 0;
                        int.TryParse(textBox2.Text, out userId); // קוד חותם
                        comm.Parameters.AddWithValue("@senderCode", userId); // קוד חותם
                        comm.Parameters.AddWithValue("@senderRole", Ujob);// textBox3.Text); // תאור תפקיד חותם
                        comm.Parameters.AddWithValue("@references", textBox7.Text);
                        comm.Parameters.AddWithValue("@shortRefs", PublicFuncsNvars.removeNansButLetters(textBox7.Text));
                        classCode = PublicFuncsNvars.getClassificationCode(comboBox1.Text);
                        comm.Parameters.AddWithValue("@classificationCode", classCode);
                        comm.Parameters.AddWithValue("@publicationDate", "00000000");
                        comm.Parameters.AddWithValue("@isActive", true);
                        comm.Parameters.AddWithValue("@InReference", textBox23.Text);
                        comm.Parameters.AddWithValue("@notes", textBox8.Text);
                        proId = 0;
                        short.TryParse(textBox9.Text, out proId);
                        comm.Parameters.AddWithValue("@projectID", proId);
                        comm.Parameters.AddWithValue("@user", PublicFuncsNvars.curUser.userCode);
                        if (userId == PublicFuncsNvars.curUser.userCode)
                        {
                            comm.Parameters.AddWithValue("@isTransferedToSign", true);
                            comm.Parameters.AddWithValue("@dateTransferedToSign", DateTime.Today);
                        }
                        else
                        {
                            comm.Parameters.AddWithValue("@isTransferedToSign", false);
                            comm.Parameters.AddWithValue("@dateTransferedToSign", SqlDateTime.MinValue);
                        }

                        conn.Open();
                        SqlDataReader sdr = comm.ExecuteReader();
                        sdr.Read();
                        id = sdr.GetInt32(0);
                        conn.Close();
                    }

                    bool docCreated = false;
                    bool askedInOut = false;
                    string filePath = "";
                    Exception exp = null;
                    for (int j = 0; j < 5; j++)
                    {
                        string directoryPath = Program.folderPath + "\\";
                        string subject = textBox1.Text.Replace("\"", "''");
                        
                        
                        try
                        {
                            comm = new SqlCommand("UPDATE dbo.tm_doc_parm SET sn_gnrl=(sn_gnrl+1) WHERE dkey=1", conn);
                            conn.Open();
                            comm.ExecuteNonQuery();
                            conn.Close();
                            textBox7.Text += " - " + id.ToString();
                            short template = 0;
                            bool inOrOut = false, isRapat = false, fileOrDoc = false;
                            short projectId;
                            #region switch curDocWI
                            switch (curDocWI)
                            {
                                case WayIn.existing:
                                    #region WayIn.existing
                                    bool IsDocNewVersion = false;

                                    comm = new SqlCommand("SELECT file_extension, kod_sholeah, tarich_hamichtav, kod_sivug_bitchoni, hanadon FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
                                    comm.Parameters.AddWithValue("@docId", existingId);


                                    conn.Open();
                                    SqlDataReader sdr2 = comm.ExecuteReader();
                                    sdr2.Read();
                                    extension = sdr2.GetString(0);
                                    userId = sdr2.GetInt32(1);


                                    string oldSubject = sdr2.GetString(4);
                                    DateTime date = DateTime.ParseExact(sdr2.GetString(2).Trim(), "yyyyMMdd", new CultureInfo("he-IL"));
                                    Classification oldClassification = PublicFuncsNvars.getClassification(sdr2.GetInt16(3));
                                    conn.Close();
                                    comm = new SqlCommand("INSERT INTO dbo.documents(shotef_mismach, mismach_or_kovetz, hanadon, is_nichnas,"
                                        + " shotef_nichnas_yotze, tarich_hamichtav, tarich_hazana, zman_hazana, kod_sholeah, teur_tafkid_sholeah, simuchin,"
                                        + " simuchin_metzumtzam, kod_sivug_bitchoni, kod_meabed_tamlilim, msd_template, is_kayam, is_hufatz, tarich_hafatza,"
                                        + " is_pail, makor, asmachta_makor, hearot, msd_proiect, is_hasum, is_rapat, is_sodi, user_metaiek, Txt,"
                                        + " LastTxtUpdateDate, file_data, file_extension, docType, isTransferedToSign, dateTransferedToSign)"
                                        + Environment.NewLine + "output inserted.shotef_mismach, inserted.is_nichnas, inserted.simuchin,"
                                        + " inserted.kod_sivug_bitchoni, inserted.mismach_or_kovetz, inserted.hanadon, inserted.teur_tafkid_sholeah,"
                                        + " inserted.is_rapat, inserted.hearot" + Environment.NewLine
                                        + "     SELECT (SELECT MAX(shotef_mismach) FROM dbo.documents WHERE shotef_mismach<90000000) + 1, mismach_or_kovetz,"
                                        + " hanadon, is_nichnas, shotef_nichnas_yotze, @creationDate, @insertionDate, @insertionTime, @senderId,"
                                        + " @senderRole, simuchin, simuchin_metzumtzam, @classification, kod_meabed_tamlilim, msd_template,"
                                        + " is_kayam, 0, '00000000', 1, makor, asmachta_makor, @notes, @projectId, is_hasum, is_rapat,"
                                        + " is_sodi, @filingUser, Txt, LastTxtUpdateDate, file_data, file_extension, docType, @isTransferedToSign,"
                                        + " @dateTransferedToSign FROM dbo.documents WHERE shotef_mismach=@docId", conn);
                                    comm.Parameters.AddWithValue("@docId", existingId);
                                    comm.Parameters.AddWithValue("@filingUser", PublicFuncsNvars.curUser.userCode.ToString());
                                    comm.Parameters.AddWithValue("@creationDate", DateTime.Today.ToString("yyyyMMdd"));
                                    comm.Parameters.AddWithValue("@insertionDate", DateTime.Today.ToString("yyyyMMdd"));
                                    comm.Parameters.AddWithValue("@insertionTime", DateTime.Now.ToString("hhmmss"));
                                    comm.Parameters.AddWithValue("@senderId", int.Parse(textBox2.Text));
                                    comm.Parameters.AddWithValue("@senderRole", Ujob);// textBox3.Text);
                                    comm.Parameters.AddWithValue("@classification", PublicFuncsNvars.getClassificationCode(comboBox1.SelectedItem.ToString()));
                                    comm.Parameters.AddWithValue("@projectId", short.TryParse(textBox9.Text, out projectId) ? projectId : 0);
                                    comm.Parameters.AddWithValue("@notes", textBox8.Text);
                                    if (userId == PublicFuncsNvars.curUser.userCode)
                                    {
                                        comm.Parameters.AddWithValue("@isTransferedToSign", true);
                                        comm.Parameters.AddWithValue("@dateTransferedToSign", DateTime.Today);
                                    }
                                    else
                                    {
                                        comm.Parameters.AddWithValue("@isTransferedToSign", false);
                                        comm.Parameters.AddWithValue("@dateTransferedToSign", SqlDateTime.MinValue);
                                    }

                                    conn.Open();
                                    sdr2 = comm.ExecuteReader();
                                    sdr2.Read();
                                    id = sdr2.GetInt32(0);
                                    inOrOut = sdr2.GetBoolean(1);
                                    string refs = sdr2.GetString(2).Trim(), senderRole = sdr2.GetString(6).Trim(), notes = sdr2.GetString(8).Trim();
                                    classCode = sdr2.GetInt16(3);
                                    bool docOrFile = sdr2.GetBoolean(4);
                                    isRapat = sdr2.GetBoolean(7);
                                    subject = sdr2.GetString(5).Trim();
                                    conn.Close();
                                    int primeDirectory = int.Parse(textBox12.Text);
                                    SqlCommand comm2 = new SqlCommand("INSERT INTO dbo.tiukim(kod_marechet, shotef_klali, mispar_nose, is_rashi, mispar_in_tik)" +
                                        Environment.NewLine + "output inserted.mispar_in_tik" + Environment.NewLine +
                                        " VALUES(2, @id, @directory, @isPrimery, " +
                                        "(SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 1" +
                                        Environment.NewLine + "ELSE MAX(mispar_in_tik) + 1" + Environment.NewLine + "END" + Environment.NewLine +
                                        "FROM dbo.tiukim WHERE mispar_nose=@directory))", conn);
                                    comm2.Parameters.AddWithValue("@id", id);
                                    comm2.Parameters.AddWithValue("@directory", primeDirectory);
                                    comm2.Parameters.AddWithValue("@isPrimery", true);
                                    conn.Open();
                                    int numberInPrimeDirectory = (int)comm2.ExecuteScalar();
                                    conn.Close();

                                    string newRefs = primeDirectory.ToString() + " - " + numberInPrimeDirectory.ToString() + " - " + id.ToString();
                                    if (extension.ToLower() == "doc" || extension.ToLower() == "docx" && !inOrOut)
                                    {
                                        string oldFilePath = directoryPath + existingId + ".";
                                        filePath = directoryPath + id + ".";
                                        filePath += extension.ToLower();
                                        oldFilePath += extension.ToLower();

                                        SqlCommand comm3 = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                                        comm3.Parameters.AddWithValue("@id", id);
                                        conn.Open();
                                        SqlDataReader sdr = comm3.ExecuteReader();
                                        if (sdr.Read())
                                        {
                                            byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                                            if (!File.Exists(filePath))
                                                File.WriteAllBytes(oldFilePath, fileData);
                                            else
                                            {
                                                string newExistingId = existingId + "_2";
                                                oldFilePath = oldFilePath.Replace(existingId.ToString(), newExistingId);
                                                File.WriteAllBytes(oldFilePath, fileData);
                                            }
                                        }

                                        try
                                        {
                                            Wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                                        }
                                        catch
                                        {
                                            Wapp = new Word.Application();
                                        }
                                        //Wapp = new Word.Application();
                                        bool iswAppVisible = Wapp.Visible;
                                        if (iswAppVisible)
                                            Wapp.Visible = false;
                                        doc = Wapp.Documents.Add(oldFilePath);
                                        var ciNew = CultureInfo.CreateSpecificCulture("he-IL");
                                        ciNew.DateTimeFormat.Calendar = new HebrewCalendar();
                                        string oldHebrewDate = date.ToString("d", ciNew);
                                        string newHebrewDate = DateTime.Today.ToString("d", ciNew);
                                        string oldNonHebrewDate = date.ToLongDateString().Replace(date.ToString("dddd", ciNew) + " ", "");
                                        string newNonHebrewDate = DateTime.Now.ToString("dd\t\t\tMMMM yyyy");

                                        dynamic customPropertiess = doc.CustomDocumentProperties;
                                        doc.BuiltInDocumentProperties["Subject"].Value = id.ToString();
                                        int procID = GetProccessIdByWindowTitle(id.ToString());
                                        Classification newClassification = PublicFuncsNvars.getClassification(classCode);
                                        string NewClassificationString= PublicFuncsNvars.getClassificationByEnum(newClassification);
                                        
                                        try
                                        {
                                            dynamic existingProperty = customPropertiess["נמענים_לפעולה"];
                                            IsDocNewVersion = true;
                                        }
                                        catch { }
                                        if (IsDocNewVersion)
                                        {
                                            customPropertiess["סימוכין"].Value = newRefs;
                                            try
                                            {
                                                customPropertiess["תאריך_עברי"].Value = newHebrewDate;
                                                customPropertiess["תאריך_לועזי"].Value = newNonHebrewDate;
                                                doc.BuiltInDocumentProperties["Category"].Value = NewClassificationString;
                                                if (userId != int.Parse(textBox2.Text))
                                                {
                                                    conn.Close();
                                                    comm = new SqlCommand("SELECT hatimh FROM dbo.tmtafkidu WHERE kod_tpkid=@userId", conn);
                                                    comm.Parameters.AddWithValue("@userId", int.Parse(textBox2.Text));
                                                    conn.Open();
                                                    sdr = comm.ExecuteReader();
                                                    string[] linesNew = { "", "", "" };
                                                    if (sdr.Read())
                                                    {
                                                        string[] existingLines = sdr.GetString(0).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                                        int max = 0;
                                                        if (existingLines.Length > 3) max = 3; else max = existingLines.Length;
                                                        for (int i = 0; i < max; i++)
                                                            linesNew[i] = existingLines[i];
                                                    }
                                                    customPropertiess["חתימה_שורה_א"].Value = linesNew[0];
                                                    customPropertiess["חתימה_שורה_ב"].Value = linesNew[1];
                                                    customPropertiess["חתימה_שורה_ג"].Value = linesNew[2];
                                                    //conn.Close();
                                                }
                                            }
                                            catch { }
                                        }
                                        else
                                        {

                                            //PublicFuncsNvars.replaceInWordDoc(Wapp, refs, newRefs);
                                            //PublicFuncsNvars.replaceInWordDoc(Wapp, oldHebrewDate, newHebrewDate);
                                            //PublicFuncsNvars.replaceInWordDoc(Wapp, oldNonHebrewDate, newNonHebrewDate);
                                            //PublicFuncsNvars.replaceTextInHeaderFooter(doc, existingId.ToString(), id.ToString());

                                            //PublicFuncsNvars.replaceTextInHeaderFooter(doc, PublicFuncsNvars.getClassificationByEnum(oldClassification), PublicFuncsNvars.getClassificationByEnum(newClassification));
                                            Word.Bookmark bookmark1 = null;
                                            Word.Bookmark bookmark2 = null;
                                            Word.Range range = null;
                                            try {
                                                 bookmark1 = doc.Bookmarks["לידיעה"];
                                                 bookmark2 = doc.Bookmarks["חתימה"];
                                                 range = doc.Range(bookmark1.Range.End, bookmark2.Range.Start);
                                                 range.Copy();
                                                topaste = true;
                                            }

                                            catch
                                            {
                                                MessageBox.Show("לא נמצאו סימניות ולכן לא ניתן לשכפל את תוכן מסמך המקור", "בעיה בשכפול מסמך קיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                                               MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                            }
                                            doc.SaveAs2(filePath);
                                            doc.Close();
                                            File.Delete(filePath);
                                            template = (short)(int.Parse(textBox14.Text));
                                            SqlConnection conn3 = new SqlConnection(Global.ConStr);
                                            SqlCommand comm4 = new SqlCommand("SELECT file_data,nam_word_template FROM dbo.tm_templ_bhi WHERE onum_word_template=@pat", conn3);
                                            comm4.Parameters.AddWithValue("@pat", template);
                                            conn3.Open();
                                            SqlDataReader sdr4 = comm4.ExecuteReader();
                                            sdr4.Read();
                                            string templatePathE = sdr4.GetString(1);
                                            //templatePath = templatePath.Split('.')[0];
                                            //templatePath = templatePath + "_3.dotx";
                                            templatePathE = Global.P_APP + templatePathE;
                                            conn3.Close();

                                            doc = Wapp.Documents.Add(templatePathE);
                                            customPropertiess = doc.CustomDocumentProperties;
                                            customPropertiess["סימוכין"].Value = newRefs;
                                            customPropertiess["תאריך_עברי"].Value = newHebrewDate;
                                            customPropertiess["תאריך_לועזי"].Value = newNonHebrewDate;
                                            doc.BuiltInDocumentProperties["Category"].Value = NewClassificationString;
                                            doc.BuiltInDocumentProperties["Subject"].Value = id.ToString();
                                            conn.Close();

                                            comm = new SqlCommand("SELECT hatimh FROM dbo.tmtafkidu WHERE kod_tpkid=@userId", conn);
                                            comm.Parameters.AddWithValue("@userId", int.Parse(textBox2.Text));
                                            conn.Open();
                                            sdr = comm.ExecuteReader();
                                            string[] linesNew = { "", "", "" };
                                            if (sdr.Read())
                                            {
                                                string[] existingLines = sdr.GetString(0).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                                int max = 0;
                                                if (existingLines.Length > 3) max = 3; else max = existingLines.Length;
                                                for (int i = 0; i < max; i++)
                                                    linesNew[i] = existingLines[i];
                                            }
                                            
                                            customPropertiess["חתימה_שורה_א"].Value = linesNew[0];
                                            customPropertiess["חתימה_שורה_ב"].Value = linesNew[1];
                                            customPropertiess["חתימה_שורה_ג"].Value = linesNew[2];

                                            conn.Close();

                                            comm = new SqlCommand("SELECT hanadon FROM dbo.documents (nolock)"
                                                + " WHERE shotef_mismach=@id ", conn);
                                            comm.Parameters.AddWithValue("@id", existingId);
                                            conn.Open();
                                            sdr = comm.ExecuteReader();
                                            if (sdr.Read())
                                            {
                                                //string hanadon = sdr.GetString(0).Trim() + "_2";
                                                doc.BuiltInDocumentProperties["Title"].Value = sdr.GetString(0).Trim();
                                            }
                                            conn.Close();

                                            comm = new SqlCommand("SELECT row_0, row_1, row_2, row_3, row_4, row_5, row_6, row_7, row_8, row_9 FROM dbo.tmtafkidu t join dbo.tm_kubiot k on k.cod_kobyh=t.kod_kobih WHERE t.kod_tpkid=@userId", conn);
                                            comm.Parameters.AddWithValue("@userId", res2);
                                            conn.Open();
                                            sdr = comm.ExecuteReader();
                                            if (sdr.Read())
                                            {
                                                customPropertiess["קוביה_א"].Value = sdr.GetString(0).Trim();
                                                customPropertiess["קוביה_ב"].Value = sdr.GetString(1).Trim();
                                                customPropertiess["קוביה_ג"].Value = sdr.GetString(2).Trim();
                                                customPropertiess["קוביה_ד"].Value = sdr.GetString(3).Trim();
                                                customPropertiess["קוביה_ה"].Value = sdr.GetString(4).Trim();
                                                customPropertiess["קוביה_ו"].Value = sdr.GetString(5).Trim();
                                                customPropertiess["קוביה_ז"].Value = sdr.GetString(6).Trim();
                                                customPropertiess["קוביה_ח"].Value = sdr.GetString(7).Trim();
                                                customPropertiess["קוביה_ט"].Value = sdr.GetString(8).Trim();
                                                customPropertiess["קוביה_י"].Value = sdr.GetString(9).Trim();

                                            }
                                            conn.Close();
                                            doc.Fields.Update();
                                            Word.Range docRange = doc.Range();
                                            if (docRange.Tables.Count > 0)
                                            {
                                                foreach (Word.Row row in docRange.Tables[1].Rows)
                                                {
                                                    bool del = true;
                                                    foreach (Word.Cell cell in row.Cells)
                                                    {
                                                        if (!cell.Range.Text.Equals("\v\r\a"))
                                                        {
                                                            del = false;
                                                            break;
                                                        }
                                                    }
                                                    if (del)
                                                        row.Delete();
                                                }
                                            } 
                                        }

                

                                        doc.Fields.Update();
                                        if (topaste)
                                        {

                                            foreach (Word.Paragraph paragraph in doc.Paragraphs)
                                            {
                                                if (paragraph.Range.ListFormat.ListType == Word.WdListType.wdListSimpleNumbering)
                                                {
                                                    Word.Range afterRange = doc.Range(paragraph.Range.End, paragraph.Range.End);
                                                    afterRange.Paste();
                                                    topaste = false;
                                                    break;
                                                }
                                            }
                                            //foreach (Word.Field field in doc.Fields)
                                            //{
                                            //    if (field.Code.Text.Contains("Title") || field.Code.Text.Contains("כותרת"))
                                            //    {
                                            //        Word.Range fieldRange = field.Result;
                                            //        Word.Range beforefieldRange = doc.Range(fieldRange.End, fieldRange.End);
                                            //        beforefieldRange.Paste();
                                            //        topaste = false;
                                            //        break;
                                            //    }
                                            //}
                                        }
                                       // doc.Fields.Update();
                                        doc.SaveAs2(filePath);
                                        string Text = PublicFuncsNvars.docToTxt(doc, filePath);
                                        doc.Close();
                                        if (iswAppVisible)
                                            Wapp.Visible = true;
                                        //Wapp.Quit();
                                        saveOriginalFileBlob(filePath, id, extension);
                                        byte[] newFileData = File.ReadAllBytes(filePath);
                                        PublicFuncsNvars.saveDocToDB(ref newFileData, id, filePath, ref comm, ref conn , Text);//2
                                        File.Delete(filePath);
                                        File.Delete(oldFilePath);

                                        //if (oldClassification != newClassification)

                                        //PublicFuncsNvars.updateSubjectAndClassificationInWordDoc(id, subject, subject, oldClassification, newClassification);
                                    }
                                    comm = new SqlCommand("UPDATE dbo.documents SET simuchin=@refferences, simuchin_metzumtzam=@shortRefs"
                                        + " WHERE shotef_mismach=@id", conn);
                                    comm.Parameters.AddWithValue("@id", id);
                                    comm.Parameters.AddWithValue("@refferences", newRefs);
                                    comm.Parameters.AddWithValue("@shortRefs", PublicFuncsNvars.removeNansButLetters(newRefs));
                                    conn.Open();
                                    comm.ExecuteNonQuery();
                                    conn.Close();
                                    comm = new SqlCommand("INSERT INTO dbo.doc_mech(kod_marechet, shotef_klali, msd, kod_mechutav, tiur_tafkid, is_lepeula,"
                                        + " is_ishu_kabala, is_lishloh_mail, ktovet_mail) SELECT kod_marechet, @docId, msd, kod_mechutav, tiur_tafkid, is_lepeula,"
                                        + " is_ishu_kabala, is_lishloh_mail, ktovet_mail from dbo.doc_mech WHERE shotef_klali=@id", conn);
                                    comm.Parameters.AddWithValue("@id", existingId);
                                    comm.Parameters.AddWithValue("@docId", id);
                                    conn.Open();
                                    comm.ExecuteNonQuery();
                                    conn.Close();
                                    comm = new SqlCommand("INSERT INTO dbo.doc_Authorizations(docId, roleCode, isForEdit) SELECT @id, roleCode,"
                                        + " isForEdit FROM dbo.doc_Authorizations WHERE docId=@existingId", conn);
                                    comm.Parameters.AddWithValue("@existingId", existingId);
                                    comm.Parameters.AddWithValue("@id", id);
                                    conn.Open();
                                    comm.ExecuteNonQuery();
                                    conn.Close();
                                    curDoc = new Document(id, docOrFile, subject, inOrOut, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"),
                                        "00000000", int.Parse(textBox2.Text), senderRole, newRefs, PublicFuncsNvars.getClassification(classCode), false, true, isRapat, null,
                                        PublicFuncsNvars.getDirectoriesForDoc(id), null, notes, DocType.normal, false, DateTime.MinValue);
                                    docCreatedd = true;
                                    if (!IsDocNewVersion)
                                        MessageBox.Show("המסמך בתבנית ישנה." + Environment.NewLine + "נתוני המסמך שוכפלו." + Environment.NewLine + "יש לערוך ידנית את המסמך", "מסמך ישן", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                    break;
                                    #endregion
                                case WayIn.word:
                                    #region WayIn.word
                                    if (textBox15.Text.Equals("") || textBox15.Text.Equals("שם תבנית"))
                                    {
                                        MessageBox.Show("עליך לבחור תבנית למסמך", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error,
                                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                        Cursor.Current = Cursors.Default;
                                        return;
                                    }
                                    extension = "docx";
                                    template = (short)(int.Parse(textBox14.Text));
                                    SqlConnection conn1 = new SqlConnection(Global.ConStr);
                                    SqlCommand comm1 = new SqlCommand("SELECT file_data,nam_word_template FROM dbo.tm_templ_bhi WHERE onum_word_template=@pat", conn1);
                                    comm1.Parameters.AddWithValue("@pat", template);
                                    conn1.Open();
                                    SqlDataReader sdr1 = comm1.ExecuteReader();
                                    sdr1.Read();
                                    string templatePath = sdr1.GetString(1);
                                    //templatePath = templatePath.Split('.')[0];
                                    //templatePath = templatePath + "_3.dotx";
                                    templatePath = Global.P_APP + templatePath;
                                    conn1.Close();
                                    filePath = directoryPath + id + ".";
                                    filePath += extension;
                                    try
                                    {
                                        Wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                                    }
                                    catch
                                    {
                                        Wapp = new Word.Application();
                                    }
                                    bool isWAppVisible = Wapp.Visible;
                                    if (isWAppVisible)
                                        Wapp.Visible = false;
                                    //Wapp = new ();
                                    doc = Wapp.Documents.Add(templatePath);
                                    dynamic customProperties = doc.CustomDocumentProperties;
                                    
                                    comm1 = new SqlCommand("SELECT row_0, row_1, row_2, row_3, row_4, row_5, row_6, row_7, row_8, row_9 FROM dbo.tmtafkidu t join dbo.tm_kubiot k on k.cod_kobyh=t.kod_kobih WHERE t.kod_tpkid=@userId", conn1);
                                    comm1.Parameters.AddWithValue("@userId", res2);
                                    conn1.Open();
                                    sdr1 = comm1.ExecuteReader();
                                    if (sdr1.Read())
                                    {
                                        customProperties["קוביה_א"].Value = sdr1.GetString(0).Trim(); 
                                        customProperties["קוביה_ב"].Value = sdr1.GetString(1).Trim();
                                        customProperties["קוביה_ג"].Value = sdr1.GetString(2).Trim();
                                        customProperties["קוביה_ד"].Value = sdr1.GetString(3).Trim();
                                        customProperties["קוביה_ה"].Value = sdr1.GetString(4).Trim();
                                        customProperties["קוביה_ו"].Value = sdr1.GetString(5).Trim(); 
                                        customProperties["קוביה_ז"].Value = sdr1.GetString(6).Trim();
                                        customProperties["קוביה_ח"].Value = sdr1.GetString(7).Trim();
                                        customProperties["קוביה_ט"].Value = sdr1.GetString(8).Trim();
                                        customProperties["קוביה_י"].Value = sdr1.GetString(9).Trim();
                                       
                                    }
                   

                                    doc.Fields.Update();
                                    conn1.Close();
                                    customProperties["סימוכין"].Value = textBox7.Text;
                                    var ci = CultureInfo.CreateSpecificCulture("he-IL");
                                    ci.DateTimeFormat.Calendar = new HebrewCalendar();
                                    var ciNeww = CultureInfo.CreateSpecificCulture("he-IL");
                                    ciNeww.DateTimeFormat.Calendar = new HebrewCalendar();
                                    customProperties["תאריך_עברי"].Value = dateTimePicker1.Value.ToString("d", ci);
                                    customProperties["תאריך_לועזי"].Value = DateTime.Now.ToString("dd\t\t\tMMMM yyyy");

                                    customProperties["תאריך"].Value = dateTimePicker1.Value.ToShortDateString();
                                    //customProperties["נדון"].Value = textBox1.Text;
                                    doc.BuiltInDocumentProperties["Title"].Value = textBox1.Text;//Ahava 10.01.2024 update the custom property.//שיניתי
                                    doc.BuiltInDocumentProperties["Subject"].Value = id.ToString();//Ahava 10.01.2024 update the custom property.
                                    customProperties["תוכן"].Value = "";
                                    Word.Range rng = doc.Range(); 
                                    if (rng.Tables.Count > 0)            
                                        foreach (Word.Row row in rng.Tables[1].Rows)
                                        {
                                            bool del = true;
                                            foreach (Word.Cell cell in row.Cells)
                                            {
                                                if (!cell.Range.Text.Equals("\v\r\a"))
                                                {
                                                    del = false;
                                                    break;
                                                }
                                            }
                                            if (del)
                                                row.Delete();
                                        }
                                    comm1 = new SqlCommand("SELECT hatimh FROM dbo.tmtafkidu WHERE kod_tpkid=@userId", conn1);
                                    comm1.Parameters.AddWithValue("@userId", res2);
                                    conn1.Open();
                                    sdr1 = comm1.ExecuteReader();
                                    if (sdr1.Read())
                                    {
                                        string[] lines = { "", "", "" };
                                        string[] existingLines = sdr1.GetString(0).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                                        int max = 0;
                                        if (existingLines.Length > 3) max = 3; else max = existingLines.Length;
                                        for (int i = 0; i < max; i++)
                                            lines[i] = existingLines[i];
                                        customProperties["חתימה_שורה_א"].Value =lines[0];
                                        customProperties["חתימה_שורה_ב"].Value = lines[1];
                                        customProperties["חתימה_שורה_ג"].Value = lines[2];
                                    }
                                    conn1.Close();
                                    
                                    doc.BuiltInDocumentProperties["Category"].Value = comboBox1.Text;//Ahava 10.01.2024 update the custom property.
                                    bool docOpened = false; 
                                    exp = null;
                                    doc.Fields.Update();
                                    doc.SaveAs2(filePath);
                                    docOpened = true;
                                    doc.Close();
                                    if (isWAppVisible)
                                        Wapp.Visible = true;
                                    if (!docOpened)        
                                    {
                                        PublicFuncsNvars.saveLogError(FindForm().Name, exp.ToString(), exp.Message);
                                        comm = new SqlCommand("DELETE FROM dbo.documents WHERE shotef_mismach=@id", conn);
                                        comm.Parameters.AddWithValue("@id", id);
                                        if (conn.State == ConnectionState.Closed)
                                            conn.Open();
                                        comm.ExecuteNonQuery();
                                        conn.Close();
                                        MessageBox.Show("יצירת מסמך נכשלה, אנא נסו שוב.");
                                        ControlBox = true;
                                        textBox7.Text = textBox7.Text.Remove(textBox7.Text.Length - id.ToString().Length - 3);
                                        this.Visible = true;
                                        this.BringToFront();
                                        Cursor.Current = Cursors.Default;
                                        return;
                                    }

                                    Marshal.ReleaseComObject(doc);
                                    //Wapp.Quit();
                                    //Marshal.ReleaseComObject(Wapp);
                                    saveOriginalFileBlob(filePath, id, extension);
                                    File.Delete(filePath);
                                    this.Visible = false;
                                    docCreatedd = true;
                                    break;
                                    #endregion
                                case WayIn.excel:
                                    #region WayIn.excel
                                    extension = "xlsx";
                                    Eapp = new Excel.Application();
                                    Excel.Workbook wb = Eapp.Workbooks.Add();
                                    wb.SaveAs(filePath + extension, XlFileFormat.xlWorkbookDefault);
                                    wb.Close();
                                    this.Enabled = false;
                                    Marshal.ReleaseComObject(wb);
                                    Eapp.Quit();
                                    Marshal.ReleaseComObject(Eapp);
                                    saveOriginalFileBlob(filePath, id, extension);
                                    docCreatedd = true;
                                    File.Delete(filePath);
                                    PublicFuncsNvars.viewDocForEdit(id);
                                    break;
                                    #endregion
                                case WayIn.scan:
                                case WayIn.file:
                                    fileOrDoc = true;
                                    File.Copy(path, filePath + extension);
                                    if (!askedInOut)
                                    {
                                        Cursor.Current = Cursors.Default;
                                        inOrOut = true;
                                        DialogResult result3 = MessageBox.Show("האם זה מסמך מרפ\"ט?", "מסמך רפ\"ט", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                               MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                        if (result3 == DialogResult.Yes)
                                        {
                                            isRapat = true;
                                        }

                                        Cursor.Current = Cursors.WaitCursor;
                                    }
                                    filePath += extension;
                                    saveOriginalFileBlob(filePath, id, extension);
                                    docCreatedd = true;
                                    File.Delete(filePath);
                                    while (File.Exists(filePath)) ;
                                    break;

                                case WayIn.drag:
                                    fileOrDoc = true;
                                    inOrOut = true;
                                    saveOriginalFileBlob(mailPath, id, extension);
                                    if (extension == "msg"  && mailPath.ToLower().Contains("temp")) File.Delete(mailPath);
                                    docCreatedd = true;
                                    break;
                            }
                            #endregion
                            if (curDocWI != WayIn.existing)
                            {
                                comm = new SqlCommand("UPDATE dbo.documents SET mismach_or_kovetz=@fileOrDoc, is_nichnas=@isIncoming,"
                                    + " msd_template=@docPattern, is_rapat=@isRapat, simuchin=@refferences, simuchin_metzumtzam=@shortRefs WHERE shotef_mismach=@id",
                                    conn);
                                comm.Parameters.AddWithValue("@id", id);
                                comm.Parameters.AddWithValue("@fileOrDoc", fileOrDoc);
                                comm.Parameters.AddWithValue("@isIncoming", inOrOut);
                                comm.Parameters.AddWithValue("@docPattern", template);
                                comm.Parameters.AddWithValue("@isRapat", isRapat);
                                comm.Parameters.AddWithValue("@refferences", textBox7.Text);
                                comm.Parameters.AddWithValue("@shortRefs", PublicFuncsNvars.removeNansButLetters(textBox7.Text));
                                conn.Open();
                                comm.ExecuteNonQuery();
                                conn.Close();
                                comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                                    "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@directory",
                                    conn);
                                comm.Parameters.AddWithValue("@directory", int.Parse(textBox12.Text));
                                conn.Open();
                                int nif = (int)comm.ExecuteScalar() + 1;
                                conn.Close();
                                comm.CommandText  = "INSERT INTO dbo.tiukim(kod_marechet, shotef_klali, mispar_nose, is_rashi, mispar_in_tik) VALUES(2, @id, @directory, 1, @nif)";
                                comm.Parameters.AddWithValue("@id", id);
                                comm.Parameters.AddWithValue("@nif", nif);
                                conn.Open();
                                comm.ExecuteNonQuery();
                                conn.Close();

                                curDoc = new Document(id, fileOrDoc, textBox1.Text, inOrOut, dateTimePicker1.Value.ToString("yyyyMMdd"),
                                    dateTimePicker2.Value.ToString("yyyyMMdd"), "00000000", int.Parse(textBox2.Text), Ujob/*textBox3.Text*/, textBox7.Text,
                                    PublicFuncsNvars.getClassification(classCode), false, true, isRapat, null, PublicFuncsNvars.getDirectoriesForDoc(id), null,
                                    textBox8.Text, DocType.normal, false, DateTime.MinValue);

                                User u = PublicFuncsNvars.getUserByCode(userId);
                                if (u != null)
                                {
                                    Dictionary<int, bool> authorizedUsers = PublicFuncsNvars.getUserByCode(userId).getAutoAuthorizedUsers();
                                    foreach (KeyValuePair<int, bool> au in authorizedUsers)
                                        curDoc.addAuthorization(au.Key, au.Value);
                                    if (PublicFuncsNvars.getClassification(classCode) == Classification.sensitivePersonal)
                                        curDoc.addAuthorization(PublicFuncsNvars.curUser.userCode, true);
                                }
                            }
                            makeDataControlsEnDisabled(false);
                            changeSelectedFileControlsVisibility(false);

                            button13.Visible = true;// עבור לטיפול במסמך שנוצר
                            button14.Visible = true;// צור מסמך חדש
                            button13.BringToFront();
                            button14.BringToFront();
                            button12.Visible = false;// אישור
                            docCreated = true; 
                            customLabel2.Visible = false;// מספר שוטף
                            textBox17.Visible = false;// מספר שוטף
                            button17.Visible = false;// חפש (שוטף)
                            break;
                        }
                        catch (Exception ex)
                        {
                            exp = ex;
                            MessageBox.Show(ex.Message + "\n" + ex.StackTrace, ex.GetType().ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                            textBox7.Text = textBox7.Text.Remove(textBox7.Text.Length - id.ToString().Length - 3);
                            Thread.Sleep(100);
                        }
                    }
                   
                    if (!docCreated)
                    {
                        success = false;
                        PublicFuncsNvars.saveLogError(FindForm().Name, exp.ToString(), exp.Message);
                        comm = new SqlCommand("DELETE FROM dbo.documents WHERE shotef_mismach=@id", conn);
                        comm.Parameters.AddWithValue("@id", id);
                        if (conn.State == ConnectionState.Closed)
                            conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        comm = new SqlCommand("DELETE FROM dbo.tiukim WHERE kod_marechet=2 AND shotef_klali=@id", conn);
                        comm.Parameters.AddWithValue("@id", id);
                        if (conn.State == ConnectionState.Closed)
                            conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        MessageBox.Show("יצירת מסמך נכשלה, נשלח אליכם דוא\"ל בנושא.");
                        List<Tuple<byte[], string>> atts = new List<Tuple<byte[], string>>();
                        if (fileData.Length > 0)
                            atts.Add(new Tuple<byte[], string>(fileData, filePath.Substring(filePath.LastIndexOf('\\') + 1)));
                        PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - הודעות מחשוב", PublicFuncsNvars.curUser.email + ";mntkmihshuv@modnet.il", null, null,
                            "יצירת שוטף נכשלה", "השוטף שניסית ליצור לא הצליח להישמר בבסיס הנתונים." + Environment.NewLine + "מצורף המסמך עליו עבדת.",
                            atts);
                    }
                    else
                    {
                        success = true;
                    }
                }
            }
            ControlBox = true;
            this.Cursor = Cursors.Default;
            this.Enabled = true;
            this.TopMost = true;
            this.TopMost = false;
            this.Activate();
            this.BringToFront();
            Cursor.Current = Cursors.Default;
            if (success)
            {
                if (docCreatedd)
                    button13_Click(sender, e);
                


                
                else
                {
                   

                    MyGlobals.dragFlag = true;
                    Thread newDocThread = new Thread(openNewDocForm);
                   newDocThread.SetApartmentState(ApartmentState.STA);
                   newDocThread.Start();
                    this.Close();
                    
                }
            }
        }

        private int GetProccessIdByWindowTitle(string appID)
        {
            Process[] P_CESSES = Process.GetProcesses();
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Contains(appID) && !P_CESSES[p_count].MainWindowTitle.Contains("טיפול במסמך שוטף"))
                {
                    return P_CESSES[p_count].Id;
                }
            }

            return int.MaxValue;
        }

        private void saveOriginalFileBlob(string filePath, int id, string extension)
        {
            FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            BinaryReader reader = new BinaryReader(fs);
            byte[] fileData = reader.ReadBytes((int)fs.Length);
            reader.Close();
            fs.Close();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET file_data=@fileData, file_extension=@fileExt WHERE shotef_mismach=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@fileData", fileData);
            comm.Parameters.AddWithValue("@fileExt", extension);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("שוטף " + id + " נוצר בהצלחה");
        }

        private void wb_BeforeClose(ref bool Cancel)
        {
            isExcelClosed = !Cancel;
        }

        private void makeDataControlsEnDisabled(bool b)
        {
            Control[] controls = { textBox1,textBox2,textBox3, textBox5,textBox6,textBox7,textBox8,textBox9,textBox10,textBox11,textBox12,dateTimePicker1,
                                     dateTimePicker2,comboBox1,/*button4,*/button12,button1,button2,button3, button15, textBox23, textBox20};
            PublicFuncsNvars.makeControlsEnDisabled(b, controls);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            makeDirectoriesTableInVisible();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox12.Text = "קוד";
            makeDirectoriesTableInVisible();
        }

        private void button10_Click(object sender, EventArgs e)// סיום (טבלת תיקים)
        {
            makeProjectsTableInVisible();
        }

        private void button11_Click(object sender, EventArgs e)// ביטול (טבלת תיקים)
        {
            textBox9.Text = "קוד";
            makeProjectsTableInVisible();
        }

        private void NewDocument_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)// עבור לטיפול במסמך שנוצר
        {
            Thread docHandleThread = new Thread(openDocumentHandlingForm);
            docHandleThread.SetApartmentState(ApartmentState.STA);
            docHandleThread.Start(curDoc.getID());

            /*DocumentHandling dh = new DocumentHandling(curDoc.getID());
            dh.Activate();
            dh.ShowDialog();*/
            this.Close();
        }
        private void openDocumentHandlingForm(object obj)
        {
            int d = (int)obj;
            dh = new DocumentHandling(d);
            dh.Activate();
            dh.ShowDialog();
        }
        private void button14_Click(object sender, EventArgs e)
        {       // צור מסמך חדש
            changeDataControlsVisiblity(false);
            changeInputControlsVisiblity(false);
            label13.Visible = false;// תבנית מסמך
            textBox14.Visible = false;// קוד תבנית
            textBox15.Visible = false;// שם תבנית
            dataGridViewProjects.Visible = false;
            dataGridViewFolders.Visible = false;
            dataGridViewUsers.Visible = false;
            button13.Visible = false;// עבור לטיפול במסמך שנוצר
            button14.Visible = false;// צור מסמך חדש
            
            makeDataControlsEnDisabled(true);
            clearContents();
        }

        private void clearContents()
        {
            textBox1.Clear();
            textBox2.Text = PublicFuncsNvars.curUser.userCode.ToString();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            dateTimePicker1.ResetText();
            dateTimePicker2.ResetText();
            comboBox1.SelectedIndex = -1;
            textBox13.Clear();
            textBox14.Clear();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            int res;
            if (int.TryParse(textBox14.Text, out res))
            {
                try
                {
                    textBox15.Text = patterns[res];
                }
                catch (Exception ex)
                {
                    textBox14.Text = "";
                }
            }
            else
                textBox15.Text = "שם תבנית";
        }

        private void makePatternsTableInVisible()
        {
            Control[] controls = { label14, dataGridViewPatterns, button19, button20 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void textBox14_Click(object sender, EventArgs e)
        {
            //makeDirectoriesTableInVisible();
            //makeProjectsTableInVisible();
            //changeUsersTableVisiblity(false);
            //label14.Visible = true;
            ChangeDataGrid("patterns");
            /*while (true)
            {
                try
                {
                    dataGridViewPatterns.SelectionChanged -= dataGridView5_SelectionChanged;
                    dataGridViewPatterns.Visible = true;
                    dataGridViewPatterns.SelectionChanged += dataGridView5_SelectionChanged;
                    break;
                }
                catch { }
            }*/
            //button19.Visible = true;
            //button20.Visible = true;
            okPat = true;
        }

        private void textBox14_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox14.Text, out res))
            {
                textBox14.Text = "";
            }
            okPat = true;
        }

        private void textBox14_Leave(object sender, EventArgs e)
        {
            if (textBox14.Text.Equals(""))
                textBox14.Text = "קוד";
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            if (okPat)
                textBox14.Text = dataGridViewPatterns.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }

        private void dataGridView5_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridViewPatterns.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewPatterns.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
                    break;
                }
            }
        }

        private void dataGridView5_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView5_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                row5 = e.RowIndex;
                dataGridViewPatterns.Rows[row5].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            textBox14.Text = "קוד";
            makePatternsTableInVisible();
            okPat = false;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            makePatternsTableInVisible();
            okPat = false;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value.Date;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker1.MaxDate = dateTimePicker2.Value;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
            {
                textBox2.Text = "";
               // textBox4.Text = "";
                //textBox16.Text = "";

            }
            else if (textBox3.Text.Equals("שם / תפקיד"))
            {
                textBox2.Text = "קוד";
               // textBox4.Text = "שם פרטי";
               // textBox16.Text = "שם משפחה";
            }
            else if(textBox2.Text!="99999")
            {
                //List<User> FilteredUsers = users.Where(word => word.job.ToString().StartsWith(textBox3.Text)).ToList();
                //dataGridView2.Columns[0].DataPropertyName = "id";
                //dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                //dataGridView2.Columns[2].DataPropertyName = "description";
                /*dataGridView2.DataSource = FilteredUsers.Select(item => new
                {//userCode, u.firstName, u.lastName, u.job
                    item.userCode,
                    item.firstName,
                    item.lastName,
                    item.job
                }).ToList();*/
                /*foreach (User u in users)
                    dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
                dataGridView2.Refresh();
                int numOfRows = dataGridView2.Rows.Cast<DataGridViewRow>().Count(row => row.Visible);
                //int index = 0;
                if (numOfRows > 0)
                {
                    dataGridView2.FirstDisplayedScrollingRowIndex = 0;
                }
                else
                {
                    textBox12.Text = "";
                    dataGridView2.DataSource = users.Select(item => new
                    {
                        item.userCode,
                        item.firstName,
                        item.lastName,
                        item.job
                    }).ToList();
                    dataGridView2.Refresh();
                }*/
                //foreach (User u in users)
                //dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
                //dataGridView2.Sort(dataGridView2.Columns[3], ListSortDirection.Ascending);
                int index = 0;
                foreach (DataGridViewRow row in dataGridViewUsers.Rows)
                {
                    if (row.Cells[3].Value != null && row.Cells[3].Value.ToString().StartsWith(textBox3.Text))
                    {
                        index = row.Index;
                        dataGridView2.FirstDisplayedScrollingRowIndex = index;
                        break;
                    }
                }
                
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text.Equals(""))
            {
                textBox2.Text = "";
                textBox3.Text = "";
                textBox16.Text = "";

            }
            else if (textBox4.Text.Equals("שם פרטי"))
            {
                textBox2.Text = "קוד";
                textBox3.Text = "תפקיד";
                textBox16.Text = "שם משפחה";
            }
            else if(textBox2.Text!="99999")
            {

                dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);
                int index = 0;
                foreach (DataGridViewRow row in dataGridViewUsers.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString().StartsWith(textBox4.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            if (textBox16.Text.Equals(""))
            {
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";

            }
            else if (textBox16.Text.Equals("שם"))
            {
                textBox2.Text = "קוד";
                textBox3.Text = "תפקיד";
                textBox4.Text = "שם פרטי";
            }
            else if(textBox2.Text!="99999" && textBox2.Text != "קוד")
            {
                dataGridView2.Sort(dataGridView2.Columns[2], ListSortDirection.Ascending);
                int index = 0;
                foreach (DataGridViewRow row in dataGridViewUsers.Rows)
                {
                    if (row.Cells[2].Value != null && row.Cells[2].Value.ToString().StartsWith(textBox16.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox16_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox16_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox16.Text.Equals("שם משפחה") && textBox2.Text != "99999")
            {
                strTyped = "";
                textBox16.Text = "";
            }
        }

        private void textBox16_Leave(object sender, EventArgs e)
        {
            if (textBox16.Text.Equals(""))
                textBox16.Text = "שם משפחה";
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox3.Text=="תפקיד" && textBox2.Text!="99999")
            {
                strTyped = "";
                textBox3.Text = "";
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
                textBox3.Text = "תפקיד";
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox4.Text.Equals("שם פרטי") && textBox2.Text != "99999")
            {
                strTyped = "";
                textBox4.Text = "";
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            if (textBox4.Text.Equals(""))
                textBox4.Text = "שם פרטי";
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }



 

        private void panel_dropDown_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void panel_dropDown_DragDrop(object sender, DragEventArgs e)
        {
            if (app == null)
                app = new Microsoft.Office.Interop.Outlook.Application();
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {

                Microsoft.Office.Interop.Outlook.Explorer oExplorer = app.ActiveExplorer();
                Microsoft.Office.Interop.Outlook.Selection oSelection = oExplorer.Selection;
                if (oSelection.Count != 1) return;
                Microsoft.Office.Interop.Outlook.MailItem mi = null;
                try {
                     mi = (Microsoft.Office.Interop.Outlook.MailItem)oSelection[1];
                }

                catch { return; }
                string emailAddress = "";
                if (mi.SenderEmailType == "SMTP")
                    emailAddress = mi.SenderEmailAddress.ToLower().Trim();
                else
                    emailAddress = mi.Sender.GetExchangeUser().PrimarySmtpAddress.ToLower().Trim();

                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm1 = new SqlCommand("SELECT kod_tpkid FROM dbo.tmtafkidu WHERE isActive = 1 AND LOWER(doal) LIKE '%' + @userMail + '%'", conn);
                comm1.Parameters.AddWithValue("@userMail", emailAddress);
                conn.Open();
                SqlDataReader sdr = comm1.ExecuteReader();
                
                while (sdr.Read())
                {
                    textBox2.Text = sdr.GetInt32(0).ToString();
                }

                conn.Close();
                dateTimePicker1.Value = mi.CreationTime.Date;
                
                string s = mi.Subject.Replace("@שמור@", "").Replace("@בלמ\"ס@", "").Replace("@סודי@", "").Replace("@סודי ביותר@", "");
                textBox1.Text = s;
                string body = Regex.Replace(mi.Body, @"[\r\n]+", "\r\n");
                textBox8.Text = body;
                // אתחול סיווג מסמך לפי סיווג בכותרת המייל
                if (mi.Subject.Contains("@בלמ\"ס")) comboBox1.SelectedIndex = 0;
                else if (mi.Subject.Contains("@שמור")) comboBox1.SelectedIndex = 1;
                else if (mi.Subject.Contains("@סודי@")) comboBox1.SelectedIndex = 2;
                else if (mi.Subject.Contains("@סודי ביותר")) comboBox1.SelectedIndex = 3;

                
                isDragged = true;
                mailPath = string.Join("", mi.Subject.Split(Path.GetInvalidFileNameChars()));
                mailPath += ".msg";
                extension = "msg";
                mailPath = "c:\\temp\\" + mailPath;
                mi.SaveAs(mailPath);
        

            }

            else
            {

                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Any())
                {
                  


                    FileInfo fi = new FileInfo(files[0]);
                    mailPath = fi.FullName;
                    extension = fi.Extension;
                    extension = extension.Remove(0, 1);
                    dateTimePicker1.Value = File.GetLastWriteTime(files[0]).Date;
                    textBox1.Text = fi.Name;
                }






            }

            curDocWI = WayIn.drag;
            panel1.Visible = true;
            pnlWayInFiles.Visible = false;
            changeShotefContrlVisibility(false);
            changeDataControlsVisiblity(true);
            changeTemplateControlsVisiblity(true);
            changeInputControlsVisiblity(false);
            changeSelectedFileControlsVisibility(false);
            panel_dropDown.Height = 61;
            this.TopMost = false;
        }

        private void panel_dropDown_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)// שכפול שוטף קיים
        {
            dataGridView2.Visible = false;
            button5.Visible = false;
            textBox24.Visible = false;
            label2.Visible = false;
            curDocWI = WayIn.existing;
         //   searched = false;
            panel1.Visible = true;
            customLabel2.Visible = true;// מספר שוטף
            textBox17.Visible = true;// מספר שוטף
            button17.Visible = true; // חפש (שוטף)
            button12.Visible = true;// אישור
            changeTemplateControlsVisiblity(false);
            changeInputControlsVisiblity(false);
            changeSelectedFileControlsVisibility(false);
         //   label11.Visible = false;// הקובץ הנבחר
         //   textBox13.Visible = false;// הקובץ הנבחר
            changeNadonControlsVisiblity(false);
            panel_dropDown.Height = 61;
            this.TopMost = false;

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (startToLookAtIndex0)
            {
                dataGridViewFolders.FirstDisplayedScrollingRowIndex = 0;
                startToLookAtIndex0 = false;
            }
            foreach (DataGridViewRow row in dataGridViewFolders.Rows)
                if (row.Index > dataGridViewFolders.FirstDisplayedScrollingRowIndex)
                    foreach (DataGridViewCell cell in row.Cells)
                        if (cell.Value.ToString().Contains(textBox22.Text))
                        {
                            dataGridViewFolders.FirstDisplayedScrollingRowIndex = row.Index;
                            row.Selected = true;
                            return;
                        }
            MessageBox.Show("לא נמצאו תוצאות חיפוש נוספות.", "חיפוש תיקים", MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            if (TableType == "users")
            {
                dataGridView2.Rows.Clear();
                List<User> FilteredUsers = users.Where(word => word.firstName.Contains(textBox24.Text) || word.userCode.ToString().Contains(textBox24.Text) || word.lastName.Contains(textBox24.Text) || word.job.ToString().Contains(textBox24.Text)).ToList();
                foreach (User u in FilteredUsers)
                    dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job, u.userCode + ";" + u.firstName + ";" + u.lastName + ";" + u.job);
                dataGridView2.Refresh();
            }
            else if (TableType == "folders")
            {
                ClearTable();
                dataGridView2.Columns.Add("Col1", "תיק");
                dataGridView2.Columns.Add("Col2", "שם מקוצר");
                dataGridView2.Columns.Add("Col3", "שם תיק");
                dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                List<Folder> FilteredDirectories = directories.Where(word => word.description.Contains(textBox24.Text) || word.id.ToString().Contains(textBox24.Text) || word.shortDescription.Contains(textBox24.Text)).ToList();
                dataGridView2.Columns[0].DataPropertyName = "id";
                dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                dataGridView2.Columns[2].DataPropertyName = "description";
                dataGridView2.DataSource = FilteredDirectories.Select(item => new
                {
                    item.id,
                    item.shortDescription,
                    item.description
                }).ToList();
                dataGridView2.Refresh();
            }
            else if (TableType == "projects")
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value.ToString().Contains(textBox24.Text)|| row.Cells[1].Value.ToString().Contains(textBox24.Text))
                        row.Visible = true;
                
                    else
                        row.Visible = false;
                }
            }
            else if (TableType == "patterns")
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value.ToString().Contains(textBox24.Text) || row.Cells[1].Value.ToString().Contains(textBox24.Text) || row.Cells[1].Value.ToString().Contains(textBox24.Text))
                        row.Visible = true;

                    else
                        row.Visible = false;
                }
            }

            /*foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //string rowString = string.Join(" ", row.Cells.Cast<DataGridViewCell>().Select(cell => cell.Value?.ToString() ?? string.Empty));
                if (row.Cells[4].Value.ToString().Contains(textBox24.Text))
                    row.Visible = true;
                
                else
                    row.Visible = false;

            }*/
            /*var rowstoshow = new ConcurrentBag<DataGridViewRow>();
            Parallel.ForEach(dataGridView2.Rows.Cast<DataGridViewRow>(), (row, state) =>
            {
                if (row.Cells.Cast<DataGridViewCell>().Any(cell => cell.Value != null && cell.Value.ToString().Contains(textBox24.Text)))
                {
                    rowstoshow.Add(row);
                }
            });

            dataGridView2.Invoke(new System.Action(() =>
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    row.Visible = rowstoshow.Contains(row);
                }
            }));*/
            UpdateRowCount();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            textBox24.Text = "";
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            startToLookAtIndex0 = true;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            int docId;

            if (int.TryParse(textBox17.Text, out docId))
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT kod_sholeah, msd_proiect, kod_sivug_bitchoni, hearot FROM dbo.documents (nolock) WHERE shotef_mismach=@docId",
                    conn);
                comm.Parameters.AddWithValue("@docId", docId);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    textBox2.Text = sdr.GetInt32(0).ToString();
                    textBox9.Text = sdr.GetInt16(1).ToString();
                    comboBox1.SelectedItem = PublicFuncsNvars.getClassificationByEnum(PublicFuncsNvars.getClassification(sdr.GetInt16(2)));
                    textBox8.Text = sdr.GetString(3);

                    conn.Close();
                    comm.CommandText = "SELECT mispar_nose FROM dbo.tiukim WHERE shotef_klali=@docId AND is_rashi=1";
                    conn.Open();
                    sdr = comm.ExecuteReader();
                    if (sdr.Read())
                        textBox12.Text = sdr.GetInt32(0).ToString();

                  //  searched = true;
                }
                conn.Close();

            }
        }

        private void btnOKWayInType_Click(object sender, EventArgs e)//לא בשימוש
        {
            if (rbBrowse.Checked)
                CreateFromBrowse();
            else CreateFromScan();
            pnlWayInFiles.Visible = false;
        }

        private void changeWriterControlsVisibility(bool b)
        {
            customLabel3.Visible = b;
            //textBox18.Visible = b;// שם משפחה כותב
            textBox19.Visible = b;// קוד כותב
            textBox20.Visible = b;// תפקיד כותב
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void changeInputControlsVisiblity(bool b)
        {
            customLabel1.Visible = b;// סימוכין נכנס
            textBox23.Visible = b;// סימוכין נכנס
        }


        private void openNewDocForm(object state)
        {
            NewDocument nd = new NewDocument();
            nd.Activate();
            nd.ShowDialog();
        }
        private void changeDataControlsVisiblity(bool b)
        {
            // lblNadon.Visible = b; // הנדון
            changeNadonControlsVisiblity(b);
            // lblSigner.Visible = b; // חותם
            changeWriterControlsVisibility(b);
            lblSigner.Visible = b;// חותם
            textBox2.Visible = b;// חותם
            textBox3.Visible = b;// // תפקיד חותם
            //textBox4.Visible = b;// שם פרטי חותם
            //textBox16.Visible = b;// שם משפחה חותם
            label3.Visible = b; // סיווג
            label4.Visible = b;// תאריך הזנה
            label5.Visible = b;// תאריך מכתב
            label6.Visible = b;// סימוכין
            label7.Visible = b;// הערות
            label8.Visible = b;// פרויקט
            label10.Visible = b;// תיק
            customLabel3.Visible = b; // כותב

            // textBox1.Visible = b;// נדון
            //textBox2.Visible = b;// חותם
            //textBox3.Visible = b;// // תפקיד חותם
            //textBox4.Visible = b;// שם פרטי חותם
            //textBox16.Visible = b;// שם משפחה חותם
            //textBox5.Visible = b;// שם תיק
            textBox6.Visible = b;// מספר בתיק
            textBox7.Visible = b;// סימוכין
            textBox8.Visible = b;// הערות
            textBox9.Visible = b;// קוד פרויקט
            textBox10.Visible = b;// שם פרויקט
            textBox11.Visible = b;// שם תיק מקוצר
            textBox12.Visible = b;// קוד תיק
            //textBox18.Visible = b;// שם משפחה כותב
            //textBox19.Visible = b;// קוד כותב
            //textBox20.Visible = b;// תפקיד כותב
            //textBox21.Visible = b;// שם פרטי
            //dateTimePicker1.Visible = b;// תאריך המכתב
            //dateTimePicker2.Visible = b;// תאריך הזנה
            comboBox1.Visible = b;// סיווג
            //button4.Visible = !b;
            button12.Visible = b;// אישור
        }


        private void ChangeDataGrid(string tableType)
        {
            if (tableType == "users")
            {
                
                if (dataGridView2.Columns[0].HeaderText != "משתמש")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "משתמש");
                    dataGridView2.Columns.Add("Col1", "שם פרטי");
                    dataGridView2.Columns.Add("Col3", "שם משפחה");
                    dataGridView2.Columns.Add("Col4", "תפקיד");
                    users = PublicFuncsNvars.users.Where(x => x.isActive).ToList();
                    foreach (User u in users)
                        dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
                    textBox24.Text = "";
                }
                UpdateRowCount();
                dataGridView2.Visible = true;
                textBox24.Visible = true;
                label2.Visible = true;
                label2.BringToFront();
                button5.Visible = true;
                button5.BringToFront();
                TableType = "users";
            }
            else if (tableType == "folders")
            {
                textBox24.Text = "";
                if (dataGridView2.Columns[0].HeaderText != "תיק")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "תיק");
                    dataGridView2.Columns.Add("Col2", "שם מקוצר");
                    dataGridView2.Columns.Add("Col3", "שם תיק");
                    dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    directories = PublicFuncsNvars.folders.Where(x => x.isActive && (x.branch == PublicFuncsNvars.curUser.branch ||
                PublicFuncsNvars.curUser.roleType == RoleType.computers|| int.Parse(Global.P_MAX_CLASS)>3)).ToList();// || x.shortDescription == "מפ - 1" || x.shortDescription == "לש - 320")).ToList();
                    dataGridView2.Columns[0].DataPropertyName = "id";
                    dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                    dataGridView2.Columns[2].DataPropertyName = "description";
                    dataGridView2.DataSource = directories.Select(x => new
                    {
                        x.id,
                        x.shortDescription,
                        x.description
                    }).ToList();
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                dataGridView2.DataSource = directories.Select(x => new
                {
                    x.id,
                    x.shortDescription,
                    x.description
                }).ToList();
                dataGridView2.Refresh();
                dataGridView2.Visible = true;
                textBox24.Visible = true;
                label2.Visible = true;
                button5.Visible = true;
                button5.BringToFront();
                label2.BringToFront();
                TableType = "folders";
            }
            else if (tableType == "projects")
            {
                textBox6.Text = "";
                if (dataGridView2.Columns[0].HeaderText != "מספר")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "מספר");
                    dataGridView2.Columns.Add("Col2", "שם פרויטק");
                    dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    projects = PublicFuncsNvars.projects;
                    foreach (KeyValuePair<int, string> p in projects)
                        dataGridView2.Rows.Add(p.Key.ToString(), p.Value);
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                dataGridView2.Visible = true;
                textBox24.Visible = true;
                label2.Visible = true;
                button5.Visible = true;
                button5.BringToFront();
                label2.BringToFront();
                TableType = "projects";
            }
            else if (tableType == "patterns")
            {
                textBox24.Text = "";
                if (dataGridView2.Columns[1].HeaderText != "שם תבנית")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "מספר תבנית");
                    dataGridView2.Columns.Add("Col2", "שם תבנית");
                    dataGridView2.Columns.Add("Col2", "נתיב לתבנית");
                    dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("SELECT onum_word_template, dscr_template,nam_word_template FROM dbo.tm_templ_bhi", conn); //WHERE file_data<> 0x00000000
                    conn.Open();
                    SqlDataReader sdr = comm.ExecuteReader();
                    while (sdr.Read())
                        dataGridView2.Rows.Add(sdr.GetInt16(0), sdr.GetString(1).Trim(), sdr.GetString(2).Trim());

                    conn.Close();
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                dataGridView2.Visible = true;
                textBox24.Visible = true;
                label2.Visible = true;
                button5.Visible = true;
                button5.BringToFront();
                label2.BringToFront();
                TableType = "patterns";
            }
        }
        private void DataGridView2_RowsChanged(object sender, DataGridViewRowsAddedEventArgs e)
        {
            UpdateRowCount();
        }
        private void DataGridView2_RowsChanged(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            UpdateRowCount();
        }
        private void UpdateRowCount()
        {
            label2.Text = "נמצאו " + dataGridView2.Rows.Cast<DataGridViewRow>().Count(row => row.Visible) + " רשומות";
        }
        private void ClearTable()
        {
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
        }
        private void dataGridViewTable_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.Columns[0].HeaderText == "משתמש")
            {
                textBox2.TextChanged -= textBox2_TextChanged;
                textBox2.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox2.TextChanged += textBox2_TextChanged;

                textBox3.TextChanged -= textBox3_TextChanged;
                textBox3.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                Ujob= dataGridView2.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                textBox3.TextChanged += textBox3_TextChanged;

                /*textBox4.TextChanged -= textBox4_TextChanged;
                textBox4.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox4.TextChanged += textBox4_TextChanged;

                textBox16.TextChanged -= textBox16_TextChanged;
                textBox16.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                textBox16.TextChanged += textBox16_TextChanged;*/
            }
            else if (dataGridView2.Columns[1].HeaderText == "שם מקוצר")
            {
                textBox12.TextChanged -= textBox12_TextChanged;
                textBox12.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox12.TextChanged += textBox12_TextChanged;
                textBox11.TextChanged -= textBox11_TextChanged;
                textBox11.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox11.TextChanged += textBox11_TextChanged;
                textBox5.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                    "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@id", conn);
                comm.Parameters.AddWithValue("@id", int.Parse(textBox12.Text));
                conn.Open();
                textBox6.Text = (int.Parse(comm.ExecuteScalar().ToString()) + 1).ToString();
                conn.Close();
                textBox7.Text = textBox11.Text + " - " + textBox6.Text;
            }
            else if (dataGridView2.Columns[0].HeaderText == "פרויטק")
            {
                textBox9.Text = dataGridView2.Rows[dataGridView2.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            }
            else if(dataGridView2.Columns[0].HeaderText == "מספר תבנית")
                textBox14.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }
    }
}
