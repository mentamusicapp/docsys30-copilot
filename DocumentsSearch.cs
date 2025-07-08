//using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Data.SqlTypes;

namespace DocumentsModule
{
    public partial class DocumentsSearch : Form
    {

        List<DataGridViewRow> rows = new List<DataGridViewRow>();
        ToolStripMenuItem tsmiForAtts;
        internal SearchPatternScreen sps = null;
        internal DocumentHandling dh = null;
        string conStr = Global.ConStr;
        bool okID = false, //האם לסנן לפי שוטף
            //okDate = false, //האם לסנן לפי תאריך
            okBranch = false, //האם לסנן לפי ענף
            okPublished = false, //האם לסנן לפי האם הופץ
            okConIn = false, //האם מסונן לפי נכנס שמור
            okConOut = false, //האם מסונן לפי יוצא שמור
            okTopIn = false, //האם מסונן לפי נכנס סודי
            okTopOut = false, //האם מסונן לפי יוצא סודי
            okUncl = false, //האם מסונן לפי בלמ"ס
            okRapat = false, //האם מסונן לפי רפ"ט
            okInactive = false, //האם מסונן לפי לא פעילים
            okSP = false,//האם מסונן לפי רגיש אישי
            okSender = false, okRef = false, okDir = false, okFromPro = false, okToPro = false;
        int row1 = 0, row2 = 0, row3 = 0;
        public static List<KeyValuePair<int, int>> documents;
        List<Folder> directories;
        List<User> users;
        List<string> searchPatterns;
        string strTyped = "";
        Dictionary<int, string> projects;
        private List<int> searchResults = new List<int>();
        int multiplier = 0;
        DataGridViewColumn currentlySortedColumn = null;
        bool startToLookAtIndex0;
        string tableType;
        DataTable DataTable= new DataTable();
        string FilterOfNispah = "";
        string FilterString = "";
        List<int> ListOfNispah = new List<int>();

        public DocumentsSearch()
        {
            InitializeComponent();
            KeyPreview = true;
            dataGridViewDocs.CellMouseEnter += dataGridView1_CellMouseEnter;
            dataGridView2.CellDoubleClick += dataGridViewTable_CellDoubleClick;
            dataGridView2.RowsAdded += DataGridView2_RowsChanged;
            dataGridView2.RowsRemoved += DataGridView2_RowsChanged;
            DocumentsMenu.PathTemplate(this.button1, 30);
            DocumentsMenu.PathTemplate(this.button6, 30);
            DocumentsMenu.PathTemplate(this.button7, 30);
            DocumentsMenu.PathTemplate(this.button10, 30);
            DocumentsMenu.PathTemplate(this.button11, 30);
            DocumentsMenu.PathTemplate(this.button16, 20);
            DocumentsMenu.PathTemplate(this.button2, 30);
            DocumentsMenu.PathTemplate(this.button3, 30);
            DocumentsMenu.PathTemplate(this.button4, 30);
            DocumentsMenu.PathTemplate(this.button5, 30);

        }

        public DocumentsSearch(string shotef)
        {
            InitializeComponent();
            KeyPreview = true;
            DocumentsSearch_Load(null, null);
            int id = int.Parse(shotef);
            if (!PublicFuncsNvars.docExists(id))
            {
                MessageBox.Show("שוטף" + " " + id + " " + "לא קיים", "", MessageBoxButtons.OK, MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                Environment.Exit(1);
                return;
            }
            if (!PublicFuncsNvars.dhFormsOpen.Contains(id))
            {
                KeyValuePair<int, int> d = new KeyValuePair<int, int>(id, PublicFuncsNvars.curUser.userCode); // getDocById(id);

                if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                {
                    Thread docHandleThread = new Thread(openDocumentHandlingForm);
                    docHandleThread.SetApartmentState(ApartmentState.STA);
                    docHandleThread.Start(d.Key);

                }
                else
                {
                    MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    Environment.Exit(1);
                }
            }
        }

        private void DocumentsSearch_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.tiukim' table. You can move, or remove it, as needed.
            //this.tiukimTableAdapter.Fill(this.mantakDBDataSetDocuments.tiukim);

            this.Icon = Global.AppIcon;
            documents = new List<KeyValuePair<int, int>>();

            users = PublicFuncsNvars.users;

            //dateTimePicker2.MaxDate = DateTime.Today;
            //dateTimePicker1.MaxDate = DateTime.Today;
            //dateTimePicker2.Value = dateTimePicker2.Value.AddMonths(-1);

            dataGridViewDocs.CellFormatting += dataGridViewDocs_CellFormatting;


            tsmiForAtts = new ToolStripMenuItem("הצגת נספח");
            tsmiForAtts.Click += rightClickViewDoc;

            ToolStripMenuItem[] m2 = new ToolStripMenuItem[1];
            ToolStripMenuItem useToFilterByDir = new ToolStripMenuItem("השתמש כתיק לסינון");
            useToFilterByDir.Click += setAsDirFilter;
            m2[0] = useToFilterByDir;
            dataGridViewFolders.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewFolders.ContextMenuStrip.Items.AddRange(m2);

            ToolStripMenuItem[] m3 = new ToolStripMenuItem[2];
            ToolStripMenuItem useToFilterBySender = new ToolStripMenuItem("השתמש כשולח לסינון");
            useToFilterBySender.Click += setAsSenderFilter;
            m3[0] = useToFilterBySender;
            ToolStripMenuItem useToFilterByRef = new ToolStripMenuItem("השתמש כמכותב לסינון");
            useToFilterByRef.Click += setAsRefFilter;
            m3[1] = useToFilterByRef;
            dataGridViewUsers.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewUsers.ContextMenuStrip.Items.AddRange(m3);

            /*users = PublicFuncsNvars.users;
            foreach (User u in users)
                dataGridViewUsers.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);

            directories = PublicFuncsNvars.folders;
            foreach (Folder d in directories)
                dataGridViewFolders.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch));

            projects = PublicFuncsNvars.projects;
            foreach (KeyValuePair<int, string> p in projects)
                dataGridViewProjects.Rows.Add(p.Key.ToString(), p.Value);*/

            reloadPatterns();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            comboBox8.SelectedIndex = 6;
            comboBox8.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox9.SelectedIndex = 0;
        }

        private void rightClickViewDocForEdit(object sender, EventArgs e)
        {
            if (row1 >= 0)
            {
                int id = (int)dataGridViewDocs.Rows[row1].Cells["docIdColumn"].Value;
                if (PublicFuncsNvars.isNormalDoc(id))
                {
                    string v = dataGridViewDocs.Rows[row1].Cells["attachmentsCol"].Value.ToString();
                    int res;
                    if (!int.TryParse(v, out res))
                    {
                        int owner = getDocById(id).Value;
                        if (PublicFuncsNvars.isAuthorizedUser(owner, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToEditDoc(id))
                        {
                            ThreadPool.QueueUserWorkItem(viewDocForEdit, id);
                        }
                        else
                        {
                            MessageBox.Show("אינך מורשה/ית לערוך מסמך זה.", "אין הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        }
                    }
                    else
                    {
                        int rowIndex = row1 - 1;
                        v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                        while (!v.Equals("-"))
                        {
                            rowIndex--;
                            v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                        }
                        ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>((int)dataGridViewDocs.Rows[rowIndex].Cells["docIdColumn"].Value, id));
                    }
                }
                else
                {
                    MessageBox.Show("לא ניתן לערוך מסמך בפורמט זה." + Environment.NewLine + "לעוד פרטים אנא פנו לצוות מחשוב.",
                        "עריכה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
        }

        private void viewDocForEdit(object idObj)
        {
            Cursor.Current = Cursors.WaitCursor;
            //  (new PublicFuncsNvars()).viewDocForEdit((int)idObj);
            PublicFuncsNvars.viewDocForEdit((int)idObj);
            Cursor.Current = Cursors.Default;

        }

        private void usePattern(object sender, EventArgs e)//Ahava W. 03/06/2024 Not in use.
        {
            ToolStripMenuItem t = sender as ToolStripMenuItem;
            loadPattern(t.OwnerItem.Text);
        }

        private void loadPattern(string p)//Ahava W. 03/06/2024 Not in use.
        {
            SqlConnection conn = new SqlConnection(conStr);
            SqlCommand comm = new SqlCommand("SELECT fromID, toID, fromDate" +
                ", toDate, subject, senderCode, senderFirstName, senderRole, referencedCode, referencedRole," +
                " directoryCode, branch, isPublished, isActive, docContent, docOrDirect, isForAct, fromProject, toProject, senderLastName FROM dbo.docSearchPatterns WHERE" +
                " userCode=@userCode AND pattName=@pattName", conn);
            comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
            comm.Parameters.AddWithValue("@pattName", p);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                textBox1.Text = sdr.GetInt32(0).ToString();
                textBox2.Text = sdr.GetInt32(1).ToString();
                if (textBox1.Text.Equals("0") && textBox2.Text.Equals("0"))
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                }
                if (!sdr.GetString(2).Trim().Equals(""))
                {
                    dateTimePicker2.Value = DateTime.ParseExact(sdr.GetString(2).Trim(), "dd/MM/yyyy", new CultureInfo("he-IL"));
                    dateTimePicker1.Value = DateTime.ParseExact(sdr.GetString(3).Trim(), "dd/MM/yyyy", new CultureInfo("he-IL"));
                    checkBox2.Checked = false;
                }
                else
                    checkBox2.Checked = true;
                textBox5.Text = sdr.GetString(4).Trim();
                int sc = sdr.GetInt32(5), rc = sdr.GetInt32(8), dc = sdr.GetInt32(10);
                if (sc != 0)
                    textBox7.Text = sc.ToString();
                else
                {
                    textBox8.Text = sdr.GetString(7).Trim();
                    textBox3.Text = sdr.GetString(6).Trim();
                    textBox20.Text = sdr.GetString(19).Trim();
                }
                if (rc != 0)
                    textBox10.Text = rc.ToString();
                else
                    textBox9.Text = sdr.GetString(9).Trim();
                if (dc != 0)
                    textBox12.Text = dc.ToString();
                comboBox3.Text = sdr.GetString(11).Trim();
                comboBox4.Text = sdr.GetString(14).Trim();
                checkBox1.Checked = sdr.GetBoolean(13);
                textBox17.Text = sdr.GetString(14).Trim();
                comboBox2.Text = sdr.GetString(15).Trim();
                comboBox1.Text = sdr.GetString(19).Trim();
                string pr1 = sdr.GetInt32(17).ToString(), pr2 = sdr.GetInt32(18).ToString();
                if (!pr1.Equals("0"))
                    textBox14.Text = pr1;
                else
                    textBox14.Text = "קוד";
                if (!pr2.Equals("0"))
                    textBox13.Text = pr1;
                else
                    textBox13.Text = "קוד";
            }
            conn.Close();
        }

        private void updatePattern(object sender, EventArgs e)//Ahava W. 03/06/2024 Not in use.
        {
            ToolStripMenuItem t = sender as ToolStripMenuItem;
            sps = new SearchPatternScreen(false, t.OwnerItem.Text);
            sps.Activate();
            sps.Show();
            this.Hide();
        }

        private void setAsRefFilter(object sender, EventArgs e)
        {
            textBox10.Text = dataGridViewUsers.Rows[row3].Cells[0].Value.ToString();
            label24.Visible = false;
            dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void setAsSenderFilter(object sender, EventArgs e)
        {
            textBox7.Text = dataGridViewUsers.Rows[row3].Cells[0].Value.ToString();
            label24.Visible = false;
            dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void setAsDirFilter(object sender, EventArgs e)
        {
            textBox12.Text = dataGridViewFolders.Rows[row2].Cells[0].Value.ToString();
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            comboBox5.SelectedIndex = 0;
        }

        private void viewAtts(object sender, EventArgs e)
        {
            if (row1 >= 0)
            {
                DataGridViewCellEventArgs ce = new DataGridViewCellEventArgs(0, row1);
                dataGridView1_CellContentClick(sender, ce);
            }
        }

        private void rightClickViewDoc(object sender, EventArgs e)
        {
            if (row1 >= 0)
            {
                int id = (int)dataGridViewDocs.Rows[row1].Cells["docIdColumn"].Value;
                string v = dataGridViewDocs.Rows[row1].Cells["attachmentsCol"].Value.ToString();
                int res;
                if (!int.TryParse(v, out res))
                {
                    if (PublicFuncsNvars.isAuthorizedUser(getDocById(id).Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id))
                    {
                        ThreadPool.QueueUserWorkItem(viewDoc, id);
                    }
                    else
                    {
                        MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                }
                else
                {
                    int rowIndex = row1 - 1;
                    v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                    while (!v.Equals("-"))
                    {
                        rowIndex--;
                        v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                    }
                    ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>((int)dataGridViewDocs.Rows[rowIndex].Cells["docIdColumn"].Value, id));
                }
            }

        }

        private void viewDoc(object idObj)
        {
            Cursor.Current = Cursors.WaitCursor;
            (new PublicFuncsNvars()).viewDoc((int)idObj);
            Cursor.Current = Cursors.Default;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //button8.Visible = true;
            //button9.Visible = true;
            //button12.Visible = true;
            //button13.Visible = true;
            //comboBox6.Visible = true;
            searchResults.Clear();
            documents.Clear();
            multiplier = 0;
            dataGridView2.Visible = false;
            textBox6.Visible = false;
            label7.Visible = false;
            button14.Visible = false;

            if (textBox1.Text == "" && textBox2.Text == "") okID = false; else okID = true;

            /*if (dataGridViewDocs.DataSource is System.Data.DataTable dataTable)
                dataTable.Clear();//ahava 10/06/2024
            */
            DataTable.Rows.Clear();

            dataGridViewDocs.Refresh();
            DialogResult res = System.Windows.Forms.DialogResult.Yes;
            if (res == DialogResult.Yes)
            {
                Cursor = Cursors.WaitCursor;
                if (getDocuments("shotef_mismach", false) == 0)
                    MessageBox.Show("לא נמצאו מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                //dataGridViewDocs.Sort(dataGridViewDocs.Columns["docIdColumn"], ListSortDirection.Descending);
                //dataGridViewDocs.Columns["docIdColumn"].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Descending;
                //currentlySortedColumn = dataGridViewDocs.Columns["docIdColumn"];
                Cursor = Cursors.Default;
            }
        }

        private int getDocuments(string sortColumn, bool order)
        {
            int res10;
            if (int.TryParse(textBox10.Text, out res10) && PublicFuncsNvars.getUserNameByUserCode(res10) == null && textBox10.Text != "0")
            {
                MessageBox.Show("אין מכותב עם מספר משתמש זה", "חיפוש מכותב שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return -1;
            }
            else if (textBox4.Text.Equals(""))
            {
                MessageBox.Show("אין תיק עם מספר מזהה או שם מזהה אלו", "חיפוש תיק שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return -1;
            }

            SqlConnection conn = new SqlConnection(conStr);
            conn.Open();
            SqlCommand COMMAND = new SqlCommand("SP_GetDocsList_2", conn);
            COMMAND.CommandType = CommandType.StoredProcedure;

            if (okID)
            {
                int res1 = 0, res2;
                if (int.TryParse(textBox1.Text, out res1) || textBox1.Text == "")
                    COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMe", res1));
                else
                    res1 = -1;

                if (int.TryParse(textBox2.Text, out res2))
                    COMMAND.Parameters.Add(new SqlParameter("@P_ShotefAd", res2));
                else if (textBox2.Text == "")
                    res2 = 0;
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
            }
            else
            {
                COMMAND.Parameters.Add(new SqlParameter( "@P_ShotefMe", "0"));
                COMMAND.Parameters.Add(new SqlParameter("@P_ShotefAd", "0"));
            }

            DateTime today = DateTime.Today;
            COMMAND.Parameters.Add(new SqlParameter("@P_TaarichAd", comboBox8.SelectedIndex != 10 ? today.ToString("yyyyMMdd") : ""));
            if (comboBox8.SelectedIndex!=10)
            {
                if (comboBox8.SelectedIndex == 0)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 1)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddDays(-7).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 2)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddMonths(-1).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 3)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddMonths(-2).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 4)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddMonths(-3).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 5)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddMonths(-6).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 6)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddYears(-1).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 7)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddYears(-2).ToString("yyyyMMdd")));
                else if (comboBox8.SelectedIndex == 8)
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddYears(-3).ToString("yyyyMMdd")));
                else
                    COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddYears(-5).ToString("yyyyMMdd")));
            }
            else
                COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", ""));
            
            short res13 = -1, res14 = -1;
            if (short.TryParse(textBox14.Text, out res14))
            {
                if (!projects.ContainsKey(res14))
                {
                    MessageBox.Show("אין פרויקט עם מספר מזהה זה", "חיפוש מפרויקט שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjMe", res14));
            }
            else
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjMe", "0"));

            if (short.TryParse(textBox13.Text, out res13))
            {
                if (!projects.ContainsKey(res13))
                {
                    MessageBox.Show("אין פרויקט עם מספר מזהה זה", "חיפוש עד פרויקט שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjAd", res13));
            }
            else
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjAd", "0"));

            if (res13 != -1 && res14 != -1 && res14 > res13)
            {
                MessageBox.Show("לא ניתן לחפש פרויקט סיום שקטן מפרויקט התחלה", "חיפוש פרויקט שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return -1;
            }
            int rc;
            if (textBox7.Text != "0")
            {
                if (int.TryParse(textBox7.Text, out rc))
                {
                    if (PublicFuncsNvars.getUserNameByUserCode(rc) == null)
                    {
                        MessageBox.Show("אין שולח עם מספר משתמש זה.", "חיפוש שולח שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        return -1;
                    }
                }
                else if (textBox7.Text != "" && !textBox7.Text.Equals("קוד"))
                {
                    MessageBox.Show("מספר שולח לא יכול להכיל תווים שאינם ספרות.", "חיפוש שולח שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
            }
            //צריך להיות בתיעוד
            /*string name = "";
            if (!textBox3.Text.Equals("") && !textBox3.Text.Equals("שם פרטי"))
            {
                name += textBox3.Text;
                if (!textBox20.Text.Equals("") && !textBox20.Text.Equals("שם משפחה"))
                {
                    name += " "+ textBox3.Text;
                }
            }
            else if (!textBox20.Text.Equals("") && !textBox20.Text.Equals("שם משפחה"))
            {
                name += textBox20.Text;
            }*/
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahShem", name));//צריך להיות בתיעוד
            COMMAND.Parameters.Add(new SqlParameter("@P_Nadon", !textBox5.Text.Equals("") ? textBox5.Text : ""));
           // COMMAND.Parameters.Add(new SqlParameter("@P_TafkidTeur", !textBox8.Text.Equals("") && !textBox8.Text.Equals("שם / תפקיד") ? textBox8.Text : ""));//בתיעור
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahShem", ""));
            //COMMAND.Parameters.Add(new SqlParameter("@P_TafkidTeur", ""));//למחוק
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahTeur", !textBox8.Text.Equals("") && !textBox8.Text.Equals("שם / תפקיד") ? textBox8.Text : ""));
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahKod", int.TryParse(textBox7.Text, out rc) ? rc : 0));
            COMMAND.Parameters.Add(new SqlParameter("@P_Mehutav",(!textBox9.Text.Equals("") && !textBox9.Text.Equals("שם / תפקיד")) ? textBox9.Text : ""));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsLePeula", (!comboBox1.Text.Equals("הכל")) ? comboBox1.SelectedIndex -1 : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsPail", checkBox1.Checked ? 1 : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_Tik", (!textBox4.Text.Equals("") && !textBox4.Text.Equals("שם תיק")) ? long.Parse(textBox12.Text) : 0));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsHufats", (!comboBox4.Text.Equals("הכל")) ? comboBox4.SelectedIndex - 1 : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_SugMismach", (!comboBox2.Text.Equals("הכל")) ? comboBox2.SelectedIndex - 1 : 0));
            COMMAND.Parameters.Add(new SqlParameter("@P_Top", (!comboBox9.Text.Equals("הכל")) ? int.Parse(comboBox9.Text) : 999999999));
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahAnaf", (!comboBox3.Text.Equals("הכל") && (textBox7.Text.Equals("קוד") || textBox7.Text.Equals("") || textBox7.Text.Equals("99999") ||
                users.Where(x => x.userCode == int.Parse(textBox7.Text)).ToList()[0].branch != (Branch)PublicFuncsNvars.getBranchByString(comboBox3.Text))) ? PublicFuncsNvars.getBranchByString(comboBox3.Text) : '\0'));
            CleanFilters();
            FilterString = "";
            using (SqlDataReader reader= COMMAND.ExecuteReader())
            {
                DataTable dataTable = new DataTable();
                DataTable.Clear();
                DataTable.Load(reader);
                conn.Close();
                
                try
                {
                    dataGridViewDocs.DataSource = DataTable.Select("Nispah = 0").CopyToDataTable();
                }
                catch
                {
                    dataGridViewDocs.DataSource = DataTable;
                }
                FilterOfNispah = "";
                documents = DataTable.AsEnumerable().Select(row =>
                {
                    return new KeyValuePair<int, int>(row.Field<int>("Shotef"), row.Field<int>("SholeahKod"));
                }).ToList();
                dataGridViewDocs.Columns["Nikud"].Visible = false;
                dataGridViewDocs.Columns["Taarich2"].Visible = false;
                dataGridViewDocs.Columns["SholeahKod"].Visible = false;
                dataGridViewDocs.Columns["Txt"].Visible = false;
                dataGridViewDocs.Refresh();
                label16.Visible = true;
                label16.Text = "נמצאו " + dataGridViewDocs.Rows.Cast<DataGridViewRow>().Count(row => row.Visible) + " מסמכים";
                /*
                dataGridViewDocs.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
                dataGridViewDocs.ContextMenuStrip.Items.AddRange(tsmiForDocs);*/
                Cursor = Cursors.Default;
                return dataGridViewDocs.RowCount;
            }
            /*command3 += " AND CONTAINS(dbo.docnisp.file_data, @sp0)";//בדיקה לפי תוכן המסמך
                command2 += " AND CONTAINS(file_data, @sp0)";*/
        }

        private bool displayNext15Docs()//למחוק
        {
            return true;
        }

        private void dataGridViewDocs_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridViewDocs.Rows[e.RowIndex].Cells["DocStatus"].Value.ToString()=="הופץ")
            {
                dataGridViewDocs.Rows[e.RowIndex].DefaultCellStyle.Font = new Font(dataGridViewDocs.Font, FontStyle.Bold);
            }
            /*if (!Convert.ToBoolean(dataGridViewDocs.Rows[e.RowIndex].Cells["ragish"].Value))
            {
                ToolStripMenuItem[] tsmiForDocs = new ToolStripMenuItem[4];
                ToolStripMenuItem viewDocMenu = new ToolStripMenuItem("הצגת מסמך");
                viewDocMenu.Click += rightClickViewDoc;
                tsmiForDocs[0] = viewDocMenu;
                ToolStripMenuItem viewDocForEditMenu = new ToolStripMenuItem("עריכת מסמך");
                viewDocForEditMenu.Click += rightClickViewDocForEdit;
                tsmiForDocs[1] = viewDocForEditMenu;
                ToolStripMenuItem viewAttMenu = new ToolStripMenuItem("פתח נספחים");
                viewAttMenu.Click += viewAtts;
                tsmiForDocs[2] = viewAttMenu;
                ToolStripMenuItem shareDocMenu = new ToolStripMenuItem("שיתוף מסמך");
                shareDocMenu.Click += shareDoc;
                tsmiForDocs[3] = shareDocMenu;
                dataGridViewDocs.Rows[e.RowIndex].ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
                dataGridViewDocs.Rows[e.RowIndex].ContextMenuStrip.Items.AddRange(tsmiForDocs);
            }*/

        }

        /*private void addDoc(int id, string subject, bool inOrOut, string creationDate, int senderUser, string sender,
            string refferences, short classification, bool isActive, bool isRapat, bool hasAtts, bool isPublished, bool isRagish)
        {
            documents.Add(new KeyValuePair<int, int>(id, senderUser));
            DateTime cd;
            Classification c = PublicFuncsNvars.getClassification(classification);
            string inOrOutWord;
            if (inOrOut)
                inOrOutWord = "נכנס";
            else
                inOrOutWord = "יוצא";

            if (DateTime.TryParseExact(creationDate, "yyyyMMdd", new CultureInfo("he-IL"), DateTimeStyles.None, out cd))
                dataGridViewDocs.Rows.Add("", id, inOrOutWord, subject, sender, refferences, cd, PublicFuncsNvars.getClassificationByEnum(c));
            else
                dataGridViewDocs.Rows.Add("", id, subject, sender, refferences, null);
            DataGridViewRow row = dataGridViewDocs.Rows[dataGridViewDocs.Rows.Count - 1];
            if (isPublished)
                row.DefaultCellStyle.Font = new Font(dataGridViewDocs.Font, FontStyle.Bold);
            ToolStripMenuItem[] tsmiForDocs = new ToolStripMenuItem[4];
            ToolStripMenuItem viewDocMenu = new ToolStripMenuItem("הצגת מסמך");
            viewDocMenu.Click += rightClickViewDoc;
            tsmiForDocs[0] = viewDocMenu;
            ToolStripMenuItem viewDocForEditMenu = new ToolStripMenuItem("עריכת מסמך");
            viewDocForEditMenu.Click += rightClickViewDocForEdit;
            tsmiForDocs[1] = viewDocForEditMenu;
            ToolStripMenuItem viewAttMenu = new ToolStripMenuItem("פתח נספחים");
            viewAttMenu.Click += viewAtts;
            tsmiForDocs[2] = viewAttMenu;
            ToolStripMenuItem shareDocMenu = new ToolStripMenuItem("שיתוף מסמך");
            shareDocMenu.Click += shareDoc;
            tsmiForDocs[3] = shareDocMenu;
            row.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            row.ContextMenuStrip.Items.AddRange(tsmiForDocs);
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT TOP 1 shotef_nisph FROM dbo.docnisp WHERE shotef_mchtv=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
                if (hasAtts)
                    row.Cells["attachmentsCol"].Value = "+";
            conn.Close();
            dataGridViewDocs.Refresh();


            bool allColors = false;

            if (!okConOut && !okTopIn && !okTopOut && !okConIn && !okRapat && !okInactive && !okSP)
                allColors = true;

            if (isRagish)
            {
                row.Cells["ragish"].Value = true;
            }

            if (isPublished)
            {
                row.Cells["col_published"].Value = true;
            }
            if (isActive)
            {
                row.Cells["col_active"].Value = true;
            }
            if (isRapat)
            {
                if (!allColors && !okRapat)
                    row.Visible = false;
                else
                    row.Visible = true;
            }

            else
            {
                switch (c)
                {
                    case Classification.unclassified:
                        if (!allColors && !okUncl)
                            row.Visible = false;
                        break;
                    case Classification.restricted:
                    case Classification.confidetial:
                        if (inOrOut)
                        {
                            if (!allColors && !okConIn)
                                row.Visible = false;
                        }
                        else
                        {
                            if (!allColors && !okConOut)
                                row.Visible = false;
                        }
                        break;
                    case Classification.secret:
                    case Classification.topSecret:

                        if (inOrOut)
                        {
                            if (!allColors && !okTopIn)
                                row.Visible = false;
                        }
                        else
                        {
                            if (!allColors && !okTopOut)
                                row.Visible = false;
                        }
                        break;
                    case Classification.sensitivePersonal:
                        if (!allColors && !okSP)
                            row.Visible = false;
                        break;
                    //יערה שינתה ב-19.03.23 כדי שסיווגים סודי ומעלה יצבעו בסגולה
                    case Classification.topSecret_shos:
                    case Classification.secret_shos:

                        if (inOrOut)
                        {
                            if (!allColors && !okTopIn)
                                row.Visible = false;
                        }
                        else
                        {
                            if (!allColors && !okTopOut)
                                row.Visible = false;
                        }
                        break;
                }
            }
        }*/

        private void shareDoc(object sender, EventArgs e)
        {
            string to = "";
            string cc = "";
            string bcc = "";
            string body = "";
            string mailSubject = "שיתוף מסמך ";
            int id = (int)dataGridViewDocs.Rows[row1].Cells["docIdColumn"].Value;
            string subject = (string)dataGridViewDocs.Rows[row1].Cells["docSubjectColumn"].Value;
            mailSubject += id + " - " + subject;
            List<Tuple<byte[], string, bool>> attachments = new List<Tuple<byte[], string, bool>>();


            string from = PublicFuncsNvars.curUser.email;
            string res = PublicFuncsNvars.CreateShortcut(id);

            body += res;
            PublicFuncsNvars.sendShareMail(from, to, cc, bcc, mailSubject, body, null);
        }

        private bool documentsContainsID(int id)
        {
            foreach (KeyValuePair<int, int> d in documents)
            {
                if (d.Key == id)
                    return true;
            }
            return false;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            ChangeDataGrid("users");
            tableType = "send";
            if (textBox7.Text.Equals("")|| textBox7.Text.Equals("0"))
            {
                textBox8.ReadOnly = false;
                //textBox8.Text = "";
                //textBox8.ReadOnly = true;
                //textBox3.Text = "";
                //textBox20.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox7.Text.Equals("קוד"))
            {
                textBox8.Text = "שם / תפקיד";
                textBox8.ReadOnly = true;
                //textBox3.Text = "שם פרטי";
                //textBox20.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            /*else if (textBox7.Text.Equals("0"))
                textBox8.ReadOnly = false;*/
            else
            {
                int res;
                if (int.TryParse(textBox7.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            textBox8.Text = u.firstName + " " + u.lastName + " - " + u.job;
                            //textBox3.Text = u.firstName;
                            //textBox20.Text = u.lastName;
                            comboBox3.Text = PublicFuncsNvars.getBranchString(u.branch);
                        }
                    }
                textBox8.ReadOnly = true;
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox7.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox7.Text))
                {
                    MessageBox.Show("אין משתמש עם יוזר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox7.Text = "";
                }

            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)//למחוק
        {
            tableType = "send";
            if (textBox20.Text.Equals(""))
            {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox3.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox20.Text.Equals("שם"))
            {
                textBox7.Text = "קוד";
                textBox8.Text = "תפקיד";
                textBox3.Text = "שם פרטי";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[2].Value != null && row.Cells[2].Value.ToString().StartsWith(textBox20.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            tableType = "ref";
            if (textBox10.Text.Equals("")|| textBox10.Text.Equals("0"))
                textBox9.ReadOnly = false;
                //textBox9.Text = "";
            else if (textBox10.Text.Equals("קוד"))
            {
                textBox9.Text = "שם / תפקיד";
                textBox9.ReadOnly = true;
                comboBox1.Text = "הכל";
            }
            /*else if (textBox10.Text.Equals("0"))
                textBox9.ReadOnly = false;*/
            else
            {
                int res;
                if (int.TryParse(textBox10.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            textBox9.Text = u.job;
                        }
                    }
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox10.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox10.Text))
                {
                    MessageBox.Show("אין משתמש עם יוזר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox9.Text = "";
                }
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            tableType = "folders";
            textBox11.TextChanged -= textBox11_TextChanged;
            TextBox tb = null;
            PublicFuncsNvars.directoryByCode(ref textBox12, ref textBox11, ref textBox4, ref tb, "קוד", "שם מקוצר",
                "SELECT shm_mshimh, shm_mkotzr FROM dbo.tm_mesimot WHERE ms_mshimh=@id AND shm_mkotzr<>''", "@id", typeof(int));
            textBox11.TextChanged += textBox11_TextChanged;
            if (!textBox12.Text.Equals("קוד") && !textBox12.Text.Equals(""))
            {
                dataGridView2.Columns[0].DataPropertyName = "id";
                List<Folder> FilteredDirectories = directories.Where(word => word.id.ToString().StartsWith(textBox12.Text)).ToList();
                //dataGridView2.Columns[0].DataPropertyName = "id";
                //dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                //dataGridView2.Columns[2].DataPropertyName = "description";
                dataGridView2.DataSource = FilteredDirectories.Select(item => new
                {
                    item.id,
                    item.shortDescription,
                    item.description,
                    branch = PublicFuncsNvars.getBranchString(item.branch)
                }).ToList();
                dataGridView2.Refresh();
                int numOfRows = dataGridView2.Rows.Cast<DataGridViewRow>().Count(row => row.Visible);
                //int index = 0;
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
                    dataGridView2.FirstDisplayedScrollingRowIndex = 0;
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
                        item.description,
                        branch = PublicFuncsNvars.getBranchString(item.branch)
                    }).ToList();
                    dataGridView2.Refresh();
                }

                /*int index = 0;
                dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Visible && row.Cells[0].Value.ToString().StartsWith(textBox12.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
                if (index == 0 && !dataGridView2.Rows[0].Cells[0].Value.ToString().StartsWith(textBox12.Text))
                {
                    MessageBox.Show("אין תיק עם מספר זה במערכת התואם לחתך התיקים שנבחר.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    textBox12.Text = "";
                }*/
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //if (checkBox3.Checked)
            //{
            //    okID = false;
            //    textBox1.Clear();
            //    textBox2.Clear();
            //    textBox1.Enabled = false;
            //    textBox2.Enabled = false;
            //}
            //else
            //{
            //    okID = true;
            //    textBox1.Enabled = true;
            //    textBox2.Enabled = true;
            //}
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)//לא בשימוש
        {
            if (checkBox2.Checked)
            {
                //okDate = false;
                //dateTimePicker1.ResetText();
                //dateTimePicker1.Enabled = false;
                //dateTimePicker2.Enabled = false;
                //dateTimePicker2.Value = DateTime.Today.AddMonths(-1);
            }
            else
            {
                //okDate = true;
               // dateTimePicker1.Enabled = true;
               // dateTimePicker2.Enabled = true;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text.Equals("הכל"))
            {
                okBranch = false;
            }
            else
            {
                okBranch = true;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text.Equals("הכל"))
            {
                okPublished = false;
            }
            else
            {
                okPublished = true;
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
             if (e.RowIndex >= 0)
            {
                int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                //string v = dataGridViewDocs.Rows[e.RowIndex].Cells["attachmentsCol"].Value.ToString();
                int nispah= (int)dataGridViewDocs.Rows[e.RowIndex].Cells["Nispah"].Value;
                //int res;
                if (nispah==0)//(!int.TryParse(v, out res))
                {
                    if (!PublicFuncsNvars.dhFormsOpen.Contains(id))
                    {
                        KeyValuePair<int, int> d = getDocById(id);
                        if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                        {
                            Thread docHandleThread = new Thread(openDocumentHandlingForm);
                            docHandleThread.SetApartmentState(ApartmentState.STA);
                            docHandleThread.Start(d.Key);
                        }
                        else
                        {
                            MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        }
                    }
                    else
                    {
                        MessageBox.Show("המסך של מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסך מספר פעמים", "מסך פתוח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
                else
                {
                    //int rowIndex = e.RowIndex - 1;
                    //v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                    /*while (!v.Equals("-"))
                    {
                        rowIndex--;
                        v = dataGridViewDocs.Rows[rowIndex].Cells["attachmentsCol"].Value.ToString();
                    }*/
                    ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value, nispah));// id));//rowIndex-e.RowIndex
                }
            }
        }

        private void openDocumentHandlingForm(object obj)
        {
            int d = (int)obj;
            dh = new DocumentHandling(d);
            dh.Activate();
            dh.ShowDialog();
        }

        private KeyValuePair<int, int> getDocById(int id)
        {
            foreach (KeyValuePair<int, int> doc in documents)
            {
                if (doc.Key == id)
                    return doc;
            }
            return new KeyValuePair<int, int>(-1, -1);
        }

        private void viewAtt(object idObj)
        {
            KeyValuePair<int, int> ids = (KeyValuePair<int, int>)idObj;
            PublicFuncsNvars.viewAtt(ids.Key, ids.Value);
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                row1 = e.RowIndex;
                dataGridViewDocs.Rows[row1].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void textBox7_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox7.Text, out res) || textBox7.Text == "0")
            {
                textBox7.Text = "";
            }
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            if (textBox7.Text.Equals(""))
                textBox7.Text = "קוד";
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox10.Text, out res))
            {
                textBox10.Text = "";
            }
        }

        private void textBox10_Leave(object sender, EventArgs e)
        {
            if (textBox10.Text.Equals(""))
                textBox10.Text = "קוד";
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

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            tableType = "folders";
            textBox12.TextChanged -= textBox12_TextChanged;
            TextBox tb = null;
            PublicFuncsNvars.directoryByCode(ref textBox11, ref textBox12, ref textBox4, ref tb, "שם מקוצר", "קוד",
                "SELECT shm_mshimh, ms_mshimh FROM dbo.tm_mesimot WHERE shm_mkotzr=@shortName", "@shortName", typeof(string));
            textBox12.TextChanged += textBox12_TextChanged;
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
                    item.description,
                    branch = PublicFuncsNvars.getBranchString(item.branch)
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
                        item.description,
                        branch = PublicFuncsNvars.getBranchString(item.branch)
                    }).ToList();
                    dataGridView2.Refresh();
                }

                /*int index = 0;
                dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);
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
            strTyped = "";
            //label25.Visible = true;
            //customLabel1.Visible = true;
            //comboBox5.Visible = true;
            //dataGridViewFolders.Visible = true;
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

        //private bool legalChar(char c)
        //{
        //    if ((c > '0' && c < '9') || (c > 'a' && c < 'z') || (c > 'A' && c < 'Z') || (c > 'א' && c < 'ת') || c == '.' || c == ' ' || c == '-' || c == '_' || c == '/' || c == '\\' || c == '*' || c == '+' || c == '=' || c == '!' || c == '@' || c == '#' || c == '$' || c == '%' || c == '^' || c == '(' || c == ')' || c == '~' || c == ':')
        //        return true;
        //    return false;
        //}

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
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row3 = e.RowIndex;
                dataGridViewUsers.Rows[row3].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            string value = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
            if (okSender)
            {
                textBox7.TextChanged -= textBox7_TextChanged;
                textBox7.Text = value;
                textBox7.TextChanged += textBox7_TextChanged;

                textBox8.TextChanged -= textBox8_TextChanged;
                textBox8.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                textBox8.TextChanged += textBox8_TextChanged;

                /*textBox3.TextChanged -= textBox3_TextChanged;
                textBox3.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox3.TextChanged += textBox3_TextChanged;

                textBox20.TextChanged -= textBox20_TextChanged;
                textBox20.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                textBox20.TextChanged += textBox20_TextChanged;*/
            }
            else if (okRef)
            {
                textBox10.TextChanged -= textBox10_TextChanged;
                textBox10.Text = value;
                textBox10.TextChanged += textBox10_TextChanged;

                textBox9.TextChanged -= textBox9_TextChanged;
                textBox9.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                textBox9.TextChanged += textBox9_TextChanged;
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
                textBox4.Text = dataGridViewFolders.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
            }
        }

        private void dataGridView2_Leave(object sender, EventArgs e)
        {

        }

        private void DocumentHandling_Click(object sender, EventArgs e)
        {
            label24.Visible = false;
            dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            dataGridView2.Visible = false;
            button14.Visible = false;
            textBox6.Visible = false;
            label7.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            dataGridViewProjects.Visible = false;
            button10.Visible = false;
            button11.Visible = false;
            textBox22.Visible = false;
            button16.Visible = false;
            comboBox5.SelectedIndex = 0;
        }

        private void selectingSender()
        {
            //label24.Visible = true;
            //dataGridViewUsers.Visible = true;
            tableType = "send";
            dataGridView2.Visible = true;
            ChangeDataGrid("users");
            //button3.Visible = true;
            //button5.Visible = true;
            //label25.Visible = false;
            //customLabel1.Visible = false; לבדוק
            //comboBox5.Visible = false; לבדוק
            //dataGridViewFolders.Visible = false;
            //button2.Visible = false;
            //button4.Visible = false;
            okSender = true;
            okRef = false;
            //makeProjectsTableInVisible();
            //comboBox5.SelectedIndex = 0;לבדוק
        }

        private void selectingRecipient()
        {
            //label24.Visible = true;
            //dataGridViewUsers.Visible = true;
            dataGridView2.Visible = true;
            ChangeDataGrid("users");
            tableType = "ref";
            //button3.Visible = true;
            //button5.Visible = true;
            //label25.Visible = false;
            //customLabel1.Visible = false;
            //comboBox5.Visible = false;
            //dataGridViewFolders.Visible = false;
            //button2.Visible = false;
            //button4.Visible = false;
            okSender = false;
            okRef = true;
            //makeProjectsTableInVisible();
            //comboBox5.SelectedIndex = 0;
        }

        private void textBox7_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox10_Click(object sender, EventArgs e)
        {
            selectingRecipient();
        }

        private void textBox19_Click(object sender, EventArgs e)
        {
            label24.Visible = true;
            dataGridViewUsers.Visible = true;
            button3.Visible = true;
            button5.Visible = true;
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            okSender = false;
            okRef = false;
            makeProjectsTableInVisible();
            comboBox5.SelectedIndex = 0;
        }

        private void textBox12_Click(object sender, EventArgs e)
        {
            //label25.Visible = true;
            //customLabel1.Visible = true; kcsue
            //comboBox5.Visible = true;
            //dataGridViewFolders.Visible = true;
            tableType = "folders";
            dataGridView2.Visible = true;
            ChangeDataGrid("folders");
            //button2.Visible = true;
            //button4.Visible = true;
            //label24.Visible = false;
            //dataGridViewUsers.Visible = false;
            //button3.Visible = false;
            //button5.Visible = false;
            okDir = true;
            //makeProjectsTableInVisible();

        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            tableType = "folders";
            dataGridView2.Visible = true;
            ChangeDataGrid("folders");
            // makeProjectsTableInVisible();
            //makeUsersTableVisibleOrInvisible(false);
            //label25.Visible = true;
            //customLabel1.Visible = true;
            //comboBox5.Visible = true;
            //dataGridViewFolders.Visible = true;
            //button2.Visible = true;
            //button4.Visible = true;
            //label24.Visible = false;
            //dataGridViewUsers.Visible = false;
            //button3.Visible = false;
            // button5.Visible = false;
            okDir = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            textBox6.Visible = false;
            label7.Visible = false;
            button14.Visible = false;
            checkBox10.Checked = false;
            checkBox9.Checked = false;
            checkBox6.Checked = false;
            checkBox5.Checked = false;
            checkBox4.Checked = false;
            //button8.Visible = false;
            //button9.Visible = false;
            // button12.Visible = false;
            //button13.Visible = false;
            comboBox6.Visible = false;
            //  checkBox3.Checked = true;
            //checkBox2.Checked = true;
            textBox5.Clear();
            textBox7.Text = "קוד";
            textBox10.Text = "קוד";
            textBox12.Text = "קוד";
            textBox14.Clear();
            textBox13.Clear();
            comboBox2.Text = "הכל";
            comboBox8.SelectedIndex = 6;
            comboBox3.Text = "הכל";
            comboBox4.Text = "הכל";
            comboBox7.Text = "הכל";
            checkBox1.Checked = false;
            //textBox17.Clear();
            DataTable.Rows.Clear();
            dataGridViewDocs.DataSource = DataTable;
            dataGridViewDocs.Refresh();
            //dataGridViewDocs.Rows.Clear();
            //dateTimePicker1.Value = DateTime.Today;
            //dateTimePicker2.Value = dateTimePicker1.Value.AddMonths(-1);
            textBox1.Clear();
            textBox2.Clear();
            label16.Visible = false;
            textBox6.Visible = false;
            label7.Visible = false;
            button14.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (okSender)
                textBox7.Text = PublicFuncsNvars.curUser.userCode.ToString();
            else if (okRef)
                textBox10.Text = "קוד";
            label24.Visible = false;
            //dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            label24.Visible = false;
            dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox12.Text = "קוד";
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            comboBox5.SelectedIndex = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            comboBox5.SelectedIndex = 0;
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 5) && e.RowIndex >= 0)
            {
                DataGridViewCell cell = dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string cellValue = Convert.ToString(cell.Value);
                dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = cellValue;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.ColumnIndex == 10|| e.ColumnIndex == 9) && e.RowIndex >= 0) //0-10-9
            {
                int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                KeyValuePair<int, int> doc = getDocById(id);
                if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(doc.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))//if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                {
                    DataGridViewRow row = dataGridViewDocs.Rows[e.RowIndex];
                    int a = dataGridViewDocs.FirstDisplayedScrollingRowIndex;
                    if (row.Cells["attachmentsCol"].Value.Equals("+"))
                    {
                        ListOfNispah.Add(id);
                        FilterOfNispah += FilterOfNispah == "" ? $"Nispah = 0 or Shotef='{id}'" : $" or Shotef='{id}'";
                        DataTable data;
                        if (FilterString == "")
                        {
                            data = DataTable.Select(FilterOfNispah).CopyToDataTable();
                        }
                        else
                        {
                            string filterstring = $"({FilterString}){FilterOfNispah.Replace("Nispah = 0", "")}";

                            DataRow[] dataRow = DataTable.Select(filterstring);
                            //dataRow = dataRow.Select(filterstring);
                            data = dataRow.CopyToDataTable();
                        }
                        //data = data.Select(FilterString).CopyToDataTable();
                        dataGridViewDocs.DataSource = data;
                        dataGridViewDocs.Columns[e.ColumnIndex].ReadOnly = false;
                        row.Cells[e.ColumnIndex].Value = "-";
                        BindingSource bs = new BindingSource();
                        bs.DataSource = data;
                        dataGridViewDocs.DataSource = bs;
                        data.DefaultView.Sort = "Shotef DESC, Nispah ASC";
                        DataView datav = data.DefaultView;
                        foreach (int idNispah in ListOfNispah)
                        {
                            int index = datav.Cast<DataRowView>()
                                .Select((rowv, idx) => new { RowView = rowv, Index = idx }).FirstOrDefault(x => x.RowView.Row.Field<int>("Shotef") == idNispah)?.Index ?? -1;
                            dataGridViewDocs.Rows[index].Cells[e.ColumnIndex].Value = "-";
                            dataGridViewDocs.CommitEdit(DataGridViewDataErrorContexts.Commit);
                            int ia = 1;

                            while (dataGridViewDocs.Rows[index + ia].Cells["Nispah"].Value.ToString() != "0")
                            {
                                dataGridViewDocs.Rows[index + ia].Cells[e.ColumnIndex].Value = ia;
                                dataGridViewDocs.Rows[index + ia].DefaultCellStyle.ForeColor = Color.Maroon;
                                dataGridViewDocs.Rows[index + ia].DefaultCellStyle.BackColor = row.DefaultCellStyle.BackColor;
                                /*dataGridViewDocs.Rows[index + ia].ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
                                dataGridViewDocs.Rows[index + ia].ContextMenuStrip.Items.Add(tsmiForAtts);*/
                                ia++;
                                if (index + ia >= dataGridViewDocs.Rows.Count)
                                    break;
                            }
                        }
                        dataGridViewDocs.Refresh();
                    }
                    else if (row.Cells["attachmentsCol"].Value.Equals("-"))
                    {
                        try
                        {
                            while (dataGridViewDocs.Rows[e.RowIndex + 1].Cells["Nispah"].Value.ToString() != "0")
                                dataGridViewDocs.Rows.RemoveAt(e.RowIndex + 1);
                        }
                        catch { }
                        row.Cells["attachmentsCol"].Value = "+";
                        if (FilterOfNispah.Contains($" or Shotef='{id}'"))
                            FilterOfNispah=FilterOfNispah.Replace($" or Shotef='{id}'", "");
                        Console.WriteLine(FilterOfNispah);
                        ListOfNispah.Remove(id);
                    }
                    dataGridViewDocs.FirstDisplayedScrollingRowIndex = a;

                }
                else
                {
                    MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
        }

        private void newPatternToolStripMenuItem_Click(object sender, EventArgs e)//Ahava W. 03/06/2024 Not in use.
        {
            sps = new SearchPatternScreen(true, null);
            sps.Activate();
            sps.Show();
            this.Hide();
        }

        internal void reloadPatterns()//Ahava W. 03/06/2024 Not in use.
        {
            exsitingPatternsToolStripMenuItem.DropDownItems.Clear();
            SqlConnection conn = new SqlConnection(conStr);
            SqlCommand comm = new SqlCommand("SELECT pattName FROM dbo.docSearchPatterns where userCode=@userCode", conn);
            comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.curUser.userCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            searchPatterns = new List<string>();
            while (sdr.Read())
                searchPatterns.Add(sdr.GetString(0).Trim());
            conn.Close();

            foreach (string pat in searchPatterns)
            {
                ToolStripMenuItem updateTsmi = new ToolStripMenuItem();
                updateTsmi.Name = "updateToolStripMenuItem";
                updateTsmi.Text = "עדכן תבנית";
                updateTsmi.Click += new EventHandler(updatePattern);

                ToolStripMenuItem useTsmi = new ToolStripMenuItem();
                useTsmi.Name = "useToolStripMenuItem";
                useTsmi.Text = "השתמש בתבנית";
                useTsmi.Click += new EventHandler(usePattern);

                ToolStripMenuItem tsmi = new ToolStripMenuItem();
                tsmi.Text = pat;
                tsmi.DropDownItems.Add(updateTsmi);
                tsmi.DropDownItems.Add(useTsmi);
                exsitingPatternsToolStripMenuItem.DropDownItems.Add(tsmi);
            }
        }

        private void DocumentsSearch_FormClosed(object sender, FormClosedEventArgs e)
        {
            Program.dm.ds = null;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            makeProjectsTableInVisible();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (okFromPro)
                textBox14.Text = "קוד";
            else if (okToPro)
                textBox13.Text = "קוד";
            makeProjectsTableInVisible();
        }

        private void makeProjectsTableInVisible()
        {
            Control[] controls = { label26, dataGridViewProjects, button10, button11, button16, textBox22 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            tableType = "fromPro";
            try
            {
                int res;
                if (int.TryParse(textBox14.Text, out res))
                {
                    textBox15.Text = projects[res];
                    textBox16.Text = projects[res];
                    textBox13.Text = textBox14.Text;
                }


                else
                {
                    textBox15.Text = "שם פרויקט";
                    textBox16.Text = "שם פרויקט";
                }
            }
            catch (Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show("אין פרויקט עם מספר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                textBox14.Text = "";
                textBox13.Text = "";
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            tableType = "toPro";
            try
            {
                int res;
                if (int.TryParse(textBox13.Text, out res))
                    textBox16.Text = projects[res];
                else
                    textBox16.Text = "שם פרויקט";
            }
            catch (Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show("אין פרויקט עם מספר זה במערכת", "", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                textBox13.Text = "";
            }
        }

        private void textBox14_Click(object sender, EventArgs e)
        {
            strTyped = "";
            //label26.Visible = true;
            //dataGridViewProjects.Visible = true;
            dataGridView2.Visible = true;
            ChangeDataGrid("projects");
            tableType = "fromPro";
            //button10.Visible = true;
            //button11.Visible = true;
            //makeUsersTableVisibleOrInvisible(false);
            //makeDirectoriesTableInVisible();
            okFromPro = true;
            okToPro = false;
            //textBox22.Visible = true;
            //button16.Visible  = true;
        }

        private void textBox15_Click(object sender, EventArgs e)
        {
            strTyped = "";
            //label26.Visible = true;
            //dataGridViewProjects.Visible = true;
            dataGridView2.Visible = true;
            ChangeDataGrid("projects");
            tableType = "fromPro";
            //button10.Visible = true;
            //button11.Visible = true;
            //makeUsersTableVisibleOrInvisible(false);
            //makeDirectoriesTableInVisible();
            okFromPro = true;
            okToPro = false;
        }
        private void textBox16_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = true;
            ChangeDataGrid("projects");
            tableType = "toPro";
            strTyped = "";
            okFromPro = false;
            okToPro = true;
        }
        private void textBox13_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = true;
            ChangeDataGrid("projects");
            tableType = "toPro";
            strTyped = "";
            //label26.Visible = true;
            //dataGridViewProjects.Visible = true;
            //button10.Visible = true;
            //button11.Visible = true;
            //makeUsersTableVisibleOrInvisible(false);
            //makeDirectoriesTableInVisible();
            okFromPro = false;
            okToPro = true;
            //textBox22.Visible = true;
            //button16.Visible = true;


        }

        private void makeDirectoriesTableInVisible()
        {
            Control[] controls = { label25, customLabel1, comboBox5, dataGridViewFolders, button2, button4 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
            comboBox5.SelectedIndex = 0;
        }

        private void makeUsersTableVisibleOrInvisible(bool b)
        {
            Control[] controls = { label24, dataGridViewUsers, button3, button5 };
            PublicFuncsNvars.changeControlsVisiblity(b, controls.ToList());
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (okFromPro)
                textBox14.Text = dataGridViewProjects.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
            else if(okToPro)
                textBox13.Text = dataGridViewProjects.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }
        
        private void dataGridView4_KeyPress(object sender, KeyPressEventArgs e)
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

        private void dataGridView4_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView4_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row1 = e.RowIndex;
                dataGridViewProjects.Rows[row1].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)//למחוק
        {
            if(dateTimePicker1.Value<dateTimePicker2.Value)
            {
                MessageBox.Show("לא ניתן לבחור תאריך סיום הקודם לתאריך ההתחלה", "בחירת תאריך שגויה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)//למחוק
        {
            if (dateTimePicker1.Value < dateTimePicker2.Value)
            {
                MessageBox.Show("לא ניתן לבחור תאריך התחלה המאוחר מתאריך הסיום", "בחירת תאריך שגויה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }

        private void exsitingPatternsToolStripMenuItem_Click(object sender, EventArgs e)//Ahava W. 03/06/2024 Not in use.
        {
            if (searchPatterns.Count == 0)
            {
                MessageBox.Show("אין לך תבניות חיפוש קיימות", "אין תבניות חיפוש", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            tableType = "send";
            if (textBox8.Text.Equals(""))
            {
                if (textBox7.Text != "0")
                    textBox7.Text = "";
                // textBox3.Text = "";
                // textBox20.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox8.Text.Equals("שם / תפקיד"))
            {
                textBox7.Text = "קוד";
                //textBox3.Text = "שם פרטי";
                //textBox20.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[3].Value != null && row.Cells[3].Value.ToString().StartsWith(textBox8.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)//למחוק
        {
            tableType = "send";
            if (textBox3.Text.Equals(""))
            {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox20.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox3.Text.Equals("שם פרטי"))
            {
                textBox7.Text = "קוד";
                textBox8.Text = "תפקיד";
                textBox20.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString().StartsWith(textBox3.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            tableType = "ref";
            if (textBox9.Text.Equals(""))
            {
                if (textBox10.Text != "0")
                    textBox10.Text = "";
            }
            else if (textBox9.Text.Equals("שם / תפקיד"))
            {
                textBox10.Text = "קוד";
                comboBox1.Text = "הכל";
            }
            else
            {
                int index = 0;
                dataGridView2.Sort(dataGridView2.Columns[3], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[3].Value != null && row.Cells[3].Value.ToString().StartsWith(textBox9.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox8.Text.Equals("שם / תפקיד"))
            {
                strTyped = "";
                textBox8.Text = "";
            }
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            if (textBox8.Text.Equals(""))
                textBox8.Text = "שם / תפקיד";
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox3.Text.Equals("שם פרטי"))
            {
                strTyped = "";
                textBox3.Text = "";
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
                textBox3.Text = "שם פרטי";
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals("שם / תפקיד"))
            {
                strTyped = "";
                textBox9.Text = "";
            }
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
                textBox9.Text = "שם / תפקיד";
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
        }

        private void textBox14_Leave(object sender, EventArgs e)
        {
            if (textBox14.Text.Equals(""))
                textBox14.Text = "קוד";
        }

        private void textBox13_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox13.Text, out res))
            {
                textBox13.Text = "";
            }
        }

        private void textBox13_Leave(object sender, EventArgs e)
        {
            if (textBox13.Text.Equals(""))
                textBox13.Text = "קוד";
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            try
            {
                double from;
                double to;

                from = double.Parse(textBox1.Text);
                to = double.Parse(textBox2.Text);

                if (from > to) textBox2.Text = textBox1.Text;
            }

            catch { }
        }

        private void DocumentsSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            
                if (startToLookAtIndex0)
                {
                dataGridViewProjects.FirstDisplayedScrollingRowIndex = 0;
                    startToLookAtIndex0 = false;
                }
                foreach (DataGridViewRow row in dataGridViewProjects.Rows)
                    if (row.Index > dataGridViewProjects.FirstDisplayedScrollingRowIndex)
                        foreach (DataGridViewCell cell in row.Cells)
                            if (cell.Value.ToString().Contains(textBox22.Text))
                            {
                            dataGridViewProjects.FirstDisplayedScrollingRowIndex = row.Index;
                                row.Selected = true;
                                return;
                            }
                MessageBox.Show("לא נמצאו תוצאות חיפוש נוספות.", "חיפוש תיקים", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            startToLookAtIndex0 = true;
        }

        private void dataGridViewDocs_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void dataGridViewDocs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Shift && e.KeyCode == Keys.Delete)
            {
                var y = dataGridViewDocs.SelectedRows;
                foreach( DataGridViewColumn col in dataGridViewDocs.Columns)
                {
                    if (col.DataPropertyName == "Shotef")
                    {
                        string shotef = y[0].Cells[col.Index].Value.ToString();

                        var x = MessageBox.Show("To Delete " + shotef + " ?", "מחיקת שוטף", MessageBoxButtons.YesNo);
                        if (x == DialogResult.Yes)
                        {
                            SqlConnection conn = new SqlConnection(Global.ConStr);
                            SqlCommand cmd = new SqlCommand("dbo.SP_DeleteDocument", conn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add(new SqlParameter("@ShoteToDel", shotef));
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();

                            MessageBox.Show("DELETED " + shotef);
                        }
                    }
                }
                
            }
        }
        

        private void customLabel2_Click(object sender, EventArgs e)
        {
          
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)//בלמ"ס
        {
            FilterOfDocTable();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)// שמור
        {
            FilterOfDocTable();
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)//  סודי
        {
            FilterOfDocTable();        }
        

        private void checkBox9_CheckedChanged(object sender, EventArgs e)//רגיש
        {
            FilterOfDocTable();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)//יוצא, נכנס, הכל
        {
            FilterOfDocTable();
        }

        private void FilterOfDocTable()
        {
            FilterString = "";
            if (dataGridViewDocs.Rows.Count > 0)
            {
                if (checkBox6.Checked || checkBox5.Checked || checkBox4.Checked)
                {
                    if (checkBox4.Checked)
                        FilterString += FilterString == "" ? "(Sivug = 'בלמ\"ס'" : " OR Sivug = 'בלמ\"ס'";
                    if (checkBox5.Checked)
                        FilterString += FilterString == "" ? "(Sivug = 'שמור'" : " OR Sivug = 'שמור'";
                    if (checkBox6.Checked)
                        FilterString += FilterString == "" ? "(Sivug = 'סודי'" : " OR Sivug = 'סודי'";
                    FilterString += ")";
                }
                if (checkBox9.Checked)
                    FilterString += FilterString != "" ? " AND isRagish = 1" : "isRagish = 1";
                if (checkBox10.Checked)
                    FilterString += FilterString != "" ? " AND isPail = 0" : "isPail = 0";
                if (comboBox7.SelectedIndex != 0)
                {
                    if (comboBox7.Text == "נכנס")
                        FilterString += FilterString != "" ? " AND Makor= 'נכנס'" : "Makor= 'נכנס'";
                    else if (comboBox7.Text == "יוצא")
                        FilterString += FilterString != "" ? " AND Makor= 'יוצא'" : "Makor= 'יוצא'";
                }
                FilterString += FilterString != "" ? " AND Nispah = 0" : "Nispah = 0";
                int NumOfRows = DataTable.Select(FilterString).Count();
                if (NumOfRows > 0)
                    dataGridViewDocs.DataSource = DataTable.Select(FilterString).CopyToDataTable();
                else
                {
                    MessageBox.Show("לא נמצאו מסמכים לסינון."); dataGridViewDocs.DataSource = DataTable;
                    CleanFilters();
                }
                dataGridViewDocs.Refresh();
                label16.Text = "נמצאו " + dataGridViewDocs.Rows.Cast<DataGridViewRow>().Count(row => row.Visible) + " מסמכים";
            }
            else
                CleanFilters();
        }

        private void CleanFilters()
        {
            checkBox6.Checked = false;
            checkBox5.Checked = false;
            checkBox4.Checked = false;
            checkBox10.Checked = false;
            checkBox9.Checked = false;
            comboBox7.SelectedIndex = 0;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void customLabel2_Click_1(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)//לא פעיל
        {
            FilterOfDocTable();/*
            if (checkBox10.Checked)
            {
                foreach (DataGridViewRow row in dataGridViewDocs.Rows)
                {
                    string classification = row.Cells["docClassificationColumn"].Value.ToString();
                    string InOrOut = row.Cells["docInOrOutColumn"].Value.ToString();
                    bool ragish = Convert.ToBoolean(row.Cells["ragish"].Value);
                    bool active = Convert.ToBoolean(row.Cells["col_active"].Value);
                    if (row.Visible && active)
                    {
                        if (!((checkBox4.Checked && classification == "בלמ\"ס") ||
                            (checkBox6.Checked && classification == "סודי") ||
                            (checkBox5.Checked && classification == "שמור") ||
                            (checkBox9.Checked && ragish)))
                        {
                            row.Visible = false;
                        }
                    }
                    else if (!active)
                    {
                        if (comboBox7.Text.Equals("הכל"))
                            row.Visible = true;
                        else if (comboBox7.Text.Equals("נכנס") && InOrOut == "נכנס")
                            row.Visible = true;
                        else if (comboBox7.Text.Equals("יוצא") && InOrOut == "יוצא")
                            row.Visible = true;
                        else
                            row.Visible = false;
                    }
                }
            }
            else
            {
                if (checkBox6.Checked || checkBox4.Checked || checkBox9.Checked || checkBox5.Checked )
                {
                    foreach (DataGridViewRow row in dataGridViewDocs.Rows)
                    {
                        bool active = Convert.ToBoolean(row.Cells["col_active"].Value);
                        if (!active)
                        {
                            row.Visible = false;
                        }
                    }
                }
                else if (!comboBox7.Text.Equals("הכל"))
                    comboBox7_SelectedIndexChanged(sender, e);
                else
                {
                    foreach (DataGridViewRow row in dataGridViewDocs.Rows)
                        row.Visible = true;
                }
            }
            label24.Visible = false;
            dataGridViewUsers.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
            label25.Visible = false;
            customLabel1.Visible = false;
            comboBox5.Visible = false;
            dataGridViewFolders.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            dataGridViewProjects.Visible = false;
            button10.Visible = false;
            button11.Visible = false;
            textBox22.Visible = false;
            button16.Visible = false;
            comboBox5.SelectedIndex = 0;*/
        }

        
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (tableType == "folders")
            {
                //dataGridView2.DataSource=null;
                ClearTable();
                dataGridView2.Columns.Add("Col1", "תיק");
                dataGridView2.Columns.Add("Col2", "שם מקוצר");
                dataGridView2.Columns.Add("Col3", "שם תיק");
                dataGridView2.Columns.Add("Col4", "ענף");
                dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                List<Folder> FilteredDirectories = directories.Where(word => word.description.Contains(textBox6.Text) || word.id.ToString().Contains(textBox6.Text) || word.shortDescription.Contains(textBox6.Text) || (PublicFuncsNvars.getBranchString(word.branch)).Contains(textBox6.Text)).ToList();
                dataGridView2.Columns[0].DataPropertyName = "id";
                dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                dataGridView2.Columns[2].DataPropertyName = "description";
                dataGridView2.Columns[3].DataPropertyName = "branch";
                dataGridView2.DataSource = FilteredDirectories.Select(item => new
                {
                    item.id,
                    item.shortDescription,
                    item.description,
                    branch = PublicFuncsNvars.getBranchString(item.branch)
                }).ToList();
                //dataGridView2.DataSource = FilteredDirectories;
                //foreach (Folder d in FilteredDirectories)
                // dataGridView2.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch), d.id + ";" + d.shortDescription + ";" + d.description + ";" + PublicFuncsNvars.getBranchString(d.branch));
                dataGridView2.Refresh();
            }
            else if (tableType == "send" || tableType == "ref")
            {
                ClearTable();
                dataGridView2.Columns.Add("Col1", "משתמש");
                dataGridView2.Columns.Add("Col2", "שם פרטי");
                dataGridView2.Columns.Add("Col3", "שם משפחה");
                dataGridView2.Columns.Add("Col4", "תפקיד");
                dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                List<User> FilteredUsers = users.Where(word => word.firstName.Contains(textBox6.Text) || word.userCode.ToString().Contains(textBox6.Text) || word.lastName.Contains(textBox6.Text) || word.userCode.ToString().Contains(textBox6.Text)).ToList();
                foreach (User u in FilteredUsers)
                    dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
                dataGridView2.Refresh();
            }
            else if (tableType == "toPro" || tableType == "fromPro")
            {
                ClearTable();
                dataGridView2.Columns.Add("Col1", "מספר");
                dataGridView2.Columns.Add("Col2", "שם פרויקט");
                dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                List<Folder> FilteredDirectories = directories.Where(word => word.description.Contains(textBox6.Text) || word.id.ToString().Contains(textBox6.Text) || word.shortDescription.Contains(textBox6.Text) || word.shortDescription.Contains(textBox6.Text)).ToList();
                foreach (Folder d in FilteredDirectories)
                    dataGridView2.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch), d.id + ";" + d.shortDescription + ";" + d.description + ";" + PublicFuncsNvars.getBranchString(d.branch));
                dataGridView2.Refresh();
            }


            //dataGridView2.Visible = true;
            /*foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //string rowString = string.Join(" ", row.Cells.Cast<DataGridViewCell>().Select(cell => cell.Value?.ToString() ?? string.Empty));
                if (row.Cells[4].Value.ToString().Contains(textBox6.Text))
                    row.Visible = true;

                else
                    row.Visible = false;

            }*/
            /*foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Cells[0].Value.ToString().Contains(textBox6.Text.ToString()) || row.Cells[1].Value.ToString().Contains(textBox6.Text.ToString()))
                    row.Visible = true;
                else if (row.Cells[2].Value != null)
                {
                    if (row.Cells[2].Value.ToString().Contains(textBox6.Text.ToString()) || row.Cells[3].Value.ToString().Contains(textBox6.Text.ToString()))
                        row.Visible = true;
                    else
                        row.Visible = false;
                }
                else
                    row.Visible = false;
            }
            var rowstoshow = new ConcurrentBag<DataGridViewRow>();
            Parallel.ForEach(dataGridView2.Rows.Cast<DataGridViewRow>(), (row, state) =>
            {
                if (row.Cells.Cast<DataGridViewCell>().Any(cell => cell.Value != null && cell.Value.ToString().Contains(textBox6.Text)))
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
            label7.Visible = true;
            label7.Text = "נמצאו " + dataGridView2.Rows.Cast<DataGridViewRow>().Count(row=> row.Visible) + " רשומות";
        }
        private void ClearTable()
        {
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridViewFolders_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox8_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            tableType = "fromPro";
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            tableType = "toPro";
        }

        private void button14_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void customLabel3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            selectingRecipient();
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
                    comboBox8.SelectedIndex = 10;
                }
                else
                    comboBox8.SelectedIndex = 6;
            }

            catch { }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            int res1, res2;
            bool ok1 = int.TryParse(textBox1.Text, out res1), ok2 = int.TryParse(textBox2.Text, out res2);
            if (ok2)
            {
                if (textBox1.Text.Equals(""))
                {
                    res1 = res2 + 1;
                    ok2 = true;
                }
                if (ok1)
                {
                    if (res2 < res1)
                    {
                        textBox1.Text = textBox2.Text;
                        MessageBox.Show("אין לבחור מספר שוטף סיום קטן ממספר שוטף התחלה", "בחירת שוטפים שגויה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                }
            }
        }

        private void textBox20_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox20_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            if (textBox20.Text.Equals("שם משפחה"))
            {
                strTyped = "";
                textBox20.Text = "";
            }
        }

        private void textBox20_Leave(object sender, EventArgs e)//לממוק
        {
            if (textBox20.Text.Equals(""))
                textBox20.Text = "שם משפחה";
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if(comboBox5.Text=="הכל")
            {
                foreach(DataGridViewRow row in dataGridViewFolders.Rows)
                {
                    row.Visible = true;
                }
            }
            else
            {
                foreach (DataGridViewRow row in dataGridViewFolders.Rows)
                {
                    if (row.Cells[3].Value.ToString() == comboBox5.Text)
                        row.Visible = true;
                    else
                        row.Visible = false;
                }
            }
            Cursor.Current = Cursors.Default;
            dataGridViewFolders.Select();
        }
        
        public void button7_Click(object sender, EventArgs e)// המסמכים שלי
        {
            Cursor = Cursors.WaitCursor;
            button6_Click(sender, e); // חפש מסמכים
            comboBox8.SelectedIndex = 2; // חודש
            ChangeDataGrid("users");
            textBox7.Text = PublicFuncsNvars.curUser.userCode.ToString();
            button1_Click(sender, e);
            dataGridView2.Visible = false;
            textBox6.Visible = false;
            label7.Visible = false;
            button14.Visible = false;
            //comm.Parameters.AddWithValue("@senderCode", PublicFuncsNvars.curUser.userCode);
            //comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.curUser.userCode.ToString());
            Cursor = Cursors.Default;
        }



        private void button8_Click(object sender, EventArgs e)//למחוק
        {
            if (dataGridViewDocs.Rows.Count == 15)
            {
                dataGridViewDocs.Rows.Clear();
                dataGridViewDocs.Refresh();
                DialogResult res = System.Windows.Forms.DialogResult.Yes;
                if (res == DialogResult.Yes)
                {
                    Cursor = Cursors.WaitCursor;
                    //int temp = lastDocsViewed.Count > 0 ? lastDocsViewed.Last() : -1;
                    multiplier++;
                    if (!displayNext15Docs())
                    {
                        multiplier--;
                        MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    comboBox6.SelectedIndexChanged -= comboBox6_SelectedIndexChanged;
                    comboBox6.SelectedIndex = multiplier;
                    comboBox6.SelectedIndexChanged += comboBox6_SelectedIndexChanged;
                    Cursor = Cursors.Default;
                }
            }
            else
                MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void button9_Click(object sender, EventArgs e)//למחוק
        {
            if (multiplier == 0)
            {
                MessageBox.Show("אין מסמכים קודמים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }
            dataGridViewDocs.Rows.Clear();
            dataGridViewDocs.Refresh();
            DialogResult res = System.Windows.Forms.DialogResult.Yes;
            if (res == DialogResult.Yes)
            {
                Cursor = Cursors.WaitCursor;
                multiplier--;
                if (!displayNext15Docs())
                {
                    MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
                comboBox6.SelectedIndexChanged -= comboBox6_SelectedIndexChanged;
                comboBox6.SelectedIndex = multiplier;
                comboBox6.SelectedIndexChanged += comboBox6_SelectedIndexChanged;
                Cursor = Cursors.Default;
            }
        }

        private void button12_Click(object sender, EventArgs e)//למחוק
        {
            if (dataGridViewDocs.Rows.Count == 15)
            {
                dataGridViewDocs.Rows.Clear();
                dataGridViewDocs.Refresh();
                DialogResult res = System.Windows.Forms.DialogResult.Yes;
                if (res == DialogResult.Yes)
                {
                    Cursor = Cursors.WaitCursor;
                    //int temp = lastDocsViewed.Count > 0 ? lastDocsViewed.Last() : -1;
                    int multiplierTemp = multiplier;
                    multiplier = searchResults.Count / 15;
                    if (!displayNext15Docs())
                    {
                        multiplier = multiplierTemp;
                        MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    comboBox6.SelectedIndexChanged -= comboBox6_SelectedIndexChanged;
                    comboBox6.SelectedIndex = multiplier;
                    comboBox6.SelectedIndexChanged += comboBox6_SelectedIndexChanged;
                    Cursor = Cursors.Default;
                }
            }
            else
                MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void button13_Click(object sender, EventArgs e)//למחוק
        {
            if (multiplier == 0)
            {
                MessageBox.Show("אין מסמכים קודמים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }
            dataGridViewDocs.Rows.Clear();
            dataGridViewDocs.Refresh();
            DialogResult res = System.Windows.Forms.DialogResult.Yes;
            if (res == DialogResult.Yes)
            {
                Cursor = Cursors.WaitCursor;
                multiplier = 0;
                if (!displayNext15Docs())
                {
                    MessageBox.Show("לא נמצאו עוד מסמכים התואמים את נתוני החיפוש", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
                comboBox6.SelectedIndexChanged -= comboBox6_SelectedIndexChanged;
                comboBox6.SelectedIndex = multiplier;
                comboBox6.SelectedIndexChanged += comboBox6_SelectedIndexChanged;
                Cursor = Cursors.Default;
            }
        }

        private void dataGridViewDocs_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                DataGridViewColumn newColumn = dataGridViewDocs.Columns[e.ColumnIndex];
                DataGridViewColumn oldColumn = dataGridViewDocs.SortedColumn;
                currentlySortedColumn = newColumn;

                string sortColumn = "shotef_mismach";
                switch (newColumn.Name)
                {
                    case "docIdColumn":
                        sortColumn = "shotef_mismach";
                        break;
                    case "docSubjectColumn":
                        sortColumn = "hanadon";
                        break;
                    case "docSenderColumn":
                        sortColumn = "teur_tafkid_sholeah";
                        break;
                    case "docReferencesColumn":
                        sortColumn = "simuchin";
                        break;
                    case "docDateColumn":
                        sortColumn = "tarich_hamichtav";
                        break;
                }

                if (comboBox6.Visible)
                {
                    if (oldColumn != null)
                    {
                        if (oldColumn == newColumn)
                        {
                            newColumn.HeaderCell.SortGlyphDirection = newColumn.HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Ascending ?
                                System.Windows.Forms.SortOrder.Descending : System.Windows.Forms.SortOrder.Ascending;
                            searchResults.Reverse();
                            multiplier = 0;
                            documents.Clear();
                            dataGridViewDocs.Rows.Clear();
                            dataGridViewDocs.Refresh();
                            displayNext15Docs();
                        }
                        else
                        {
                            oldColumn.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
                            newColumn.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Ascending;
                            multiplier = 0;
                            searchResults.Clear();
                            documents.Clear();
                            getDocuments(sortColumn, true);
                        }
                    }
                    else
                    {
                        newColumn.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Ascending;
                        multiplier = 0;
                        searchResults.Clear();
                        documents.Clear();
                        getDocuments(sortColumn, true);
                    }
                }
                else
                {
                    if (oldColumn != null && oldColumn!=newColumn)
                        oldColumn.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
                    //newColumn.HeaderCell.SortGlyphDirection = newColumn.HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Ascending ?
                      //  System.Windows.Forms.SortOrder.Descending : System.Windows.Forms.SortOrder.Ascending;//AW 28.05.2024 
                    dataGridViewDocs.Sort(newColumn, newColumn.HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Ascending ?
                        ListSortDirection.Descending : ListSortDirection.Ascending);
                }
            }
            Cursor.Current = Cursors.Default;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridViewDocs.Rows.Clear();
            dataGridViewDocs.Refresh();
            multiplier = (int)comboBox6.Items[comboBox6.SelectedIndex] - 1;
            displayNext15Docs();
        }


        private void ChangeDataGrid(string tableType)
        {
            if (tableType == "users")
            {
                textBox6.Text = "";
                if(dataGridView2.Columns[0].HeaderText != "משתמש")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "משתמש");
                    dataGridView2.Columns.Add("Col1", "שם פרטי");
                    dataGridView2.Columns.Add("Col3", "שם משפחה");
                    dataGridView2.Columns.Add("Col4", "תפקיד");
                    dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;/
                    //users = PublicFuncsNvars.users;
                    foreach (User u in users)
                        dataGridView2.Rows.Add(u.userCode, u.firstName, u.lastName, u.job, u.userCode + ";" + u.firstName + ";" + u.lastName + ";" + u.job);
                    dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                textBox6.Visible = true;
                button14.Visible = true;
                button14.BringToFront();
            }
            else if (tableType == "folders")
            {
                textBox6.Text = string.Empty;
                if(dataGridView2.Columns[1].HeaderText != "שם מקוצר")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "תיק");
                    dataGridView2.Columns.Add("Col2", "שם מקוצר");
                    dataGridView2.Columns.Add("Col3", "שם תיק");
                    dataGridView2.Columns.Add("Col4", "ענף");
                    dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView2.Columns[0].DataPropertyName = "id";
                    dataGridView2.Columns[1].DataPropertyName = "shortDescription";
                    dataGridView2.Columns[2].DataPropertyName = "description";
                    dataGridView2.Columns[3].DataPropertyName = "branch";
                    //dataGridView2.Rows.Clear();
                    //dataGridView2.Columns[0].HeaderText = "תיק";
                    //dataGridView2.Columns[1].HeaderText = "שם מקוצר";
                    //dataGridView2.Columns[2].HeaderText = "שם תיק";
                    //dataGridView2.Columns[3].HeaderText = "ענף";
                    //dataGridView2.Columns[2].Visible = true;
                    //dataGridView2.Columns[3].Visible = true;
                    //dataGridView2.Columns[4].Visible = false;
                    //dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                    directories = PublicFuncsNvars.folders;
                    dataGridView2.DataSource = directories.Select(item => new
                    {
                        item.id,
                        item.shortDescription,
                        item.description,
                        branch = PublicFuncsNvars.getBranchString(item.branch)
                    }).ToList();
                    //foreach (Folder d in directories)
                       // dataGridView2.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch));
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                textBox6.Visible = true;
                label7.Visible = true;
                button14.Visible = true;
                button14.BringToFront();
            }
            else if (tableType == "projects")
            {
                textBox6.Text = "";
                if(dataGridView2.Columns[1].HeaderText != "שם פרוייקט")
                {
                    ClearTable();
                    dataGridView2.Columns.Add("Col1", "מספר");
                    dataGridView2.Columns.Add("Col2", "שם פרויטק");
                    dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    projects = PublicFuncsNvars.projects;
                    foreach (KeyValuePair<int, string> p in projects)
                        dataGridView2.Rows.Add(p.Key.ToString(), p.Value, "", "", p.Key.ToString() + ";" + p.Value);
                    dataGridView2.Refresh();
                    dataGridView2.Visible = true;
                }
                dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                textBox6.Visible = true;
                label7.Visible = true;
                button14.Visible = true;
                button14.BringToFront();
            }
        }

        private void dataGridViewTable_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
            if (tableType=="send")
            {
                string value = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox7.TextChanged -= textBox7_TextChanged;
                textBox7.Text = value;
                textBox7.TextChanged += textBox7_TextChanged;

                textBox8.TextChanged -= textBox8_TextChanged;
                textBox8.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                textBox8.TextChanged += textBox8_TextChanged;

                /*textBox3.TextChanged -= textBox3_TextChanged;
                textBox3.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox3.TextChanged += textBox3_TextChanged;

                textBox20.TextChanged -= textBox20_TextChanged;
                textBox20.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
                textBox20.TextChanged += textBox20_TextChanged;*/
                
            }
            else if (tableType == "ref")
            {
                string value = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox10.TextChanged -= textBox10_TextChanged;
                textBox10.Text = value;
                textBox10.TextChanged += textBox10_TextChanged;

                textBox9.TextChanged -= textBox9_TextChanged;
                textBox9.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[3].Value.ToString();
                textBox9.TextChanged += textBox9_TextChanged;
            }
            else if (tableType == "folders")
            {
                textBox12.TextChanged -= textBox12_TextChanged;
                textBox12.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
                textBox12.TextChanged += textBox12_TextChanged;
                textBox11.TextChanged -= textBox11_TextChanged;
                textBox11.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[1].Value.ToString();
                textBox11.TextChanged += textBox11_TextChanged;
                //textBox4.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[2].Value.ToString();
            }
            else if (tableType == "fromPro")
                textBox14.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
            else if (tableType == "toPro")
                textBox13.Text = dataGridView2.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }
    }
}
