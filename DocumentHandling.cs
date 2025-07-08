using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.Globalization;
using Word = Microsoft.Office.Interop.Word;
using System.Web.Services;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Timers;

namespace DocumentsModule
{
    public partial class DocumentHandling : Form
    {
        Document doc;//the current document
        private List<User> users;//list of users(derived from PublicFuncsNvars)
        private List<Folder> folders;//list of folders(derived from PublicFuncsNvars)
        bool okRef = false,//true if recipients adjustments are being done
            okBro = false;//true if browsers adjustments are being done
        string strTyped = "";//keeps a substring by which the user can choose a value in a datagridview
        private int row4;//the current row in datagridview4
        int rowAutho;
        private Dictionary<KeyValuePair<short, short>, RecipientList> recipientsLists;//list of recipients lists(derived from PublicFuncsNvars)
        private List<Recipient> interDist;
        private string docExt;
        private bool isOriginalNull = false;
        //lists to contain groups of controls for easier visibility changes
        List<Control> newBroControls = new List<Control>();
        List<Control> newFolderControls = new List<Control>();
        List<Control> scanControls = new List<Control>();
        List<Control> AttsControls = new List<Control>();
        List<Control> publishControls=new List<Control>();
        List<Control> addExistingDocControls = new List<Control>();
        bool isAllowedToEdit;
        private DataGridViewRow recipientToMoveRow;
        private int rowIndexFromMouseDown;
        private bool hasBeenUpdated = true;
        private bool hasBeenUpdatedForALotOfRecs = true;
        private int shotef;
        Microsoft.Office.Interop.Outlook.Application app = null;
        FileInfo fi = null;
        string oldSubject;
        int oldUserId;
        string oldUserRole, oldUserName, oldClassification, oldClasfication;
        System.Timers.Timer _timer;
        bool TheDocCaChange = true;
        bool OpenedForEdit;
        public DocumentHandling(int id)
        {
            shotef = id;
            PublicFuncsNvars.dhFormsOpen.Add(id);
            this.doc = PublicFuncsNvars.createDoc(id);
            docExt = PublicFuncsNvars.getDocExt(id);
            isAllowedToEdit = PublicFuncsNvars.isAuthorizedUser(doc.getSenderRole(), PublicFuncsNvars.curUser) || doc.isCurUserAuthorizedToEdit() ||
                PublicFuncsNvars.curUser.userCode == doc.getCreatorCode();
            
            InitializeComponent();
            dataGridViewVers.CellDoubleClick += dataGridViewVers_CellDoubleClick;
        }

        private void DocumentHandling_Load(object sender, EventArgs e)
        {
            OpenedForEdit = false;
            UpdateTxtVer(shotef, 'L');
            //   dataGridViewRecipients.MouseClick+= dataGridViewRecipients_MouseClick;
            customLabel8.Visible = false;
            dataGridViewVers.Visible = false;
            DocumentsMenu.PathTemplate(this.viewDocButton, 40);
            DocumentsMenu.PathTemplate(this.BtnVer, 40);
            DocumentsMenu.PathTemplate(this.editDocButton, 40);
            DocumentsMenu.PathTemplate(this.recipientsButton, 40);
            DocumentsMenu.PathTemplate(this.foldersButton, 40);
            DocumentsMenu.PathTemplate(this.addAttButton, 40);
            DocumentsMenu.PathTemplate(this.browsersButton, 40);
            DocumentsMenu.PathTemplate(this.quickPrintButton, 40);
            DocumentsMenu.PathTemplate(this.publishDocButton, 40);
            DocumentsMenu.PathTemplate(this.button36, 40);
            DocumentsMenu.PathTemplate(this.btn_Share, 40);
            DocumentsMenu.PathTemplate(this.button1, 30);
            this.button2.FlatStyle = FlatStyle.Flat;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button18.FlatStyle = FlatStyle.Flat;
            this.button18.FlatAppearance.BorderSize = 0;
            DocumentsMenu.PathTemplate(this.button3, 30);
            DocumentsMenu.PathTemplate(this.button4, 30);
            DocumentsMenu.PathTemplate(this.button5, 30);
            DocumentsMenu.PathTemplate(this.button6, 30);
            DocumentsMenu.PathTemplate(this.button7, 30);
           // DocumentsMenu.PathTemplate(this.button8, 30);
            DocumentsMenu.PathTemplate(this.button9, 30);
            this.button10.FlatStyle = FlatStyle.Flat;
            this.button10.FlatAppearance.BorderSize = 0;
            DocumentsMenu.PathTemplate(this.button11, 30);
            DocumentsMenu.PathTemplate(this.button12, 30);
            //this.button12.FlatStyle = FlatStyle.Flat;
            //this.button12.FlatAppearance.BorderSize = 0;
            this.button13.FlatStyle = FlatStyle.Flat;
            this.button13.FlatAppearance.BorderSize = 0;
         //  DocumentsMenu.PathTemplate(this.button14, 30); // כפתור סריקה - נמחק
            DocumentsMenu.PathTemplate(this.button15, 30);
            this.button16.FlatStyle = FlatStyle.Flat;
            this.button16.FlatAppearance.BorderSize = 0;
            this.button17.FlatStyle = FlatStyle.Flat;
            this.button17.FlatAppearance.BorderSize = 0;
            DocumentsMenu.PathTemplate(this.button19, 30);
            this.button37.FlatStyle = FlatStyle.Flat;
            this.button37.FlatAppearance.BorderSize = 0;
            DocumentsMenu.PathTemplate(this.button32, 30);
            DocumentsMenu.PathTemplate(this.button31, 30);
            DocumentsMenu.PathTemplate(this.button30, 30);
            DocumentsMenu.PathTemplate(this.button29, 30);
            DocumentsMenu.PathTemplate(this.button28, 30);
            DocumentsMenu.PathTemplate(this.button27, 30);
            DocumentsMenu.PathTemplate(this.button26, 30);
            this.button25.FlatStyle = FlatStyle.Flat;
            this.button25.FlatAppearance.BorderSize = 0;
            DocumentsMenu.PathTemplate(this.button24, 30);
            DocumentsMenu.PathTemplate(this.button23, 30);
            DocumentsMenu.PathTemplate(this.button22, 30);
            DocumentsMenu.PathTemplate(this.button21, 30);
            DocumentsMenu.PathTemplate(this.button20, 20);
            DocumentsMenu.PathTemplate(this.button38, 20);


            dataGridViewUsers.Rows.Clear();
            dataGridViewFolders.Rows.Clear();
            dataGridViewRecipients.Rows.Clear();
            dataGridViewAuthorizations.Rows.Clear();
            dataGridViewInterDist.Rows.Clear();
            dataGridViewAtts.Rows.Clear();
            dataGridViewRecipientLists.Rows.Clear();

            comboBox1.DataSource = PublicFuncsNvars.sivug_by_reshet();
            
            this.Icon = Global.AppIcon;

            _timer = new System.Timers.Timer(Global.SecondToCloseDocForm);//כל 15 דקות לשמור את במסמך עם פתחו את המסמך.
            _timer.Elapsed += OnTimedEvent;
            _timer.AutoReset = true;
            _timer.Start();

            populateControlsList(ref newBroControls, label24, dataGridViewUsers, button6, button5);
            populateControlsList(ref newFolderControls, label14, customLabel6, comboBox5, dataGridViewFolders, button11, button3, textBox16,button40);
            populateControlsList(ref scanControls, button20, textBox13, customLabel4);
            populateControlsList(ref addExistingDocControls, button9, textBox13, customLabel5);
            populateControlsList(ref AttsControls, button19, button21, dataGridViewAtts, label9);
            populateControlsList(ref publishControls, dataGridViewToSend, button28, button29, label18);

            dataGridViewToSend.Columns[0].ReadOnly = true;
            dataGridViewToSend.Columns[1].ReadOnly = true;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            ToolStripMenuItem[] mRef = new ToolStripMenuItem[1];
            ToolStripMenuItem removeRecipient = new ToolStripMenuItem("הסר מכותב");
            removeRecipient.Click += removeRecipientFromDoc;
            mRef[0] = removeRecipient;
            dataGridViewRecipients.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewRecipients.ContextMenuStrip.Items.AddRange(mRef);
            if(!PublicFuncsNvars.isNormalDoc(doc.getID()))
            {
                recipientsButton.Enabled = false;
                foldersButton.Enabled = false;
                panel1.Visible = true;
                button10.Visible = false;
                button16.Visible = false;
                button17.Visible = false;
                button25.Visible = false;
                button2.Visible = false;
                button13.Visible = false;
                button18.Visible = false;
                button37.Visible = false;
                dataGridViewRecipients.Columns["recipientIFAColumn"].ReadOnly = true;
                dataGridViewRecipients.ContextMenuStrip.Items.Clear();
                editDocButton.Enabled = false;
                button33.Enabled = false;
                button34.Enabled = false;
                button35.Enabled = false;
                dataGridViewRecipients.MouseMove -= dataGridViewRecipients_MouseMove;
                dataGridViewRecipients.MouseDown -= dataGridViewRecipients_MouseDown;
                dataGridViewRecipients.DragOver -= dataGridViewRecipients_DragOver;
                dataGridViewRecipients.DragDrop -= dataGridViewRecipients_DragDrop;
            }

            
            if (doc.getCreationDate() < DateTime.Parse(Global.ReadOnly))
            {
                editDocButton.Enabled = false;
                button33.Enabled = false;
                button34.Enabled = false;
                //button35.Enabled = false;
                button1.Enabled = false;
                TheDocCaChange = false;
                
                // ASAF 5.9.24
                button2.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button7.Enabled = false;
                button10.Enabled = false;
                button13.Enabled = false;
                button15.Enabled = false;
                button16.Enabled = false;
                button17.Enabled = false;
                button22.Enabled = false;
                button25.Enabled = false;
                button26.Enabled = false;
                button27.Enabled = false;
                button31.Enabled = false;
                button32.Enabled = false;
                button35.Enabled = false;
                button37.Enabled = false;

            }

            ToolStripMenuItem[] mBro = new ToolStripMenuItem[1];
            ToolStripMenuItem removeAuthorization = new ToolStripMenuItem("הסר הרשאה");
            removeAuthorization.Click += removeAuthorizationFromDoc;
            mBro[0] = removeAuthorization;
            dataGridViewAuthorizations.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewAuthorizations.ContextMenuStrip.Items.AddRange(mBro);


            /* classification*/
            var temp_class = doc.getClassification();
            comboBox1.SelectedIndex = (int)temp_class-1;

            this.Text = "טיפול במסמך " + docExt?.ToUpper() + " שוטף מספר " + doc.getID().ToString();
            textBox1.Text = doc.getSubject();
            oldSubject = textBox1.Text;
            textBox2.Text = doc.getSenderRole().ToString();
            oldUserId = Convert.ToInt32(doc.getSenderRole());
            textBox3.Text = PublicFuncsNvars.getUserNameByUserCode(doc.getSenderRole());
            oldUserRole = textBox3.Text;
            textBox4.Text = doc.getSenderName();
            oldUserName = textBox4.Text;
           
            List<Folder> filedFolders = doc.getFolders();
            if (filedFolders.Count > 0)
            {
                Folder dir = filedFolders[0];
                textBox12.Text = dir.id.ToString();
                textBox11.Text = dir.shortDescription;
                textBox5.Text = dir.description;
                textBox6.Text = doc.getNumInFolder(dir.id).ToString();
            }
            else
            {
                MessageBox.Show("למסמך זה אין תיק ולכן תוייק לתיק זמני, נא לתייק לתיק קבוע.",
                    "מסמך ללא תיק", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

               // SqlConnection conn2 = new SqlConnection(Global.ConStr);
               
                using (SqlConnection conn2 = new SqlConnection(Global.ConStr))
                {
                    conn2.Open();
                    SqlCommand comm2 = new SqlCommand("INSERT INTO dbo.tiukim (kod_marechet, shotef_klali, mispar_nose, is_rashi, mispar_in_tik)"
                                   + Environment.NewLine +
                                   " VALUES (2,@id,16163,1,0)", conn2); // תיק ללא תיק.
                    comm2.Parameters.AddWithValue("@id", doc.getID());
                    comm2.ExecuteNonQuery();
                }

                doc = PublicFuncsNvars.createDoc(doc.getID());
                filedFolders = doc.getFolders();
                Folder dir = filedFolders[0];
                textBox12.Text = dir.id.ToString();
                textBox11.Text = dir.shortDescription;
                textBox5.Text = dir.description;
                textBox6.Text = doc.getNumInFolder(dir.id).ToString();
            }
            KeyValuePair<int, string> p = PublicFuncsNvars.getProjectById(doc.getProject());
            if (p.Key != -1)
            {
                textBox9.Text = p.Key.ToString();
                textBox10.Text = p.Value;
            }
            dateTimePicker1.Value = doc.getCreationDate();
            DateTime entryDate = doc.getEntryDate();
            if (entryDate > DateTime.MinValue)
            {
                dateTimePicker2.Value = entryDate;
            }
            else dateTimePicker2.Value = new DateTime(1900, 1, 1);
            comboBox1.Text = PublicFuncsNvars.getClassificationByEnum(doc.getClassification());
            oldClassification = comboBox1.Text;
            textBox7.Text = doc.getRefferences();
            checkBox1.Checked = doc.getIsPublished();
            checkBox3.Checked = doc.getIsRagish();
            if (checkBox1.Checked)
            {
                textBox1.ReadOnly = true;
                dateTimePicker3.Value = doc.getPublishDate();
                dateTimePicker3.Enabled = true;
            }
            else
            {
                dateTimePicker3.Text = "";
                dateTimePicker3.Enabled = false;
            }
            textBox8.Text = doc.getNotes();

            checkBox2.Checked = doc.getIsActive();

            users = PublicFuncsNvars.users.Where(x=>x.isActive).ToList();
            foreach (User u in users)
                dataGridViewUsers.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
            dataGridViewUsers.Sort(dataGridViewUsers.Columns[0], ListSortDirection.Ascending);
            folders = PublicFuncsNvars.folders.Where(x=>x.isActive).ToList();
            foreach (Folder d in folders)
                dataGridViewFolders.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch));

            foreach (Recipient r in doc.getRecipients())
            {
                string name = PublicFuncsNvars.getUserNameByUserCode(r.getId());
                if (name == null)
                    name = r.getRole();
                int rowIndex=dataGridViewRecipients.Rows.Add(r.getNID(), r.getId(), name, r.getRole());
                dataGridViewRecipients.Rows[rowIndex].Cells[4].Value = r.getIFA() ? "לפעולה" : "לידיעה";
                if (dataGridViewRecipients.Rows[rowIndex].Cells[3].Value.ToString().StartsWith("ת.פ."))
                {
                    dataGridViewRecipients.Rows[rowIndex].Cells[3].ReadOnly = true;
                }
            }
            dataGridViewRecipients.Sort(dataGridViewRecipients.Columns[0], ListSortDirection.Ascending);

            foreach (KeyValuePair<int,bool> au in doc.getAuthorizedUsers())
            {
                User u=PublicFuncsNvars.getUserByCode(au.Key);
                string name = u.getFullName(), role = u.job;
                int rowIndex = dataGridViewAuthorizations.Rows.Add(au.Key, name, role);
                dataGridViewAuthorizations.Rows[rowIndex].Cells[3].Value = au.Value ? "לעריכה" : "לצפיה";
            }

            ToolStripMenuItem[] mfolder = new ToolStripMenuItem[1];
            ToolStripMenuItem removefolder = new ToolStripMenuItem("הסר תיוק");
            removefolder.Click += removefolderFromDoc;
            mfolder[0] = removefolder;
            dataGridViewFiledFolders.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewFiledFolders.ContextMenuStrip.Items.AddRange(mfolder);

            foreach (Folder d in filedFolders)
                dataGridViewFiledFolders.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch), d.isMain);

            recipientsLists = PublicFuncsNvars.recipientsLists;
            foreach (KeyValuePair<KeyValuePair<short, short>, RecipientList> rl in recipientsLists)
            {
                RecipientList temp = rl.Value;
                if ((temp.level == RecipientListsLevel.personal && temp.owner == PublicFuncsNvars.curUser.userCode) ||
                    (temp.level == RecipientListsLevel.branch && temp.branch == PublicFuncsNvars.curUser.branch) ||
                    temp.level == RecipientListsLevel.unit)
                {
                    dataGridViewRecipientLists.Rows.Add(rl.Key.Key, rl.Key.Value, temp.name, temp.getLevelString(), temp.getOwnerName());
                }
            }

            interDist = PublicFuncsNvars.interDist;
            foreach (Recipient r in interDist)
                dataGridViewInterDist.Rows.Add(false, r.getId(), r.getRole(), r.getEmail());

            string mainFile_extention = string.Empty;

            using (SqlConnection conn1 = new SqlConnection(Global.ConStr))
            {
                SqlCommand comm1 = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn1);
                comm1.Parameters.AddWithValue("@id", doc.getID());
                conn1.Open();
                SqlDataReader sdr1 = comm1.ExecuteReader();
                sdr1.Read();
                mainFile_extention = sdr1.GetString(1);

                isOriginalNull = sdr1.IsDBNull(0);
            }

               // SqlConnection conn1 = new SqlConnection(Global.ConStr);
            
            /*if (mainFile_extention.ToLower() == "docx" || mainFile_extention.ToLower() == "doc")
                button38.Visible = true;*///תצוגה מקדימה
          //  conn1.Close();


            if (!isOriginalNull)
                dataGridViewToSend.Rows.Add(doc.getID(), doc.getSubject(), doc.getCreationDate(),mainFile_extention, false);

            ToolStripMenuItem[] mAtt = new ToolStripMenuItem[1];
            ToolStripMenuItem removeAttachment = new ToolStripMenuItem("הסר נספח");
            removeAttachment.Click += removeAttachmentFromDoc;
            mAtt[0] = removeAttachment;
            dataGridViewAtts.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewAtts.ContextMenuStrip.Items.AddRange(mAtt);


            using (SqlConnection conn = new SqlConnection(Global.ConStr))
            {
                conn.Open();
                SqlCommand comm = new SqlCommand("SELECT shotef_nisph, prtim, tarich,file_extension FROM dbo.docnisp WHERE shotef_mchtv=@id and datalength(file_data) >= 1 ", conn); // AND datalength(file_data)>0 - causes long time execution
                comm.Parameters.AddWithValue("@id", doc.getID());
                //  comm.CommandTimeout = 100;

                SqlDataReader sdr = null;
                try
                {
                    sdr = comm.ExecuteReader(CommandBehavior.SequentialAccess);
                    while (sdr.Read())
                    {
                        int id = sdr.GetInt32(0);
                        string fileName = sdr.GetString(1).Trim();
                        string sDate = sdr.GetString(2).Trim();
                        DateTime date;
                        if (!DateTime.TryParseExact(sDate, "yyyyMMdd", new CultureInfo("he-IL"), DateTimeStyles.None, out date))
                            date = doc.getCreationDate();

                        
                        string file_extension = sdr.GetString(3).Trim();

                        dataGridViewAtts.Rows.Add(id, fileName, date);
                        dataGridViewToSend.Rows.Add(id, fileName, date, file_extension, false, false);
                        bool isWord = false;
                        bool isExcel = false;
                        bool isPPT = false;
                        List<string> WordTypes = new List<string> { "doc", "docx", "dot", "dotx" };
                        List<string> ExcelTypes = new List<string> { "xl", "xlsx", "xlsm", "xlsb", "xlam", "xls" };
                        List<string> PptTypes = new List<string> { "ppt", "pptm", "pptx", "ptx" };
                        if (WordTypes.Any(s => file_extension.ToLower().Equals(s))) isWord = true;
                        if (ExcelTypes.Any(s => file_extension.ToLower().Equals(s))) isExcel = true;
                        if (PptTypes.Any(s => file_extension.ToLower().Equals(s))) isPPT = true;


                        if (!(isWord || isExcel || isPPT))
                        {
                            dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 1].Cells["check_PDF"].ReadOnly = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + '\n' + "אירעה שגיאה! יש לפנות למחשוב", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (dataGridViewToSend.Rows.Count > 0)
                    dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 1].Cells["fileToSendColumn"].Value = true;

            }
             //   SqlConnection conn = new SqlConnection(Global.ConStr);

            dataGridViewRecipients.CellValueChanged += dataGridViewRecipients_CellValueChanged;

        }
        
        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            this.Invoke(new Action(() =>
            {
                Cursor.Current = Cursors.WaitCursor;
                int id = doc.getID();
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT  file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                comm.Parameters.AddWithValue("@id", id);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();

                if (sdr.Read())
                {
                    string fileExt = sdr.GetString(0).Trim();
                    if (!fileExt.ToLower().Contains("doc") && !fileExt.ToLower().Contains("docx"))
                    {
                        Cursor.Current = Cursors.Default;
                        return;
                    }

                    string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                    if (File.Exists(filePath))//רק אם המסמך קיים ולא פתוח הוא ישמור.
                    {
                        try
                        {
                            byte[] fileData = File.ReadAllBytes(filePath);
                            Word.Document document = null;
                           string docText = PublicFuncsNvars.docToTxt(document, filePath);
                            SaveVersion(id);
                            PublicFuncsNvars.saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, docText);
                            
                            try
                            {
                                File.Delete(filePath);
                            }
                            catch { }
                        }
                        catch//מסמך פתוח
                        {
                            Cursor.Current = Cursors.Default;
                            return;
                        }
                    }
                }

                Cursor.Current = Cursors.Default;
            }));
            /*
            int id = doc.getID();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT  file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();

            if (sdr.Read())
            {
                string fileExt = sdr.GetString(0).Trim();
                if (!fileExt.ToLower().Contains("doc") && !fileExt.ToLower().Contains("docx"))
                {
                    Cursor.Current = Cursors.Default;
                    return;
                }

                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                if (File.Exists(filePath))//רק אם המסמך קיים ולא פתוח הוא ישמור.
                {
                    try
                    {
                        byte[] fileData = File.ReadAllBytes(filePath);
                        Word.Document document = null;
                        Text = PublicFuncsNvars.docToTxt(document, filePath);
                        PublicFuncsNvars.saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, Text);
                        SaveVersion(id);
                        try
                        {
                            File.Delete(filePath);
                        }
                        catch { }
                    }
                    catch//מסמך פתוח
                    {
                        Cursor.Current = Cursors.Default;
                        return;
                    }
                }
            }*/

        }
        private void savedoc()
        {
        }

        private void removefolderFromDoc(object sender, EventArgs e)
        {
            if (dataGridViewFiledFolders.SelectedCells.Count > 0 && dataGridViewFiledFolders.SelectedCells[0].OwningRow.Index >= 0)
            {
                DataGridViewRow row = dataGridViewFiledFolders.SelectedCells[0].OwningRow;
                if ((bool)row.Cells["filedFolderIsMainColumn"].Value)
                    MessageBox.Show("לא ניתן להסיר תיק ראשי.", "הסרת תיוקים", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                else
                {
                    doc.removeFolder(int.Parse(row.Cells[0].Value.ToString()));
                    dataGridViewFiledFolders.Rows.Remove(row);
                }
            }
        }

        private void removeAttachmentFromDoc(object sender, EventArgs e)
        {
            if (dataGridViewAtts.SelectedRows.Count > 0 && dataGridViewAtts.SelectedRows[0].Index >= 0)
            {
                DataGridViewRow row = dataGridViewAtts.SelectedRows[0];
                doc.removeAttachment(int.Parse(row.Cells[0].Value.ToString()));
                dataGridViewAtts.Rows.Remove(row);
            }
        }

        private void removeAuthorizationFromDoc(object sender, EventArgs e)
        {
            if (rowAutho >= 0)
            {
                DataGridViewRow row = dataGridViewAuthorizations.Rows[rowAutho];
                doc.removeAuthorization(int.Parse(row.Cells[0].Value.ToString()));
                dataGridViewAuthorizations.Rows.Remove(row);
            }
        }

        private void populateControlsList(ref List<Control> controlsList, params Control[] controls)
        {
            controlsList = controls.ToList();
        }

        private void removeRecipientFromDoc(object sender, EventArgs e)
        {
            if (row4 >= 0)
            {
                DataGridViewRow row=dataGridViewRecipients.Rows[row4];
                doc.removeRecipient(short.Parse(row.Cells[0].Value.ToString()));
                dataGridViewRecipients.Rows.Remove(row);
                hasBeenUpdated = false;
            }
        }

        private void dataGridView4_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                row4 = e.RowIndex;
                dataGridViewRecipients.Rows[row4].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void viewDocButton_Click(object sender, EventArgs e)
        {            
            PublicFuncsNvars.changeControlsVisiblity(false, publishControls);

            try
            {
                if (!hasBeenUpdated)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    PublicFuncsNvars.updateRecipientsInWordDoc(doc, hasBeenUpdated);//לשנות
                    hasBeenUpdated = true;
                    hasBeenUpdatedForALotOfRecs = true;
                    Cursor.Current = Cursors.Default;
                }
                ThreadPool.QueueUserWorkItem(viewDoc, doc.getID());
            }

            catch // לא אמור להכנס לפה , כי לא יכול להיות שאין מסמך בלי קובץ מאחוריו.
            {
                if (dataGridViewAtts.Rows.Count > 0)
                {
                    int attId = doc.getSignedAtt();
                    if (doc.getSignedAtt() == -1)
                        dataGridView1_CellDoubleClick(dataGridViewAtts, new DataGridViewCellEventArgs(0, dataGridViewAtts.Rows.Count - 1));
                    else
                        ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(doc.getID(), attId));
                }
            }


            //if (dataGridViewAtts.Rows.Count > 0)
            //{
            //    int attId = doc.getSignedAtt();
            //    if (doc.getSignedAtt() == -1)
            //        dataGridView1_CellDoubleClick(dataGridViewAtts, new DataGridViewCellEventArgs(0, dataGridViewAtts.Rows.Count - 1));
            //    else
            //        ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(doc.getID(), attId));
            //}
            //else
            //{
            //    if (!hasBeenUpdated)
            //    {
            //        Cursor.Current = Cursors.WaitCursor;
            //        PublicFuncsNvars.updateRecipientsInWordDoc(doc, hasBeenUpdated);//לשנות
            //        hasBeenUpdated = true;
            //        hasBeenUpdatedForALotOfRecs = true;
            //        Cursor.Current = Cursors.Default;
            //    }
            //    ThreadPool.QueueUserWorkItem(viewDoc, doc.getID());

            //}
        }

        private void viewDoc(object idObj)
        {
            
            Cursor.Current = Cursors.WaitCursor;
            (new PublicFuncsNvars()).viewDoc((int)idObj);
                Cursor.Current = Cursors.Default;
           
        }

        private void editDocButton_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            OpenedForEdit = true;
            dataGridViewVers.Visible = false;
            customLabel8.Visible = false;
            if (isAllowedToEdit)
            {
                if ((doc.isRegularDoc()) || (PublicFuncsNvars.isRegularDocument))
                {
                    panel1.Visible = false;
                    PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
                    if (!hasBeenUpdated)//!hasBeenUpdated
                    {
                        if (PublicFuncsNvars.updateRecipientsInWordDoc(doc, hasBeenUpdated) ||
                            PublicFuncsNvars.curUser.roleType == RoleType.computers)
                        {
                            hasBeenUpdated = true;
                            hasBeenUpdatedForALotOfRecs = true;
                            Cursor.Current = Cursors.Default;
                        }
                        else
                        {
                            Cursor.Current = Cursors.Default;
                            return;
                        }//לשנות
                    }
                    PublicFuncsNvars.viewDocForEdit(doc.getID());
                    this.Cursor = Cursors.Default;
                    UpdateTxtVer(doc.getID(), 'L');
                }
                else
                {
                    MessageBox.Show("לא ניתן לערוך מסמך בפורמט זה.",
                        "עריכה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                    this.Cursor = Cursors.Default;
                    return;
                }
            }
            else
            {
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "עריכת מסמך", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                this.Cursor = Cursors.Default;
                return;

            }
        }

        private void viewDocForEdit(object idObj)
        {
             int x = (int)idObj;
            PublicFuncsNvars.viewDocForEdit(x);

            if (!MyGlobals.afterEdit)
            {
              
                try
                {
                    this.Invoke(new MethodInvoker(delegate () { this.Cursor = Cursors.Default; }));
                }

                catch { }
            }

            else

            {
                MyGlobals.afterEdit = false;
            }
            
        }

        private void publishDocButton_Click(object sender, EventArgs e)
        {
            FillPublishTable();
            if (PublicFuncsNvars.openDocs.Contains(doc.getID()))
            {
                MessageBox.Show("מסמך זה פתוח כרגע, יש לסגור אותו לפני הפצתו.",
                            "הפצה", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (isAllowedToEdit)
            {
                PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
                PublicFuncsNvars.changeControlsVisiblity(false, scanControls);
                PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
                PublicFuncsNvars.changeControlsVisiblity(true, publishControls);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "הפצה", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

        }

        private void FillPublishTable()
        {
            dataGridViewToSend.Rows.Clear();
            dataGridViewAtts.Rows.Clear();
           
            SqlConnection conn1 = new SqlConnection(Global.ConStr);
            SqlCommand comm1 = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn1);
            comm1.Parameters.AddWithValue("@id", doc.getID());
            conn1.Open();
            SqlDataReader sdr1 = comm1.ExecuteReader();
            sdr1.Read();
            string mainFile_extention = sdr1.GetString(1);

            isOriginalNull = sdr1.IsDBNull(0);


            conn1.Close();
            if (!isOriginalNull)
                dataGridViewToSend.Rows.Add(doc.getID(), doc.getSubject(), doc.getCreationDate(), mainFile_extention, false);

            ToolStripMenuItem[] mAtt = new ToolStripMenuItem[1];
            ToolStripMenuItem removeAttachment = new ToolStripMenuItem("הסר נספח");
            removeAttachment.Click += removeAttachmentFromDoc;
            mAtt[0] = removeAttachment;
            dataGridViewAtts.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            dataGridViewAtts.ContextMenuStrip.Items.AddRange(mAtt);

            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT shotef_nisph, prtim, tarich,file_extension FROM dbo.docnisp WHERE shotef_mchtv=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", doc.getID());
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                string sDate = sdr.GetString(2).Trim();
                DateTime date;
                if (!DateTime.TryParseExact(sDate, "yyyyMMdd", new CultureInfo("he-IL"), DateTimeStyles.None, out date))
                    date = doc.getCreationDate();
                string fileName = sdr.GetString(1).Trim();
                int id = sdr.GetInt32(0);
                string file_extension = sdr.GetString(3).Trim();

                dataGridViewAtts.Rows.Add(id, fileName, date);
                dataGridViewToSend.Rows.Add(id, fileName, date, file_extension, false, false);
                bool isWord = false;
                bool isExcel = false;
                bool isPPT = false;
                List<string> WordTypes = new List<string> { "doc", "docx", "dot", "dotx" };
                List<string> ExcelTypes = new List<string> { "xl", "xlsx", "xlsm", "xlsb", "xlam", "xls" };
                List<string> PptTypes = new List<string> { "ppt", "pptm", "pptx", "ptx" };
                if (WordTypes.Any(s => file_extension.ToLower().Equals(s))) isWord = true;
                if (ExcelTypes.Any(s => file_extension.ToLower().Equals(s))) isExcel = true;
                if (PptTypes.Any(s => file_extension.ToLower().Equals(s))) isPPT = true;


                if (!(isWord || isExcel || isPPT))
                {
                    dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 1].Cells["check_PDF"].ReadOnly = true;
                }
            }

            if (dataGridViewToSend.Rows.Count > 0)
                //dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 1].Cells["fileToSendColumn"].Value = true;
                dataGridViewToSend.Rows[0].Cells["fileToSendColumn"].Value = true; // 23.09.24 ASAF MOR
            conn.Close();
        }

        private void recipientsButton_Click(object sender, EventArgs e)
        {
            if(PublicFuncsNvars.openDocs.Contains(doc.getID()))
            {
                MessageBox.Show("מסמך זה פתוח כרגע, יש לסגור אותו לפני שינוי פרטיו.",
                            "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (isAllowedToEdit)
            {
                PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
                okRef = true;
                okBro = false;
                panel1.Visible = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                //panel9.Visible = false;
                button6.Visible = false;
                customLabel8.Visible = false;
                dataGridViewVers.Visible = false;
                button21_Click(sender, e);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                panel6.Visible = false;
                panel5.Visible = true;
                textBox24.Visible = true;
                button39.Visible = true;
                button39.BringToFront();
                panel3.Visible = false;
                panel4.Visible = false;
                panel8.Visible = false;
                button6.Visible = true;
                comboBox2.Visible = true;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                DialogResult res = MessageBox.Show("האם להסיר את כל המכותבים מהמסמך?", "הסרת מכותבים", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (res == DialogResult.Yes)
                {
                    panel3.Visible = false;
                    panel4.Visible = false;
                    panel5.Visible = false;
                    textBox24.Visible = false;
                    button39.Visible = false;
                    panel6.Visible = false;
                    panel8.Visible = false;
                    button6.Visible = false;
                    comboBox2.Visible = false;
                    button23_Click(sender, e);
                    doc.removeAllRecipients();
                    dataGridViewRecipients.Rows.Clear();
                    hasBeenUpdated = false;
                }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            okRef = false;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            textBox24.Visible = false;
            button39.Visible = false;
            panel6.Visible = false;
            panel8.Visible = false;
            comboBox2.Visible = false;
            button6.Visible = false;
            button23_Click(sender, e);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (okRef)
            {
                if (dataGridViewUsers.SelectedCells.Count > 0)
                {
                    List<DataGridViewRow> rowCollection = new List<DataGridViewRow>();
                    foreach (DataGridViewCell cell in dataGridViewUsers.SelectedCells)
                    {
                        if (!rowCollection.Contains(cell.OwningRow))
                            rowCollection.Add(cell.OwningRow);
                    }
                    Cursor = Cursors.WaitCursor;
                    foreach (DataGridViewRow row in rowCollection)
                    {
                        int id = int.Parse(row.Cells[0].Value.ToString());
                        if (doc.addRecipient(new Recipient(id, (short)(doc.getMaxRecipient() + (short)1), row.Cells[3].Value.ToString(),
                            comboBox2.Text == "לפעולה", true, PublicFuncsNvars.getEmailByUserCode(id))))
                        {
                            dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), id, row.Cells[1].Value.ToString() + " " + row.Cells[2].Value.ToString(),
                                                            row.Cells[3].Value.ToString(), comboBox2.Text);
                        }
                    }
                    hasBeenUpdated = false;
                    comboBox2.Visible = false;
                    panel5.Visible = false;
                    textBox24.Visible = false;
                    button39.Visible = false;
                    button6.Visible = false;
                    Cursor = Cursors.Default;
                }
            }
            else
            {
                if (dataGridViewUsers.SelectedCells.Count > 0)
                {
                    List<DataGridViewRow> rowCollection = new List<DataGridViewRow>();
                    foreach (DataGridViewCell cell in dataGridViewUsers.SelectedCells)
                    {
                        if (!rowCollection.Contains(cell.OwningRow))
                            rowCollection.Add(cell.OwningRow);
                    }
                    Cursor = Cursors.WaitCursor;
                    foreach (DataGridViewRow row in rowCollection)
                    {
                        int id = int.Parse(row.Cells[0].Value.ToString());
                        if (doc.addAuthorization(id, comboBox6.Text=="לעריכה"))
                        {
                            dataGridViewAuthorizations.Rows.Add(id, row.Cells[1].Value.ToString() + " " + row.Cells[2].Value.ToString(),
                                                            row.Cells[3].Value.ToString(), comboBox6.Text);
                        }
                    }
                    comboBox6.Visible = false;
                    panel5.Visible = false;
                    textBox24.Visible = false;
                    button39.Visible = false;
                    button6.Visible = false;
                    Cursor = Cursors.Default;
                }
            }
            strTyped = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel6.Visible = false;
            panel8.Visible = false;
            panel5.Visible = false;
            textBox24.Visible = false;
            button39.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            button6.Visible = false;
            comboBox2.Visible = false;
            comboBox6.Visible = false;
            strTyped = "";
        }

        private void DocumentHandling_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData.HasFlag(Keys.K) && ((ModifierKeys.HasFlag(Keys.Control)) || e.Modifiers.HasFlag(Keys.Control)))
            {
                if (dataGridViewRecipients.Visible == true)
                {
                    int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
                    if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
                    {
                        MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                            "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    else
                    {
                        this.BringToFront();
                        List<Tuple<string, string, string, bool>> rec = PublicFuncsNvars.getCtrlKRecipients();
                        foreach (Tuple<string, string, string, bool> r in rec)
                        {
                            if (doc.addRecipient(new Recipient(99999, (short)(doc.getMaxRecipient() + (short)1), r.Item3, r.Item4, true, r.Item2)))
                                dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), 99999, r.Item1, r.Item3, r.Item4 ? "לפעולה" : "לידיעה");
                        }
                        hasBeenUpdated = false;
                    }
                }
            }
        }

        private void dataGridView3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                
                dataGridViewUsers.CurrentCell = dataGridViewUsers[dataGridViewUsers.SelectedCells[0].ColumnIndex, dataGridViewUsers.SelectedCells[0].RowIndex - 1];
                button5_Click(sender, e);
                return;
            }
            strTyped += e.KeyChar;

            int col = dataGridViewUsers.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewUsers.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    dataGridViewUsers.ClearSelection();
                    row.Cells[col].Selected = true;
                    dataGridViewUsers.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }
        
        private void dataGridView3_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }
        
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridViewFolders_KeyPress(object sender, KeyPressEventArgs e)
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

        private void dataGridViewFolders_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridViewFolders_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }
        
        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

     /*   private void button14_Click(object sender, EventArgs e) // כפתור סריקה - נמחק
        {
            PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
            PublicFuncsNvars.changeControlsVisiblity(true, scanControls);
            textBox13.Text = textBox1.Text;
        }*/

        private void button20_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (Directory.Exists(@"C:\Button_Data\PDF"))
                {
                    string[] files = Directory.GetFiles(@"C:\Button_Data\PDF");
                    bool isex = false;
                    foreach (string f in files)
                    {
                        try
                        {
                            File.Delete(f);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(@"נכשל במחיקת קבצים מתיקיה 'C:\Button_Data\PDF'\nיש למחוק תיקיה זו ידנית", "מערכת ניהול מסמכים", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            isex = true;
                        }
                    }
                    try
                    {
                        Directory.Delete(@"C:\Button_Data\PDF");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(@"נכשל במחיקת תיקיה 'C:\Button_Data\PDF'\nיש למחוק תיקיה זו ידנית", "מערכת ניהול מסמכים", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        isex = true;
                    }
                    if (isex) return;
                }
                Directory.CreateDirectory("C:\\Button_Data\\PDF");
                Process p = new Process();
                p.StartInfo.FileName = "C:\\Program Files\\Avision\\Button Manager\\ButtonManager.exe";
                p.StartInfo.Arguments = "/gogoscan";
                p.Start();

                while (Directory.GetFiles("C:\\Button_Data\\PDF").Count() == 0) ;
                long length;

                bool notScanned = true;
                while (true && notScanned)
                {
                    try
                    {
                        string[] scannedFiles = Directory.GetFiles("C:\\Button_Data\\PDF");
                        string path = scannedFiles[0];
                        FileInfo fi = new FileInfo(path);
                        length = fi.Length;
                        Thread.Sleep(5000);
                        while (length != (new FileInfo(path)).Length)
                        {
                            fi = new FileInfo(path);
                            length = fi.Length;
                            Thread.Sleep(5000);
                        }
                        
                        string extension = "pdf";
                        

                        string command = "INSERT INTO dbo.docnisp (shotef_mchtv, shotef_nisph, kod_marcht, kod_sug_nsph, msd_sruk, msd_df, prtim, tarich," +
                            "shm_kovtz, is_pail, shotf_mmh, kod_sivug_bithoni, is_yetzu, is_sodi, bealim, is_ishi, is_anafi, kod_kvatzaim," +
                            " user_sorek, tarich_srika, is_letzaref_mail, mail_id, ocr, colorscan, Txt, LastTxtUpdateDate, file_data, file_extension)" +
                            Environment.NewLine + "output inserted.shotef_nisph" + Environment.NewLine +
                            " VALUES (@docId, (SELECT MAX(shotef_nisph) FROM dbo.docnisp)+1," +
                            " 1, 0, @scanSerial + 1," +
                            " 0, @name, @date, @name, 1, 0, (SELECT kod_sivug_bitchoni FROM MantakDB.dbo.documents WHERE shotef_mismach=@docId), 1, 0, @owner," +
                            " 0, 0, 0, '', '00000000', 0, '', 0, 0, NULL, NULL, @data, @ext)";


                        SqlConnection conn = new SqlConnection(Global.ConStr);

                        SqlCommand c = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(msd_sruk) IS NULL THEN 0" + Environment.NewLine +
                            "ELSE MAX(msd_sruk)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.docnisp WHERE shotef_mchtv=@docId", conn);
                        c.Parameters.AddWithValue("@docId", doc.getID());
                        conn.Open();
                        int s = (int)c.ExecuteScalar();
                        conn.Close();


                        SqlCommand comm = new SqlCommand(command, conn);
                            comm.Parameters.AddWithValue("@scanSerial", s);
                        comm.Parameters.AddWithValue("@docId", doc.getID());
                        comm.Parameters.AddWithValue("@name", textBox13.Text);
                        string date = DateTime.Today.ToString("yyyyMMdd");
                        comm.Parameters.AddWithValue("@date", date);
                        comm.Parameters.AddWithValue("@owner", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
                        comm.Parameters.AddWithValue("@data", File.ReadAllBytes(path));
                        comm.Parameters.AddWithValue("@ext", extension);
                        conn.Open();
                        SqlDataReader sdr = comm.ExecuteReader();
                        sdr.Read();
                        int id = sdr.GetInt32(0);
                        notScanned = false;
                        conn.Close();
                        dataGridViewAtts.Rows.Add(id, textBox13.Text, DateTime.Today.ToShortDateString());
                        dataGridViewToSend.Rows.Add(id, textBox13.Text, DateTime.Today.ToShortDateString(), true);
                        dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 2].Cells["fileToSendColumn"].Value = false;
                        textBox13.Text = "";
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(5000);
                        continue;
                    }
                }
                Control[] controls = { button20, textBox13, button19, button21 }; // button14 - scan  - deleted
                PublicFuncsNvars.changeControlsVisiblity(false, scanControls);
                PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
                PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
                dataGridViewAtts.Visible = true;
                label9.Visible = true;
            }
            catch (Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show("בבקשה סגרו את כל קבצי ה-PDF הפתוחים מתקיית ה-button manager.");
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            textBox13.Text = "";
            PublicFuncsNvars.changeControlsVisiblity(false, scanControls);
            PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
            PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
            dataGridViewAtts.Visible = true;
            label9.Visible = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.changeControlsVisiblity(false, scanControls);
            PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            DialogResult res = ofd.ShowDialog();
            if (res == DialogResult.OK)
            {
                string path = ofd.FileName;
                string[] pathArr = path.Split('\\');
                string fileName = pathArr[pathArr.Length - 1];
                string extension = fileName.Substring(fileName.LastIndexOf('.') + 1);
                string command = "INSERT INTO dbo.docnisp (shotef_mchtv, shotef_nisph, kod_marcht, kod_sug_nsph, msd_sruk, msd_df, prtim, tarich," +
                        "shm_kovtz, is_pail, shotf_mmh, kod_sivug_bithoni, is_yetzu, is_sodi, bealim, is_ishi, is_anafi, kod_kvatzaim," +
                        " user_sorek, tarich_srika, is_letzaref_mail, mail_id, ocr, colorscan, Txt, LastTxtUpdateDate, file_data, file_extension)" +
                        Environment.NewLine + "output inserted.shotef_nisph" + Environment.NewLine +
                        " VALUES (@docId, (SELECT MAX(shotef_nisph) FROM dbo.docnisp)+1, 1, 0, @msdsruk+1," +
                        " 0, @name, @date, @name, 1, 0, (SELECT kod_sivug_bitchoni FROM MantakDB.dbo.documents WHERE shotef_mismach=@docId), 1, 0, @owner," +
                        " 0, 0, 0, '', '00000000', 0, '', 0, 0, NULL, NULL, @data, @ext)";
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(msd_sruk) IS NULL THEN 0" + Environment.NewLine +
                    "ELSE MAX(msd_sruk)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.docnisp WHERE shotef_mchtv=@docId", conn);
                comm.Parameters.AddWithValue("@docId", doc.getID());
                conn.Open();
                int msd = (int)comm.ExecuteScalar();
                conn.Close();
                comm = new SqlCommand(command, conn);
                comm.Parameters.AddWithValue("@docId", doc.getID());
                comm.Parameters.AddWithValue("@name", fileName.Substring(0, fileName.LastIndexOf('.')));
                string[] datetime = DateTime.Today.ToShortDateString().Split('/');
                string date = datetime[2].PadRight(4, '0') + datetime[1].PadRight(2, '0') + datetime[0].PadRight(2, '0');
                comm.Parameters.AddWithValue("@date", date);
                comm.Parameters.AddWithValue("@owner", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
                comm.Parameters.AddWithValue("@data", File.ReadAllBytes(path));
                comm.Parameters.AddWithValue("@ext", extension);
                comm.Parameters.AddWithValue("@msdsruk", msd);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                sdr.Read();
                int id=sdr.GetInt32(0);
                conn.Close();
                string name=fileName.Substring(0, fileName.LastIndexOf('.'));
                dataGridViewAtts.Rows.Add(id, name, DateTime.Today.ToShortDateString());
                dataGridViewToSend.Rows.Add(id, name, DateTime.Today.ToShortDateString(), true);
                dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 2].Cells["fileToSendColumn"].Value = false;
                PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
                dataGridViewAtts.Visible = true;
                label9.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Cursor.Current = Cursors.WaitCursor;

            hasBeenUpdated = false;
            if (isAllowedToEdit)
            {
                
                if (doc.isRegularDoc())
                {
                    int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());

                    if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
                    {
                        MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן פרטים.",
                                            "פרטים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    else
                    {
                        int res;
                        if (!int.TryParse(textBox2.Text, out res))
                            res = 99999;
                        //או לשים פה לפני לשנות בכללי (גרסאות) 
                      

                        hasBeenUpdated = doc.NewUpdateDetails(textBox1.Text, oldSubject, PublicFuncsNvars.getClassificationCode(oldClassification), PublicFuncsNvars.getClassificationCode(comboBox1.Text), oldUserId, res, textBox8.Text, checkBox2.Checked, checkBox3.Checked, doc, hasBeenUpdated);
                        //doc.updateDetails(textBox1.Text, PublicFuncsNvars.getClassificationCode(comboBox1.Text), textBox8.Text, checkBox2.Checked, res, checkBox3.Checked);
                        //PublicFuncsNvars.updateRecipientsInWordDoc(doc, hasBeenUpdated);
                        //hasBeenUpdated = true;
                        if (doc.getClassification() == Classification.sensitivePersonal)
                        {
                            doc.removeAllAuthorizations();
                            doc.addAuthorization(PublicFuncsNvars.curUser.userCode, true);
                            dataGridViewAuthorizations.Rows.Clear();
                            dataGridViewAuthorizations.Rows.Add(PublicFuncsNvars.curUser.userCode, PublicFuncsNvars.curUser.getFullName(),
                                PublicFuncsNvars.curUser.job, "לעריכה");
                        }

                        MessageBox.Show("פרטי המסמך נשמרו", "עדכון פרטים",
                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        
                    }

                    


                }
                else
                    MessageBox.Show("שימו לב! מסמך זה הינו מסמך ישן, ולכן לא ניתן לעדכן את המסמך אוטומטית, יש לעדכן ידני.." + Environment.NewLine +
                        "אנא פנו לצוות מחשוב", "מסמך ישן", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "עדכון פרטים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            Cursor.Current = Cursors.Default;
            UpdateTxtVer(doc.getID(), 'L');
        }

        private void textBox8_Click(object sender, EventArgs e)
        {
            button1.Visible = true;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridViewFolders.SelectedCells.Count > 0)
            {
                if (!Filed((int)dataGridViewFolders.SelectedCells[0].OwningRow.Cells[0].Value))
                {
                    DialogResult res = MessageBox.Show("האם לתייק את המסמך בתיק" + " " + dataGridViewFolders.SelectedCells[0].OwningRow.Cells[2].Value.ToString() +
                        "?", "תיוק משני", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    if (res == DialogResult.Yes)
                    {
                        SqlConnection conn = new SqlConnection(Global.ConStr);
                        SqlCommand comm = new SqlCommand("INSERT INTO dbo.tiukim(kod_marechet, shotef_klali, mispar_nose, is_rashi, mispar_in_tik)" +
                            " VALUES(2, @id, @directory, 0, (SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                            "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@directory)+1)", conn);
                        comm.Parameters.AddWithValue("@id", doc.getID());
                        DataGridViewRow row = dataGridViewFolders.SelectedCells[0].OwningRow;
                        comm.Parameters.AddWithValue("@directory", row.Cells[0].Value);
                        conn.Open();
                        try
                        {
                            comm.ExecuteNonQuery();
                            dataGridViewFiledFolders.Rows.Add(row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value, row.Cells[3].Value, false);
                            doc.addFolder(folders.Where(x => x.id == (int)row.Cells[0].Value).ToList()[0]);
                        }
                        catch(Exception ex)
                        {
                            PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                        }
                        conn.Close();
                        PublicFuncsNvars.changeControlsVisiblity(false, newFolderControls);
                        comboBox5.SelectedIndex = 0;
                        MessageBox.Show("המסמך תויק בתיק " + row.Cells[2].Value.ToString(), "תיוק משני",
                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                }
                else
                {
                    MessageBox.Show("אין אפשרות לתייק מסמך פעמיים באותו תיק", "תיוק משני", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
        }

        private bool Filed(int id)
        {
            return doc.isFiledInFolder(id);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.changeControlsVisiblity(false, newFolderControls);
            comboBox5.SelectedIndex = 0;
            panel2.Visible = false;
        }

        private void addAttButton_Click(object sender, EventArgs e)
        {
            if (PublicFuncsNvars.openDocs.Contains(doc.getID()))
            {
                MessageBox.Show("מסמך זה פתוח כרגע, יש לסגור אותו לפני שינוי פרטיו.",
                            "נספחים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (isAllowedToEdit)
            {
                PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
                PublicFuncsNvars.changeControlsVisiblity(true, AttsControls);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "נספחים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void quickPrintButton_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
            DialogResult res = MessageBox.Show("האם להדפיס מסמך זה?", "אישור הדפסה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            if (res == DialogResult.Yes)
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                comm.Parameters.AddWithValue("@id", doc.getID());
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                    string fileExt = sdr.GetString(1).Trim();
                    conn.Close();
                    string filePath = Program.folderPath + "\\print\\" + doc.getID().ToString() + "." + fileExt;
                    Cursor.Current = Cursors.WaitCursor;
                    PublicFuncsNvars.print(filePath, fileExt, fileData);
                    Cursor.Current = Cursors.Default;
                    
                }
                conn.Close();
                if (doc.getAttsIds().Count > 0)
                {
                    res = MessageBox.Show("האם להדפיס גם את הנספחים?", "אישור הדפסה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    if (DialogResult.Yes == res)
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.docnisp WHERE shotef_mchtv=@docId AND shotef_nisph=@attId AND datalength(file_data)>0", conn);
                        foreach (DataGridViewRow row in dataGridViewAtts.Rows)
                        {
                            comm.Parameters.AddWithValue("@docId", doc.getID());
                            int attId = int.Parse(row.Cells[0].Value.ToString());
                            comm.Parameters.AddWithValue("@attId", attId);
                            conn.Open();
                            sdr = comm.ExecuteReader();
                            if (sdr.Read())
                            {
                                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                                string fileExt = sdr.GetString(1).Trim();
                                conn.Close();
                                string filePath = Program.folderPath + "\\print\\" + doc.getID().ToString() + "_" + attId + "." + fileExt;
                                PublicFuncsNvars.print(filePath, fileExt, fileData);
                            }
                            comm.Parameters.Clear();
                        }
                        Cursor.Current = Cursors.Default;
                    }
                }
            }
        }

        private void browsersButton_Click(object sender, EventArgs e)
        {
            if (isAllowedToEdit)
            {
                customLabel8.Visible = false;
                dataGridViewVers.Visible = false;
                PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
                okBro = true;
                okRef = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel6.Visible = false;
                panel7.Visible = true;
                panel8.Visible = false;
                //panel9.Visible = false;
                button6.Visible = false;
                button21_Click(sender, e);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int id = (int)dataGridViewAtts.Rows[e.RowIndex].Cells["attIdColumn"].Value;
                ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(doc.getID(), id));
            }
        }

        private void viewAtt(object idObj)
        {
            KeyValuePair<int, int> ids = (KeyValuePair<int, int>)idObj;
            PublicFuncsNvars.viewAtt(ids.Key, ids.Value);
        }

        private void DocumentHandling_FormClosed(object sender, FormClosedEventArgs e)
        {
             if (!hasBeenUpdated)
            {
                Cursor.Current = Cursors.WaitCursor;
                PublicFuncsNvars.updateRecipientsInWordDoc(doc, hasBeenUpdated);
                hasBeenUpdated = true;
                hasBeenUpdatedForALotOfRecs = true;
                Cursor.Current = Cursors.Default;
            }
            PublicFuncsNvars.dhFormsOpen.Remove(doc.getID());

            if (MyGlobals.haveArgs)
            {
                Environment.Exit(1);
            }
        }

        private void foldersButton_Click(object sender, EventArgs e)
        {
            if (PublicFuncsNvars.openDocs.Contains(doc.getID()))
            {
                MessageBox.Show("מסמך זה פתוח כרגע, יש לסגור אותו לפני שינוי פרטיו.",
                            "תיקים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (isAllowedToEdit)
            {
                PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
                panel1.Visible = false;
                panel2.Visible = true;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                textBox16.Visible = false;
                button40.Visible = false;
                //panel9.Visible = false;
                customLabel8.Visible = false;
                dataGridViewVers.Visible = false;
                //panel9.Visible = false;
                button6.Visible = false;
                label14.BringToFront();

                okBro = false;
                okRef = false;
                button21_Click(sender, e);
            }
            else
                MessageBox.Show("אין לך הרשאות לערוך מסמך זה.",
                            "תיקים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void dataGridView6_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewFiledFolders.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dataGridView6_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex == 4)
                {
                    int holdingUser=PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
                    if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
                    {
                        MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל "+PublicFuncsNvars.getUserNameByUserCode(holdingUser)+", לא ניתן לעדכן תיקים.",
                                    "תיקים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    else
                    {
                        if ((bool)dataGridViewFiledFolders.Rows[e.RowIndex].Cells[4].Value)
                        {
                            DialogResult res = MessageBox.Show("האם להחליף את התיק הראשי לתיק " + dataGridViewFiledFolders.Rows[e.RowIndex].Cells[2].Value.ToString() + "?",
                                "החלפת תיק ראשי", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            if (DialogResult.Yes == res)
                            {
                                Folder d = doc.changeMainFolder((int)dataGridViewFiledFolders.Rows[e.RowIndex].Cells[0].Value);
                                if (null != d)
                                {
                                    unCheckDgv6at(d.id);
                                    setNewMainFolderDataOnScreen(dataGridViewFiledFolders.Rows[e.RowIndex]);
                                }
                                MessageBox.Show("תיק ראשי שונה ל" + dataGridViewFiledFolders.Rows[e.RowIndex].Cells[2].Value.ToString(),
                                    "החלפת תיק ראשי", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            }
                            else
                            {
                                dataGridViewFiledFolders.CellValueChanged -= dataGridView6_CellValueChanged;
                                dataGridViewFiledFolders.Rows[e.RowIndex].Cells[4].Value = false;
                                dataGridViewFiledFolders.CellValueChanged += dataGridView6_CellValueChanged;
                            }
                        }
                        else
                        {
                            MessageBox.Show("לא ניתן להשאיר מסמך ללא תיק ראשי." + Environment.NewLine + "על מנת להחליף תיק ראשי סמן תיבת סימון באחת מהשורות האחרות.",
                                    "החלפת תיק ראשי", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            dataGridViewFiledFolders.CellValueChanged -= dataGridView6_CellValueChanged;
                            dataGridViewFiledFolders.Rows[e.RowIndex].Cells[4].Value = true;
                            dataGridViewFiledFolders.CellValueChanged += dataGridView6_CellValueChanged;
                            dataGridViewFiledFolders.RefreshEdit();
                            dataGridViewFiledFolders.Refresh();
                        }
                    }
                }
            }
            Cursor.Current = Cursors.Default;
        }

        private void setNewMainFolderDataOnScreen(DataGridViewRow row)
        {
            textBox12.Text = row.Cells[0].Value.ToString();
            textBox11.Text = row.Cells[1].Value.ToString();
            textBox5.Text = row.Cells[2].Value.ToString();
            textBox6.Text = doc.getNumInFolder((int)row.Cells[0].Value).ToString();
        }

        private void unCheckDgv6at(int id)
        {
            foreach(DataGridViewRow row in dataGridViewFiledFolders.Rows)
            {
                if((int)row.Cells[0].Value==id)
                {
                    dataGridViewFiledFolders.CellValueChanged -= dataGridView6_CellValueChanged;
                    row.Cells[4].Value = false;
                    dataGridViewFiledFolders.CellValueChanged += dataGridView6_CellValueChanged;
                    break;
                }
            }
        }

        private void dataGridView6_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(e.RowIndex>=0 && e.ColumnIndex==3)
            {
                dataGridViewFiledFolders.EndEdit();
            }
            dataGridViewFiledFolders.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int holdingUser=PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                comboBox3.Text = "רמה";
                panel6.Visible = false;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel4.Visible = false;
                panel3.Visible = true;
                panel8.Visible = false;
                button6.Visible = true;
                comboBox2.Visible = false;
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            panel8.Visible = false;
            comboBox3.SelectedIndex = 0;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (okRef)
            {
                foreach (DataGridViewRow row in dataGridViewRecipientLists.SelectedRows)
                {
                    KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                    foreach (Recipient r in recipientsLists[kvp].getRecipients())
                    {
                        Recipient tempR = new Recipient(r.getId(), (short)(doc.getMaxRecipient() + (short)1), r.getRole(), r.getIFA(), true, r.getEmail());
                        if (doc.addRecipient(tempR))
                            dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), tempR.getId(),
                                PublicFuncsNvars.getUserNameByUserCode(tempR.getId()), tempR.getRole(), tempR.getIFA() ? "לפעולה" : "לידיעה");
                    }
                }

                hasBeenUpdated = false;
            }
            else
            {
                foreach (DataGridViewRow row in dataGridViewRecipientLists.SelectedRows)
                {
                    KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                    foreach (Recipient r in recipientsLists[kvp].getRecipients())
                    {
                        if (r.getId() != 99999)
                            if (doc.addAuthorization(r.getId(), comboBox6.Text == "לעריכה"))
                                dataGridViewAuthorizations.Rows.Add(r.getId(), PublicFuncsNvars.getUserNameByUserCode(r.getId()), r.getRole(), comboBox6.Text);
                    }
                }
            }
            Cursor = Cursors.Default;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = dataGridViewRecipientLists.SelectedRows[0];
            KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
            RecipientListsUpdate srl = new RecipientListsUpdate(recipientsLists[kvp], true);
            srl.Activate();
            srl.Show();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            int holdingUser=PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                panel6.Visible = false;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
                panel4.Visible = true;
                comboBox2.Visible = true;
                button6.Visible = true;
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            string interIds = "ת.פ. ";
            foreach (DataGridViewRow row8 in dataGridViewInterDist.Rows)
                if ((bool)row8.Cells["interToAddColumn"].Value)
                    interIds += row8.Cells["insIdColumn"].Value.ToString() + ",";
            if (interIds.Length > 5)
            {
                interIds = interIds.Remove(interIds.Length - 1);
                DialogResult res = MessageBox.Show("האם לכתב את " + interIds + "?", "אישור מכותב",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (DialogResult.Yes == res)
                {
                    Cursor = Cursors.WaitCursor;
                    bool k = doc.addRecipient(new Recipient(99999, (short)(doc.getMaxRecipient() + (short)1), interIds, comboBox2.Text.Equals("לפעולה"),
                        true, ""));
                    if (!k)
                    {
                        MessageBox.Show("אדם זה כבר מכותב למסמך זה", "מכותב קיים",
                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    }
                    else
                    {
                        dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), 99999, interIds, interIds, comboBox2.Text);
                        hasBeenUpdated = false;
                    }
                    panel4.Visible = false;
                    button6.Visible = false;
                }
                else
                {
                    button25_Click(button25, e);
                }
                Cursor = Cursors.Default;
            }
            else
            {
                MessageBox.Show("נא לבחור לפחות ת.פ. אחד","בחירת ת.פ. שגויה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            bool publishOriginal = false;
            bool convertOriginal = false;

            bool nothingToPublish = true;




            if (dataGridViewToSend.Rows[0].Cells["fileToSendColumn"].Value == null)
                dataGridViewToSend.Rows[0].Cells["fileToSendColumn"].Value = false;

            for (int i = 0; i < dataGridViewToSend.Rows.Count; i++)
            {
                if ((bool)dataGridViewToSend.Rows[i].Cells["fileToSendColumn"].Value)
                {
                    nothingToPublish = false;
                    break;
                }

            }

            if (nothingToPublish)
            {
                MessageBox.Show("לא נבחרו קבצים להפצה", "אין קבצים להפצה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
                if (!isOriginalNull && (bool)dataGridViewToSend.Rows[0].Cells["fileToSendColumn"].Value)
                publishOriginal = true;

            bool toConvert = false;
            if (dataGridViewToSend.Rows[0].Cells["check_PDF"].Value == null)
                toConvert = false;
            else
                toConvert = (bool)dataGridViewToSend.Rows[0].Cells["check_PDF"].Value;
            if (publishOriginal && toConvert)
                convertOriginal = true;
            int startingIndex = 0; //
            if (!isOriginalNull) startingIndex = 1; // if Original Document Exists, the nispahim is starting from row 1, if the original doc doesnt exists , the nispahim starts from row 0.
                                                    //   Dictionary<int, string> toPublish = new Dictionary<int,string>();
            Dictionary<int, Tuple<string, bool>> toPublish = new Dictionary<int, Tuple<string, bool>>();
            for (int i = startingIndex; i < dataGridViewToSend.Rows.Count; i++)
                if ((bool)dataGridViewToSend.Rows[i].Cells["fileToSendColumn"].Value)
                {
                    bool toConvert2 = false;
                    if (dataGridViewToSend.Rows[i].Cells["check_PDF"].Value == null) toConvert2 = false;
                    else
                        toConvert = (bool)dataGridViewToSend.Rows[i].Cells["check_PDF"].Value;
                    toPublish.Add((int)dataGridViewToSend.Rows[i].Cells["fileIdColumn"].Value, new Tuple<string, bool>(dataGridViewToSend.Rows[i].Cells["fileNameColumn"].Value.ToString(), toConvert2));
                }
                    if (checkBox1.Checked)
            {
                DialogResult res = MessageBox.Show("מסמך זה כבר הופץ בעבר, האם להפיץ מחדש?", "אישור הפצה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (res == DialogResult.No)
                    return;
            }
            else
            {
                DialogResult res = MessageBox.Show("האם להפיץ מסמך זה למכותבים?", "אישור הפצה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (res == DialogResult.No)
                    return;
            }
            this.Cursor = Cursors.WaitCursor;
            PublicFuncsNvars.publishDoc(doc, publishOriginal, toPublish,convertOriginal);
            this.Cursor = Cursors.Default;
            checkBox1.Checked = true;
            dateTimePicker3.Value = DateTime.Today;
            PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.changeControlsVisiblity(false, publishControls);
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }

        private void dataGridViewRecipients_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridViewRecipients.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dataGridViewRecipients_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                if (dataGridViewRecipients.SelectedCells.Count > 0)
                {
                    DataGridViewRow row = dataGridViewRecipients.SelectedCells[0].OwningRow;
                    bool ifa = row.Cells[e.ColumnIndex].Value.ToString() == "לפעולה";
                    int nid = int.Parse(row.Cells[0].Value.ToString());
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("UPDATE dbo.doc_mech SET is_lepeula=@ifa WHERE shotef_klali=@id AND msd=@nid", conn);
                    comm.Parameters.AddWithValue("@ifa", ifa);
                    comm.Parameters.AddWithValue("@id", doc.getID());
                    comm.Parameters.AddWithValue("@nid", nid);
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    doc.updateRecipientIFAByNid(nid, row.Cells[e.ColumnIndex].Value.ToString() == "לפעולה");
                    hasBeenUpdated = false;

                }
                else
                {
                    MessageBox.Show("על מנת לעדכן משתמש עליכם לסמן את המשתמש בטבלת המכותבים", "עדכון מכותב",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.changeControlsVisiblity(false, newFolderControls);
            comboBox5.SelectedIndex = 0;
            strTyped = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן תיקים.",
                                    "תיקים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                PublicFuncsNvars.changeControlsVisiblity(true, newFolderControls);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (doc.addRecipient(new Recipient(99999, -1, textBox15.Text, comboBox4.Text == "לפעולה", false, "")))
            {
                dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), 99999, textBox15.Text, textBox15.Text, comboBox4.Text);
                hasBeenUpdated = false;
            }
            Cursor = Cursors.Default;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                panel6.Visible = true;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel4.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
                button6.Visible = true;
                comboBox2.Visible = false;
            }
        }

        private void checkBox1_EnabledChanged(object sender, EventArgs e)
        {
            checkBox1.ForeColor = Color.White;
        }

        private void checkBox1_Paint(object sender, PaintEventArgs e)
        {
        // A.M   e.Graphics.DrawString(checkBox1.Text, checkBox1.Font, new SolidBrush(Color.White), e.ClipRectangle);
        }

   /*     private void button8_Click(object sender, EventArgs e)// שוטף מהמערכת
        {
            PublicFuncsNvars.changeControlsVisiblity(false, scanControls);
            PublicFuncsNvars.changeControlsVisiblity(true, addExistingDocControls);
            textBox13.Text = "";
        }*/

        private void button9_Click(object sender, EventArgs e)
        {
            int res;
            if (int.TryParse(textBox13.Text, out res))
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT hanadon, kod_sholeah, kod_sivug_bitchoni, file_data, file_extension FROM dbo.documents (nolock)"
                                                + " WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                comm.Parameters.AddWithValue("@id", res);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    string extension = sdr.GetString(4);
                    byte[] data = sdr.GetSqlBytes(3).Buffer;
                    string subject = sdr.GetString(0);
                    int owner = sdr.GetInt32(1);
                    short classification = sdr.GetInt16(2);
                    conn.Close();
                    string command = "INSERT INTO dbo.docnisp (shotef_mchtv, shotef_nisph, kod_marcht, kod_sug_nsph, msd_sruk, msd_df, prtim, tarich," +
                        "shm_kovtz, is_pail, shotf_mmh, kod_sivug_bithoni, is_yetzu, is_sodi, bealim, is_ishi, is_anafi, kod_kvatzaim," +
                        " user_sorek, tarich_srika, is_letzaref_mail, mail_id, ocr, colorscan, Txt, LastTxtUpdateDate, file_data, file_extension)" +
                        Environment.NewLine + "output inserted.shotef_nisph" + Environment.NewLine +
                        " VALUES (@docId, (SELECT MAX(shotef_nisph) FROM dbo.docnisp)+1," +
                        " 1, 0, @scanSerial+1," +
                        " 0, @name, @date, @name, 1, 1, @classification, 1, 0, @owner," +
                        " 0, 0, 0, '', '00000000', 0, '', 0, 0, NULL, NULL, @data, @ext)";
                    
                    SqlCommand c = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(msd_sruk) IS NULL THEN 0" + Environment.NewLine +
                        "ELSE MAX(msd_sruk)" + Environment.NewLine + "END" + Environment.NewLine +"FROM dbo.docnisp WHERE shotef_mchtv=@docId", conn);
                    c.Parameters.AddWithValue("@docId", doc.getID());
                    conn.Open();
                    int s = (int)c.ExecuteScalar();
                    conn.Close();


                    comm = new SqlCommand(command, conn);
                        comm.Parameters.AddWithValue("@scanSerial", s);
                    comm.Parameters.AddWithValue("@docId", doc.getID());
                    comm.Parameters.AddWithValue("@name", subject);
                    string date = DateTime.Today.ToString("yyyyMMdd");
                    comm.Parameters.AddWithValue("@date", date);
                    comm.Parameters.AddWithValue("@owner", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
                    comm.Parameters.AddWithValue("@data", data);
                    comm.Parameters.AddWithValue("@ext", extension);
                    comm.Parameters.AddWithValue("@classification", classification);
                    conn.Open();
                    sdr = comm.ExecuteReader();
                    sdr.Read();
                    int id = sdr.GetInt32(0);
                    conn.Close();
                    dataGridViewAtts.Rows.Add(id, textBox13.Text, DateTime.Today.ToShortDateString());
                    dataGridViewToSend.Rows.Add(id, textBox13.Text, DateTime.Today.ToShortDateString(), true);
                    dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 2].Cells["fileToSendColumn"].Value = false;
                    textBox13.Text = "";

                    PublicFuncsNvars.changeControlsVisiblity(false, addExistingDocControls);
                    PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
                    dataGridViewAtts.Visible = true;
                    label9.Visible = true;
                }
                else
                {
                    MessageBox.Show("אין לשוטף שנבחר נתונים בבסיס הנתונים, אנא פנו לצוות מחשוב", "הוספת נספח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
            else
            {
                MessageBox.Show("שוטף מסמך יכול להכיל רק ספרות", "הוספת נספח", MessageBoxButtons.OK, MessageBoxIcon.None,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (comboBox5.Text == "הכל")
            {
                foreach (DataGridViewRow row in dataGridViewFolders.Rows)
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

        private void button32_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            textBox24.Visible = false;
            button39.Visible = false;
            textBox24.Visible = false;
            button39.Visible = false;
            panel3.Visible = false;
            panel8.Visible = false;
            button6.Visible = true;
            comboBox6.Visible = true;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            comboBox3.Text = "רמה";
            panel5.Visible = false;
            panel3.Visible = true;
            panel8.Visible = false;
            button6.Visible = true;
            comboBox6.Visible = true;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("האם להסיר את כל ההרשאות למשתמשים מהמסמך?", "הסרת הרשאות", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            if (res == DialogResult.Yes)
            {
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
                button6.Visible = false;
                doc.removeAllAuthorizations();
                dataGridViewAuthorizations.Rows.Clear();
            }
        }

        private void dataGridViewAuthorizations_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                rowAutho = e.RowIndex;
                dataGridViewAuthorizations.Rows[rowAutho].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridViewAuthorizations_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridViewAuthorizations.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dataGridViewAuthorizations_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                if (dataGridViewAuthorizations.SelectedCells.Count > 0)
                {
                    DataGridViewRow row = dataGridViewAuthorizations.SelectedCells[0].OwningRow;
                    bool ife = row.Cells[e.ColumnIndex].Value.ToString() == "לעריכה";
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("UPDATE dbo.doc_Authorizations SET isForEdit=@ife WHERE docId=@id AND roleCode=@userCode", conn);
                    comm.Parameters.AddWithValue("@ife", ife);
                    comm.Parameters.AddWithValue("@id", doc.getID());
                    comm.Parameters.AddWithValue("@userCode", int.Parse(row.Cells[0].Value.ToString()));
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    doc.updateAuthorizationIFEByCode((int)row.Cells[0].Value, row.Cells[e.ColumnIndex].Value.ToString() == "לעריכה");

                }
                else
                {
                    MessageBox.Show("על מנת לעדכן משתמש עליכם לסמן את המשתמש בטבלת המכותבים", "עדכון מכותב",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            okBro = false;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            textBox24.Visible = false;
            button39.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            button23_Click(sender, e);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                panel6.Visible = false;
                panel8.Visible = true;
                panel5.Visible = false;
                textBox24.Visible = false;
                button39.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                button6.Visible = true;
                comboBox2.Visible = false;
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            int res;
            if (int.TryParse(textBox14.Text, out res))
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT kod_mechutav, tiur_tafkid, is_lepeula, is_lishloh_mail, ktovet_mail FROM dbo.doc_mech WHERE shotef_klali=@id", conn);
                comm.Parameters.AddWithValue("@id", res);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                while (sdr.Read())
                {
                    Recipient r = new Recipient(sdr.GetInt32(0), -1, sdr.GetString(1).Trim(), sdr.GetBoolean(2), sdr.GetBoolean(3), sdr.GetString(4).Trim());
                    if (doc.addRecipient(r))
                        dataGridViewRecipients.Rows.Add(doc.getMaxRecipient(), r.getId(), r.getRole(), r.getRole(), r.getIFA() ? "לפעולה" : "לידיעה");
                }
                conn.Close();
                hasBeenUpdated = false;
            }
            else
            {
                MessageBox.Show("שוטף יכול להכיל רק ספרות", "העתקת מכותבים ממסמך אחר",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                Cursor.Current = Cursors.WaitCursor;
                if (dataGridViewRecipients.SelectedCells.Count > 0)
                {
                    if (doc.moveRecipientUp(int.Parse(dataGridViewRecipients.SelectedCells[0].OwningRow.Cells[0].Value.ToString())))
                    {
                        DataGridViewRow row = dataGridViewRecipients.SelectedCells[0].OwningRow;
                        int index = row.Index;
                        dataGridViewRecipients.Rows.RemoveAt(index);
                        dataGridViewRecipients.Rows.Insert(index - 1, row);
                        short nid = short.Parse(row.Cells[0].Value.ToString());
                        row.Cells[0].Value = dataGridViewRecipients.Rows[index].Cells[0].Value;
                        dataGridViewRecipients.Rows[index].Cells[0].Value = nid;
                        row.Cells[0].Selected = true;
                        hasBeenUpdated = false;
                    }

                }
                else
                {
                    MessageBox.Show("אין מכותב מסומן", "שינוי סדר מכותבים",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
                Cursor.Current = Cursors.Default;
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                Cursor.Current = Cursors.WaitCursor;
                if (dataGridViewRecipients.SelectedCells.Count > 0)
                {
                    if (doc.moveRecipientDown(int.Parse(dataGridViewRecipients.SelectedCells[0].OwningRow.Cells[0].Value.ToString())))
                    {
                        DataGridViewRow row = dataGridViewRecipients.SelectedCells[0].OwningRow;
                        int index = row.Index;
                        dataGridViewRecipients.Rows.RemoveAt(index);
                        dataGridViewRecipients.Rows.Insert(index + 1, row);
                        short nid = short.Parse(row.Cells[0].Value.ToString());
                        row.Cells[0].Value = dataGridViewRecipients.Rows[index].Cells[0].Value;
                        dataGridViewRecipients.Rows[index].Cells[0].Value = nid;
                        row.Cells[0].Selected = true;
                        hasBeenUpdated = false;
                    }

                }
                else
                {
                    MessageBox.Show("אין מכותב מסומן", "שינוי סדר מכותבים",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
                Cursor.Current = Cursors.Default;
            }
        }

        private void dataGridViewRecipients_CellLeave(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridViewRecipients_Enter(object sender, EventArgs e)
        {
            dataGridViewRecipients.CurrentCell = null;
        }

        private void button35_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
            {
                MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                    "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                hasBeenUpdatedForALotOfRecs = false;
                hasBeenUpdated = false;
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            int holdingUser = PublicFuncsNvars.whoHoldsThisDoc(doc.getID());
            if (PublicFuncsNvars.curUser.userCode == holdingUser || holdingUser == 0)
            {
                if (PublicFuncsNvars.releaseHeldDoc(doc.getID()))
                    MessageBox.Show("המסמך שוחרר בהצלחה.", "שחרור מסמך", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                MessageBox.Show("לא ניתן לשחרר מסמך שמוחזק ע\"י אדם אחר.", "שחרור מסמך", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            DocumentHandling_KeyDown(sender, new KeyEventArgs(Keys.K | Keys.Control));
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            int res;
            if(int.TryParse(textBox2.Text, out res))
            {
                User u = PublicFuncsNvars.users.Find(x => x.userCode == res);
                if(u!=null)
                {
                    textBox3.Text = u.job;
                    textBox4.Text = u.getFullName();
                }
            }
        }

        private Rectangle dragBoxFromMouseDown;
        private int rowIndexOfItemUnderMouseToDrop;
        private short toMove, newLocation;

        private void dataGridViewRecipients_MouseMove(object sender, MouseEventArgs e)
        {
            if((e.Button&MouseButtons.Left)==MouseButtons.Left)
            {
                dataGridViewRecipients.CellLeave -= dataGridViewRecipients_CellLeave;
                if (dragBoxFromMouseDown != Rectangle.Empty && !dragBoxFromMouseDown.Contains(e.X, e.Y))
                {
                    DragDropEffects dropEffects = dataGridViewRecipients.DoDragDrop(dataGridViewRecipients.Rows[rowIndexFromMouseDown], DragDropEffects.Move);
                    if (rowIndexOfItemUnderMouseToDrop != rowIndexFromMouseDown)
                    {
                        doc.moveRecipient(toMove, newLocation);
                        reloadRecipients();
                        hasBeenUpdated = false;
                    }
                }
                dataGridViewRecipients.CellLeave += dataGridViewRecipients_CellLeave;
            }
        }

        private void dataGridViewRecipients_MouseDown(object sender, MouseEventArgs e)
        {
            rowIndexFromMouseDown = dataGridViewRecipients.HitTest(e.X, e.Y).RowIndex;
            if (rowIndexFromMouseDown != -1)
            {
                Size dragSize = SystemInformation.DragSize;
                dragBoxFromMouseDown = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
            }
            else
                dragBoxFromMouseDown = Rectangle.Empty;
        }

        private void dataGridViewRecipients_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;

            Rectangle r = dataGridViewRecipients.RectangleToScreen(dataGridViewRecipients.ClientRectangle);
            if (r.Top+dataGridViewRecipients.ColumnHeadersHeight >= e.Y && dataGridViewRecipients.FirstDisplayedScrollingRowIndex > 0)
                dataGridViewRecipients.FirstDisplayedScrollingRowIndex--;
            else if (r.Bottom-10 <= e.Y && dataGridViewRecipients.FirstDisplayedScrollingRowIndex < dataGridViewRecipients.Rows.Count - 1)
                dataGridViewRecipients.FirstDisplayedScrollingRowIndex++;
        }

        private void dataGridViewRecipients_DragDrop(object sender, DragEventArgs e)
        {
            Point clientPoint = dataGridViewRecipients.PointToClient(new Point(e.X, e.Y));
            rowIndexOfItemUnderMouseToDrop = dataGridViewRecipients.HitTest(clientPoint.X, clientPoint.Y).RowIndex;
            if (rowIndexOfItemUnderMouseToDrop == -1)
                rowIndexOfItemUnderMouseToDrop = dataGridViewRecipients.Rows.Count - 1;

            if(e.Effect==DragDropEffects.Move)
            {
                if (rowIndexOfItemUnderMouseToDrop != rowIndexFromMouseDown)
                {
                    DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                    toMove = (short)dataGridViewRecipients.Rows[rowIndexFromMouseDown].Cells[0].Value;
                    newLocation = (short)dataGridViewRecipients.Rows[rowIndexOfItemUnderMouseToDrop].Cells[0].Value;
                    dataGridViewRecipients.Rows.RemoveAt(rowIndexFromMouseDown);
                    dataGridViewRecipients.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove);
                }
            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                DateTime selectedTime = dateTimePicker3.Value;
                doc.setPublishDate(selectedTime,checkBox1.Checked);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                dateTimePicker3.Enabled = true;
                doc.setPublishDate(dateTimePicker3.Value,true);
            }

            else
            {
                dateTimePicker3.Enabled = false;
                doc.setPublishDate(dateTimePicker3.Value, false);
            }
        }

        private void dataGridViewFolders_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button11_Click(sender,e);
        }

        private void dataGridViewUsers_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button5_Click(sender,e);
        }

        private void dataGridViewToSend_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void dataGridViewToSend_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void dataGridViewAtts_DragDrop(object sender, DragEventArgs e)
        {
            string mailPath = "";
            
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                if (app == null)
                    app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.Explorer oExplorer = app.ActiveExplorer();
                Microsoft.Office.Interop.Outlook.Selection oSelection = oExplorer.Selection;
                if (oSelection.Count != 1) return;
                Microsoft.Office.Interop.Outlook.MailItem mi = null;
                try
                {
                    mi = (Microsoft.Office.Interop.Outlook.MailItem)oSelection[1];
                }

                catch { return; }
                mailPath = string.Join("", mi.Subject.Split(Path.GetInvalidFileNameChars()));
                mailPath += ".msg";
                mailPath = "c:\\temp\\" + mailPath;
                mi.SaveAs(mailPath);

                fi = new FileInfo(mailPath);

            }

            else
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Any())
                {
                     fi = new FileInfo(files[0]);
                }
            }


            string path = fi.FullName;
            string fileName = fi.Name; //   pathArr[pathArr.Length - 1];

                string extension = fi.Extension; // fileName.Substring(fileName.LastIndexOf('.') + 1);
                string command = "INSERT INTO dbo.docnisp (shotef_mchtv, shotef_nisph, kod_marcht, kod_sug_nsph, msd_sruk, msd_df, prtim, tarich," +
                        "shm_kovtz, is_pail, shotf_mmh, kod_sivug_bithoni, is_yetzu, is_sodi, bealim, is_ishi, is_anafi, kod_kvatzaim," +
                        " user_sorek, tarich_srika, is_letzaref_mail, mail_id, ocr, colorscan, Txt, LastTxtUpdateDate, file_data, file_extension)" +
                        Environment.NewLine + "output inserted.shotef_nisph" + Environment.NewLine +
                        " VALUES (@docId, (SELECT MAX(shotef_nisph) FROM dbo.docnisp)+1, 1, 0, @msdsruk+1," +
                        " 0, @name, @date, @name, 1, 0, (SELECT kod_sivug_bitchoni FROM MantakDB.dbo.documents WHERE shotef_mismach=@docId), 1, 0, @owner," +
                        " 0, 0, 0, '', '00000000', 0, '', 0, 0, NULL, NULL, @data, @ext)";
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(msd_sruk) IS NULL THEN 0" + Environment.NewLine +
                    "ELSE MAX(msd_sruk)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.docnisp WHERE shotef_mchtv=@docId", conn);
                comm.Parameters.AddWithValue("@docId", doc.getID());
                conn.Open();
                int msd = (int)comm.ExecuteScalar();
                conn.Close();
                comm = new SqlCommand(command, conn);
                comm.Parameters.AddWithValue("@docId", doc.getID());
                comm.Parameters.AddWithValue("@name", fileName.Substring(0, fileName.LastIndexOf('.')));
                string[] datetime = DateTime.Today.ToShortDateString().Split('/');
                string date = datetime[2].PadRight(4, '0') + datetime[1].PadRight(2, '0') + datetime[0].PadRight(2, '0');
                comm.Parameters.AddWithValue("@date", date);
                comm.Parameters.AddWithValue("@owner", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
                comm.Parameters.AddWithValue("@data", File.ReadAllBytes(path));
                comm.Parameters.AddWithValue("@ext", extension);
                comm.Parameters.AddWithValue("@msdsruk", msd);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                sdr.Read();
                int id = sdr.GetInt32(0);
                conn.Close();
                string name = fileName.Substring(0, fileName.LastIndexOf('.'));
                dataGridViewAtts.Rows.Add(id, name, DateTime.Today.ToShortDateString());
                dataGridViewToSend.Rows.Add(id, name, DateTime.Today.ToShortDateString(), true);
                dataGridViewToSend.Rows[dataGridViewToSend.Rows.Count - 2].Cells["fileToSendColumn"].Value = false;
                PublicFuncsNvars.changeControlsVisiblity(false, AttsControls);
                dataGridViewAtts.Visible = true;
                label9.Visible = true;
            }
        

        private void dataGridViewAtts_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void DocumentHandling_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Cursor == Cursors.WaitCursor)
            {
                Console.WriteLine(this.Cursor);
                Console.WriteLine(Cursors.WaitCursor);
                MessageBox.Show("אנא סגור את המסמך הפתוח לעריכה לפני יציאה מהחלון", "בעיה בסגירת חלון", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                e.Cancel = true;

            }
            if (!TheDocCaChange)
            {
                _timer.Stop();
                return;
            }
                
            if (OpenedForEdit)
            {

                bool deleteTempFile = false;

                /*
                1. סגירת חלון
2. האם המסמך קיים
	3. בודק את הגודל 
	4. אם הגודל שונה
		5. שואל האם לשמור שינויים
			6. שומר את הגירסה
			7. בודק סיומת
			8. האם הסיומת  היא doc או docx 
				9. קורא את הטקסט 
			10. מעדכן במסד נתונים
                */


                int id = doc.getID();
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT file_extension, datalength(file_data) FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data) >= 1", conn);//file_data, // datalength(file_data)>0
                comm.Parameters.AddWithValue("@id", id);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();

                if (sdr.Read())
                {
                    string fileExt = sdr.GetString(0).Trim();
                    string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                    if (File.Exists(filePath))
                    {
                        object fileDBSize = sdr.GetSqlValue(1);
                        if (long.Parse(fileDBSize.ToString()) != (new FileInfo(filePath)).Length)
                        {
                            DialogResult res = MessageBox.Show("האם לשמור השינויים בקובץ המסמך?", "שמירת שינויים", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            if (res == DialogResult.Yes)
                            {
                                try
                                {
                                    DocumentHandling.SaveVersion(id);

                                    if (fileExt.ToLower().Contains("doc"))
                                    {
                                        Word.Document document = null;
                                        Text = PublicFuncsNvars.docToTxt(document, filePath);
                                    }
                                    byte[] fileData = File.ReadAllBytes(filePath);
                                    PublicFuncsNvars.saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, Text);
                                    deleteTempFile = true;
                                }
                                catch
                                {
                                    MessageBox.Show("המסמך פתוח במחשב." + Environment.NewLine + "יש לסגור את המסמך ואז לסגור את המסך.", "שמירת שינויים",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                                    e.Cancel = true;
                                    return;
                                }

                            }
                            else
                                deleteTempFile = true;
                        }
                    }
                }
                if (deleteTempFile)
                {
                    string filePath = Program.folderPath + "\\" + shotef + "." + docExt;
                   // MessageBox.Show(filePath);
                    File.Delete(filePath);
                }
           }
            _timer.Stop();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            
            SqlConnection conn1 = new SqlConnection(Global.ConStr);
            SqlCommand comm1 = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn1);
            comm1.Parameters.AddWithValue("@id", doc.getID());
            conn1.Open();
            SqlDataReader sdr1 = comm1.ExecuteReader();
            if (sdr1.Read())
            {
                string mainFile_extention = sdr1.GetString(1);
                //isOriginalNull = sdr1.IsDBNull(0);
                byte[] fileData = sdr1.GetSqlBytes(0).Buffer;
                string filePath = Program.folderPath + "\\" + doc.getID().ToString() + "." + mainFile_extention;
                string htmlPath = Path.ChangeExtension(filePath, ".html");
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                }
                Word.Document docu = new Word.Document();
                
                //PreviewHandlerHost.Open(filePath)
                //string htmlPath = Path.ChangeExtension(filePath, ".html");
                /*if (!File.Exists(htmlPath))
                {
                    Word.Application wapp;
                    try
                    {
                        wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                    }
                    catch
                    {
                        wapp = new Word.Application();
                    }
                    Word.Document docu = wapp.Documents.Open(filePath);
                    docu.SaveAs2(htmlPath, Word.WdSaveFormat.wdFormatHTML);
                    docu.Close();
                    wapp.Quit();
                }*/
                //webBrowser1.Navigate(filePath);
                //panel9.Visible = true;
            }
            else
            {
                MessageBox.Show("הממך ריק");
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
           
        }

        private void BtnVer_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            textBox24.Visible = false;
            button39.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            button6.Visible = false;
            customLabel8.Visible = true;
            dataGridViewVers.Visible = true;
            
            /*SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT * FROM dbo.F_GetDocVers(@ShotefId)", conn);
            comm.Parameters.AddWithValue("@ShotefId", doc.getID());
            comm.Parameters.AddWithValue("@P_TopType", 'A');
            conn.Open();
            using (SqlDataReader reader = comm.ExecuteReader())
            {
                DataTable VersDataTable = new DataTable();
                VersDataTable.Clear();
                VersDataTable.Load(reader);
                dataGridViewVers.DataSource = VersDataTable;
                dataGridViewVers.Columns["VerFileData"].Visible = false;
                dataGridViewVers.Columns["VerFileExt"].Visible = false;
                dataGridViewVers.Refresh();
                dataGridViewVers.Visible = true;
            }
            //SqlDataAdapter sad = new SqlDataAdapter(comm);
            //DataTable VersDataTable = new DataTable();
            //sad.Fill(VersDataTable);
            conn.Close();*/
            /*dataGridViewVers.DataSource = VersDataTable;
            dataGridViewVers.Columns["VerFileData"].Visible = false;
            dataGridViewVers.Columns["VerFileExt"].Visible = false;
            dataGridViewVers.Refresh();
            dataGridViewVers.Visible = true;*/
        }
        public void UpdateTxtVer(int shotef, char type)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT * FROM dbo.F_GetDocVers(@ShotefId)", conn);
            comm.Parameters.AddWithValue("@ShotefId", doc.getID());
            comm.Parameters.AddWithValue("@P_TopType", type);
            conn.Open();
            using (SqlDataReader reader = comm.ExecuteReader())
            {
                DataTable VersDataTable = new DataTable();
                VersDataTable.Clear();
                VersDataTable.Load(reader);
                dataGridViewVers.DataSource = VersDataTable;
                dataGridViewVers.Columns["VerFileData"].Visible = false;
                dataGridViewVers.Columns["VerFileExt"].Visible = false;
                dataGridViewVers.Refresh();
            }
            if (dataGridViewVers.Rows.Count > 0)
            {
                label20.Text = $"({dataGridViewVers.Rows[0].Cells["VerNum"].Value} , {dataGridViewVers.Rows[0].Cells["VerDTime"].Value})";
            }
            else
                label20.Text = "";
            conn.Close();
        } 
        private void btn_Share_Click(object sender, EventArgs e)
        {
            string to = "";
            string cc = "";
            string bcc = "";
            string body = "";
            string mailSubject = "שיתוף מסמך ";
            int id = shotef;
            string subject = textBox1.Text;
            string classification = PublicFuncsNvars.getClassificationByEnum(doc.getClassification());
            mailSubject += id + " - " + subject + " @" + classification + "@";
            List<Tuple<byte[], string, bool>> attachments = new List<Tuple<byte[], string, bool>>();


            string from = PublicFuncsNvars.curUser.email;
            string res = PublicFuncsNvars.CreateShortcut(id);

            body += res;
            PublicFuncsNvars.sendShareMail(from, to, cc, bcc, mailSubject, body, null);
        }

        private void dataGridViewVers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0&& e.ColumnIndex==0)
            {
                var dgv = sender as DataGridView;
                if (dgv!= null && dgv[e.ColumnIndex, e.RowIndex] is DataGridViewLinkCell)
                {
                    var cellRectangle = dgv.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
                    var relativeMousePosition = dgv.PointToClient(Cursor.Position);
                    if (cellRectangle.Contains(relativeMousePosition))
                    {
                        dataGridViewVers_CellDoubleClick(sender, e);
                    }
                }
            }
        }
        private void dataGridViewVers_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                byte[] fileData = (byte[])dataGridViewVers.Rows[e.RowIndex].Cells["VerFileData"].Value;
                string fileExt = dataGridViewVers.Rows[e.RowIndex].Cells["VerFileExt"].Value.ToString();
                int id = doc.getID();
                string verNum;
                try
                {
                    verNum = dataGridViewVers.Rows[e.RowIndex].Cells["VerNum"].Value.ToString();

                }
                catch (Exception a)
                {
                    verNum = dataGridViewVers.Rows[e.RowIndex].Cells[0].Value.ToString();
                }

                ThreadPool.QueueUserWorkItem(new WaitCallback(OpenWordDocument), new ThreadState { FileData = fileData, FileExt = fileExt, FileName = id, FileVers = verNum });
            }
        }
        
        private void OpenWordDocument(object state)
        {
            var threadState = state as ThreadState;
            if (threadState == null)
            {
                MessageBox.Show("error");
                return;
            }
            string fileName = threadState.FileName + "_" + threadState.FileVers;
            MyGlobals.afterViewOnly = true;
            Cursor.Current = Cursors.WaitCursor;
            if (!PublicFuncsNvars.openVerDocs.Contains(fileName))
            {
                PublicFuncsNvars.openVerDocs.Add(fileName);
                string filePath = Program.folderPath + "\\" + fileName + "." + threadState.FileExt;
                if (File.Exists(filePath))
                {
                    try
                    {
                        string archiveFolder = Program.archiveFolder + "/Archive/";// Path.GetDirectoryName(filePath) + "/Archive/";
                        Directory.CreateDirectory(archiveFolder);
                        string copyTo = archiveFolder + "_" + Path.GetFileNameWithoutExtension(filePath) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(filePath);
                        File.Move(filePath, copyTo);
                    }
                    catch { }
                }

                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, threadState.FileData);
                    Process.Start(filePath);
                    Cursor.Current = Cursors.Default;
                    Application.UseWaitCursor = false;
                    while (true)
                    {
                        Thread.Sleep(5000);
                        try
                        {
                            File.Delete(filePath);
                            PublicFuncsNvars.openVerDocs.Remove(fileName);
                            break;
                        }
                        catch (Exception ex) { }
                    }
                }
                else
                {
                    PublicFuncsNvars.openVerDocs.Remove(fileName);
                }

            }
            else
            {
                MessageBox.Show("מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסמך מספר פעמים", "מסמך פתוח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }

        }

        private void dataGridViewUsers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            dataGridViewUsers.Rows.Clear();
            List<User> FilteredUsers = users.Where(word => word.firstName.Contains(textBox24.Text) || word.userCode.ToString().Contains(textBox24.Text) || word.lastName.Contains(textBox24.Text) || word.job.ToString().Contains(textBox24.Text)).ToList();
            foreach (User u in FilteredUsers)
                dataGridViewUsers.Rows.Add(u.userCode, u.firstName, u.lastName, u.job, u.userCode + ";" + u.firstName + ";" + u.lastName + ";" + u.job);
            dataGridViewUsers.Refresh();
        }

        private void button39_Click(object sender, EventArgs e)
        {
            textBox24.Text = "";
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

            dataGridViewFolders.Rows.Clear();
            foreach (Folder d in folders)
            {
                if (d.id.ToString().Contains(textBox16.Text) || d.shortDescription.ToString().Contains(textBox16.Text) || d.description.ToString().Contains(textBox16.Text))
                dataGridViewFolders.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch));
            }
            
        }

        private void button40_Click(object sender, EventArgs e)
        {
           
        }

        private void button40_Click_1(object sender, EventArgs e)
        {
            textBox16.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridViewToSend_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void reloadRecipients()
        {
            dataGridViewRecipients.Rows.Clear();
            dataGridViewRecipients.CellValueChanged -= dataGridViewRecipients_CellValueChanged;

            foreach (Recipient r in doc.getRecipients())
            {
                string name = PublicFuncsNvars.getUserNameByUserCode(r.getId());
                if (name == null)
                    name = r.getRole();
                int rowIndex = dataGridViewRecipients.Rows.Add(r.getNID(), r.getId(), name, r.getRole());
                dataGridViewRecipients.Rows[rowIndex].Cells[4].Value = r.getIFA() ? "לפעולה" : "לידיעה";
                if (dataGridViewRecipients.Rows[rowIndex].Cells[3].Value.ToString().StartsWith("ת.פ."))
                {
                    dataGridViewRecipients.Rows[rowIndex].Cells[3].ReadOnly = true;
                }
            }
            dataGridViewRecipients.Sort(dataGridViewRecipients.Columns[0], ListSortDirection.Ascending);

            dataGridViewRecipients.CellValueChanged += dataGridViewRecipients_CellValueChanged;
        }

        private void dataGridViewAtts_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex > 0)
            dataGridViewAtts.Rows[e.RowIndex].Selected = true;
        }
        
        public static void SaveVersion(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            
            SqlCommand COMMAND = new SqlCommand("SP_AddDocVer_2", conn);
            COMMAND.CommandType = CommandType.StoredProcedure;
            COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMismach", id));
            COMMAND.Parameters.Add(new SqlParameter("@P_VerUser", PublicFuncsNvars.curUser.getFullName()));

            SqlParameter rtnVal = COMMAND.Parameters.Add("@ReturnVal", SqlDbType.Int);// CreateParameter();
            rtnVal.Direction = ParameterDirection.ReturnValue;
            conn.Open();
            //COMMAND.Parameters.Add(rtnVal);
            
            try
            {
                COMMAND.ExecuteNonQuery();
                var result = rtnVal.Value;
                //int r = (int)COMMAND.Parameters["@RC"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();
            
            //UpdateTxtVer(id, "L");
        }
        //private void Importing(string FilePath)
        //{
        //    string HtmlString = string.Empty;
        //    string path = System.Web.HttpContext.Current.Server.MapPath(FilePath);
        //    if (path != null)
        //    {
        //        using (MemoryStream mStream = new MemoryStream())
        //        {
        //            new Word.Document(path).Save(mStream, FormatType.Html);
        //        }
        //    }
        //}
    }
}
