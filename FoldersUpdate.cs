using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace DocumentsModule
{
    public partial class FoldersUpdate : Form
    {
        List<Folder> foldersView;
        public FoldersUpdate()
        {
            InitializeComponent();
        }

        private void FoldersUpdate_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
            foldersView = new List<Folder>();
            foldersView.AddRange(PublicFuncsNvars.folders);
            checkedListBox1.DataSource = foldersView;
            checkedListBox1.DisplayMember = "description";
            checkedListBox1.ValueMember = "id";
            DocumentsMenu.PathTemplate(this.button1, 55);
            DocumentsMenu.PathTemplate(this.button16, 55);
            DocumentsMenu.PathTemplate(this.button7, 55);
            DocumentsMenu.PathTemplate(this.button9, 30);
            DocumentsMenu.PathTemplate(this.button8, 30);
            DocumentsMenu.PathTemplate(this.button4, 30);
            DocumentsMenu.PathTemplate(this.button5, 30);
            DocumentsMenu.PathTemplate(this.button2, 30);
            DocumentsMenu.PathTemplate(this.button3, 30);
            DocumentsMenu.PathTemplate(this.button6, 30);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.getFolders();
            dataGridViewFolders.Rows.Clear();
            int res1=0;
            bool tb1Empty=textBox1.Text.Equals(""), cbEmpty=comboBox3.Text.Equals("")||comboBox3.Text.Equals("הכל");
            if (int.TryParse(textBox1.Text, out res1)||textBox1.Text=="")
            {
                Branch b = (Branch)PublicFuncsNvars.getBranchByString(comboBox3.Text);

                var foldersToDisplay =
                    from folder in PublicFuncsNvars.folders
                    where (folder.id == res1 || tb1Empty) && folder.shortDescription.Contains(textBox2.Text) &&
                           folder.description.Contains(textBox3.Text) && (cbEmpty || folder.branch == b)
                    orderby folder.id descending
                    select folder;

                foreach (Folder f in foldersToDisplay)
                    dataGridViewFolders.Rows.Add(f.id, f.shortDescription, f.description, PublicFuncsNvars.getBranchString(f.branch));
                if (dataGridViewFolders.Rows.Count == 0)
                {
                    MessageBox.Show("לא נמצאו תיקים",
                                    "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
            else
            {
                MessageBox.Show("לא ניתן להכניס תווים שאינם ספרות לתיבת 'מספר תיק'",
                                "נתונים שגויים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                textBox1.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            comboBox3.SelectedItem = null;
            dataGridViewFolders.Rows.Clear();
        }

        private void FoldersUpdate_FormClosed(object sender, FormClosedEventArgs e)
        {
            Program.fu = null;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (PublicFuncsNvars.curUser.allowedToOpenFolders||PublicFuncsNvars.curUser.roleType==RoleType.computers)
            {
                textBox4.Clear();
                textBox5.Clear();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                panel1.Visible = false;
                panel2.Visible = true;
                panel3.Visible = false;
                button5.BringToFront();
                textBox7.Text = "";
                checkBox1.Visible = false;
                comboBox1.Enabled = true;
            }
            else
            {
                MessageBox.Show("אין לך הרשאות לפתוח תיקים",
                                "בעית הרשאות", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int res=-1;
            if(PublicFuncsNvars.shortDescExists(textBox5.Text))
            {
                MessageBox.Show("שם מקוצר זה כבר קיים עבור תיק אחר, אנא בחרו שם מקוצר אחר",
                                "שם מקוצר קיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (textBox5.Text == "" || textBox4.Text == "" || comboBox1.Text == "" || comboBox2.Text == "")
            {
                MessageBox.Show("אין להשאיר את השדות 'שם מקוצר', 'שם תיק', 'סוג תיק' ו-'סיווג' ריקים",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (comboBox1.Text == "שו\"ש" && (!int.TryParse(textBox7.Text, out res) || res <= 0))
            {
                MessageBox.Show("אין ליצור תיק מסוג שו\"ש בלי מספר שו\"ש או עם מספר שו\"ש לא חיובי",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("INSERT INTO dbo.tm_mesimot (ms_mshimh, shm_mshimh, shm_mkotzr, sog_mshimh, ms_mshimt_am, mkt, "
                    + "simn, ms_archh_shosh, msd_proiikt, is_tik, is_tik_pail, anp, harot_mshimh, iishom_shosh_isdrti, aomdn_shosh, aomdn_shosh_o, "
                    + "mtba_aomdn_shosh, alot_ishom_lin_bnih, aishom_liin_bnih, mtba_aishom_lin_b, alot_btzoa_tnk_atdi, abtzoa_tnk_atidi, mtba_abtzoa_tatdi, "
                    + "mshimh_rgishh, User_Update, Date_Update, Time_Update, User_Create, Date_Create, Time_Create, ms_pitoh, sog_ptoh_shosh_Obsolete, "
                    + "ailn_iohsin_Obsolete, rmt_ailn_Obsolete, is_Print_Tag_Obsolete, kod_sioog, kod_mtzb_mshimh, kod_adipot, tarich_hthlh_mkori, "
                    + "tarich_hthlh_mtochnn, tarich_hthlh_bpoal, tarich_siom_mkori, tarich_siom_mtochnn, tarich_siom_bpoal, mspr_mkori, hzmnh_miohdt, aopion, "
                    + "SOW, shrtoti_rpt, B_F, tik_shrtotim, mshkl_tchni, mshkl_chlchli, mshkl_loz, mshkl_niholi, tzion_sp_tchni, tzion_sp_niholi, mspr_pitoh) "
                    + Environment.NewLine+"OUTPUT INSERTED.ms_mshimh"+Environment.NewLine
                    + "VALUES((SELECT MAX(ms_mshimh) FROM dbo.tm_mesimot)+1, @description, @shortDesc, @type, 0, 0, '', @shoshNum, 0, 1, 1, @branch, @notes, '', 0, '', '', 0, '', '', 0, '', '', 0, "
                    + "@creatingUser, @creationDate, @creationTime, @creatingUser, @creationDate, @creationTime, 0, '', '', 0, 0, @classification, 1, 0, "
                    + "'00000000', '00000000', @creationDate, '00000000', '00000000', '00000000', 0, 0, '', '', '', '', '', 0, 0, 0, 0, 0, 0, '')", conn);
                comm.Parameters.AddWithValue("@description", textBox4.Text);
                comm.Parameters.AddWithValue("@shortDesc", textBox5.Text);
                comm.Parameters.AddWithValue("@type", comboBox1.Text[0]);
                comm.Parameters.AddWithValue("@shoshNum", res == -1 ? 0 : res);
                comm.Parameters.AddWithValue("@branch", (char)PublicFuncsNvars.curUser.branch);
                comm.Parameters.AddWithValue("@notes", textBox6.Text);
                comm.Parameters.AddWithValue("@creatingUser", PublicFuncsNvars.curUser.userCode);
                comm.Parameters.AddWithValue("@creationDate", DateTime.Today.ToString("yyyyMMdd"));
                comm.Parameters.AddWithValue("@creationTime", DateTime.Now.TimeOfDay.ToString("hhmmss"));
                comm.Parameters.AddWithValue("@classification", PublicFuncsNvars.getClassificationCode(comboBox2.Text));

                conn.Open();
                int id = (int)comm.ExecuteScalar();
                conn.Close();

                if (res == -1)
                    PublicFuncsNvars.folders.Add(new Folder(id, textBox4.Text, textBox5.Text, true, PublicFuncsNvars.curUser.branch, true, (FileType)(comboBox1.Text[0]),
                                   PublicFuncsNvars.getClassification(PublicFuncsNvars.getClassificationCode(comboBox2.Text))));
                else
                    PublicFuncsNvars.folders.Add(new ShoshFolder(id, textBox4.Text, textBox5.Text, true, PublicFuncsNvars.curUser.branch, true, (FileType)(comboBox1.Text[0]),
                                        PublicFuncsNvars.getClassification(PublicFuncsNvars.getClassificationCode(comboBox2.Text)), res));

                panel2.Visible = false;
                panel1.Visible = true;

                MessageBox.Show("תיק מספר "+id+": "+textBox4.Text+" נוצר בהצלחה",
                                "אישור יצירת תיק", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel1.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] headers ={"מספר תיק", "שם תיק מקוצר", "תיאור תיק", "קוד ענף", "תיאור ענף", "פעיל?"};

            DialogResult res = MessageBox.Show("האם לכלול תיקים לא פעילים?",
                "דו\"ח תיקים", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            if (res != DialogResult.Cancel)
            {
                List<string[]> values = new List<string[]>();
                IEnumerable<Folder> folders = PublicFuncsNvars.folders;
                if(res==DialogResult.No)
                {
                    folders =
                        from folder in PublicFuncsNvars.folders
                        where folder.isActive
                        orderby folder.id descending
                        select folder;
                }

                foreach(Folder f in folders)
                {
                    string[] valRow = { f.id.ToString(), f.shortDescription, f.description, ((char)f.branch).ToString(),
                                          PublicFuncsNvars.getBranchString(f.branch), f.isActive.ToString() };
                    values.Add(valRow);
                }

                PublicFuncsNvars.exportToXL("files-list", "רשימת תיקים", headers, values);
            }

            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "שו\"ש")
            {
                label9.Visible = true;
                textBox7.Visible = true;
                textBox7.Text = "";
            }
            else
            {
                label9.Visible = false;
                textBox7.Visible = false;
            }
        }

        private void foldersDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex>=0)
            {
                Folder f = PublicFuncsNvars.folders.Where(x => x.id == (int)dataGridViewFolders.Rows[e.RowIndex].Cells[0].Value).ToList()[0];
                textBox4.Text = dataGridViewFolders.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox5.Text = dataGridViewFolders.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox1.SelectedIndex = comboBox1.Items.IndexOf(PublicFuncsNvars.getfileTypeString(f.type));
                comboBox2.SelectedIndex = comboBox2.Items.IndexOf(PublicFuncsNvars.getClassificationByEnum(f.classification));
                panel1.Visible = false;
                panel2.Visible = true;
                comboBox1.Enabled = true;// false;
                textBox7.Enabled = false;
                button6.BringToFront();
                checkBox1.Visible = true;
                checkBox1.Checked = f.isActive;
                if (f.type == FileType.shosh)
                {
                    textBox7.Text = ((ShoshFolder)f).shoshNum.ToString();
                    textBox7.Visible = true;
                }
                else
                {
                    textBox7.Text = "";
                    textBox7.Visible = false;
                }
                textBox8.Text = f.id.ToString();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.tm_mesimot SET shm_mshimh=@description, shm_mkotzr=@shortDesc,sog_mshimh=@type, is_tik_pail=@isActive, "
                + "harot_mshimh=@notes, User_Update=@curUser, Date_Update=@curDate, Time_Update=@curTime, kod_sioog=@classification WHERE ms_mshimh=@id", conn);
            comm.Parameters.AddWithValue("@id", int.Parse(textBox8.Text));
            comm.Parameters.AddWithValue("@description", textBox4.Text);
            comm.Parameters.AddWithValue("@type", comboBox1.Text[0]);
            comm.Parameters.AddWithValue("@shortDesc", textBox5.Text);
            comm.Parameters.AddWithValue("@notes", textBox6.Text);
            comm.Parameters.AddWithValue("@curUser", PublicFuncsNvars.curUser.userCode);
            comm.Parameters.AddWithValue("@curDate", DateTime.Today.ToString("yyyyMMdd"));
            comm.Parameters.AddWithValue("@curTime", DateTime.Now.TimeOfDay.ToString("hhmmss"));
            comm.Parameters.AddWithValue("@classification", PublicFuncsNvars.getClassificationCode(comboBox2.Text));
            comm.Parameters.AddWithValue("@isActive", checkBox1.Checked);

            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();

            var f = PublicFuncsNvars.folders.Find(x => x.id == int.Parse(textBox8.Text));
            f.description = textBox4.Text;
            f.shortDescription = textBox5.Text;
            f.classification = (Classification)PublicFuncsNvars.getClassificationCode(comboBox2.Text);
            f.isActive = checkBox1.Checked;

            panel2.Visible = false;
            panel1.Visible = true;

            MessageBox.Show("תיק מספר " + f.id + ": " + textBox4.Text + " עודכן בהצלחה",
                            "אישור עדכון תיק", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            dataGridViewFolders.Rows.Clear();
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "שו\"ש")
            {
                customLabel3.Visible = true;
                textBox11.Visible = true;
                textBox11.Text = "";
            }
            else
            {
                customLabel3.Visible = false;
                textBox11.Visible = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int res = -1;
            List<Folder> checkedFolders=checkedListBox1.CheckedItems.Cast<Folder>().ToList();
            if (PublicFuncsNvars.shortDescExists(textBox10.Text)&&!shortDescInCheckedItems(checkedFolders, textBox10.Text))
            {
                MessageBox.Show("שם מקוצר זה כבר קיים עבור תיק אחר, אנא בחרו שם מקוצר אחר",
                                "שם מקוצר קיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (textBox10.Text == "" || textBox9.Text == "" || comboBox4.Text == "")
            {
                MessageBox.Show("אין להשאיר את השדות 'שם מקוצר', 'שם תיק', 'סוג תיק' ו-'סיווג' ריקים",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (comboBox4.Text == "שו\"ש" && (!int.TryParse(textBox11.Text, out res) || res <= 0))
            {
                MessageBox.Show("אין ליצור תיק מסוג שו\"ש בלי מספר שו\"ש או עם מספר שו\"ש לא חיובי",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("INSERT INTO dbo.tm_mesimot (ms_mshimh, shm_mshimh, shm_mkotzr, sog_mshimh, ms_mshimt_am, mkt, "
                    + "simn, ms_archh_shosh, msd_proiikt, is_tik, is_tik_pail, anp, harot_mshimh, iishom_shosh_isdrti, aomdn_shosh, aomdn_shosh_o, "
                    + "mtba_aomdn_shosh, alot_ishom_lin_bnih, aishom_liin_bnih, mtba_aishom_lin_b, alot_btzoa_tnk_atdi, abtzoa_tnk_atidi, mtba_abtzoa_tatdi, "
                    + "mshimh_rgishh, User_Update, Date_Update, Time_Update, User_Create, Date_Create, Time_Create, ms_pitoh, sog_ptoh_shosh_Obsolete, "
                    + "ailn_iohsin_Obsolete, rmt_ailn_Obsolete, is_Print_Tag_Obsolete, kod_sioog, kod_mtzb_mshimh, kod_adipot, tarich_hthlh_mkori, "
                    + "tarich_hthlh_mtochnn, tarich_hthlh_bpoal, tarich_siom_mkori, tarich_siom_mtochnn, tarich_siom_bpoal, mspr_mkori, hzmnh_miohdt, aopion, "
                    + "SOW, shrtoti_rpt, B_F, tik_shrtotim, mshkl_tchni, mshkl_chlchli, mshkl_loz, mshkl_niholi, tzion_sp_tchni, tzion_sp_niholi, mspr_pitoh) "
                    + Environment.NewLine + "OUTPUT INSERTED.ms_mshimh" + Environment.NewLine
                    + "VALUES((SELECT MAX(ms_mshimh) FROM dbo.tm_mesimot)+1, @description, @shortDesc, @type, 0, 0, '', @shoshNum, 0, 1, 1, @branch, @notes, '', 0, '', '', 0, '', '', 0, '', '', 0, "
                    + "@creatingUser, @creationDate, @creationTime, @creatingUser, @creationDate, @creationTime, 0, '', '', 0, 0, @classification, 1, 0, "
                    + "'00000000', '00000000', @creationDate, '00000000', '00000000', '00000000', 0, 0, '', '', '', '', '', 0, 0, 0, 0, 0, 0, '')", conn);
                comm.Parameters.AddWithValue("@description", textBox9.Text);
                comm.Parameters.AddWithValue("@shortDesc", textBox10.Text);
                comm.Parameters.AddWithValue("@type", comboBox4.Text[0]);
                comm.Parameters.AddWithValue("@shoshNum", res == -1 ? 0 : res);
                comm.Parameters.AddWithValue("@branch", (char)PublicFuncsNvars.curUser.branch);
                comm.Parameters.AddWithValue("@notes", textBox12.Text);
                comm.Parameters.AddWithValue("@creatingUser", PublicFuncsNvars.curUser.userCode);
                comm.Parameters.AddWithValue("@creationDate", DateTime.Today.ToString("yyyyMMdd"));
                comm.Parameters.AddWithValue("@creationTime", DateTime.Now.TimeOfDay.ToString("hhmmss"));
                short clas = getMaxClassification(checkedFolders);
                comm.Parameters.AddWithValue("@classification", clas);



                conn.Open();
                int id = (int)comm.ExecuteScalar();
                conn.Close();

                if (res == -1)
                    PublicFuncsNvars.folders.Add(new Folder(id, textBox9.Text, textBox10.Text, true, PublicFuncsNvars.curUser.branch, true, (FileType)(comboBox4.Text[0]),
                                   (Classification)clas));
                else
                    PublicFuncsNvars.folders.Add(new ShoshFolder(id, textBox9.Text, textBox10.Text, true, PublicFuncsNvars.curUser.branch, true, (FileType)(comboBox4.Text[0]),
                                        (Classification)clas, res));
                panel4.Visible = true;
                moveDocsToNewFolder(checkedFolders, id);
                panel4.Visible = false;

                MessageBox.Show("תיק מספר " + id + ": " + textBox9.Text + " נוצר בהצלחה",
                                "אישור יצירת תיק", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                textBox9.Clear();
                textBox10.Clear();
                comboBox4.SelectedItem = null;
                textBox12.Clear();
                textBox11.Clear();
                panel3.Visible = false;
                panel1.Visible = true;
            }
        }

        private void moveDocsToNewFolder(List<Folder> checkedFolders, int newFolderId)
        {
            Cursor.Current = Cursors.WaitCursor;
            customLabel6.Refresh();
            customLabel7.Refresh();
            Dictionary<int, bool> docIds = new Dictionary<int, bool>();
            progressBar1.Maximum = checkedFolders.Count + 1;
            foreach (Folder f in checkedFolders)
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT shotef_klali, is_rashi FROM dbo.tiukim WHERE mispar_nose=@folderId", conn);
                comm.Parameters.AddWithValue("@folderId", f.id);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                while (sdr.Read())
                {
                    int docId = sdr.GetInt32(0);
                    bool isMain = sdr.GetBoolean(1);
                    if (docIds.ContainsKey(docId))
                    {
                        if (!docIds[docId] && isMain)
                            docIds[docId] = isMain;
                    }
                    else
                        docIds.Add(docId, isMain);
                }
                progressBar1.Value += 2;
                progressBar1.Value -= 1;
                progressBar1.CreateGraphics().DrawString((((double)progressBar1.Value / (double)(progressBar1.Maximum - 1)) * 100).ToString("N1") + "%", new Font("Arial",
                    9.75f, FontStyle.Bold), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
            }
            progressBar1.Maximum -= 1;
            progressBar1.CreateGraphics().DrawString("100%", new Font("Arial", 9.75f, FontStyle.Bold), Brushes.Black, new PointF(progressBar1.Width / 2 - 10,
                progressBar1.Height / 2 - 7));

            progressBar2.Maximum = docIds.Count + 1;
            foreach (KeyValuePair<int, bool> doc in docIds)
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("INSERT INTO dbo.tiukim(kod_marechet, shotef_klali, mispar_nose, is_rashi, mispar_in_tik)" + Environment.NewLine +
                    "OUTPUT inserted.mispar_in_tik" + Environment.NewLine +
                    " VALUES(2, @id, @directory, @isPrimery, (SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                    "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@directory)+1)", conn);
                comm.Parameters.AddWithValue("@id", doc.Key);
                comm.Parameters.AddWithValue("@directory", newFolderId);
                comm.Parameters.AddWithValue("@isPrimery", doc.Value);
                conn.Open();
                int numberInFolder = (int)comm.ExecuteScalar();
                conn.Close();
                progressBar2.Value += 2;
                progressBar2.Value -= 1;
                progressBar2.CreateGraphics().DrawString((((double)progressBar2.Value / (double)(progressBar2.Maximum - 1)) * 100).ToString("N1") + "%", new Font("Arial",
                    9.75f, FontStyle.Bold), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));
                if(doc.Value)
                {
                    string newRefferences = textBox10.Text + " - " + numberInFolder + " - " + doc.Key;

                    comm = new SqlCommand("SELECT simuchin FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
                    comm.Parameters.AddWithValue("@docId", doc.Key);
                    conn.Open();
                    string refferences = comm.ExecuteScalar().ToString();
                    conn.Close();

                    comm = new SqlCommand("UPDATE dbo.documents SET simuchin=@refferences WHERE shotef_mismach=@docId", conn);
                    comm.Parameters.AddWithValue("@docId", doc.Key);
                    comm.Parameters.AddWithValue("@refferences", newRefferences);
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    string s = PublicFuncsNvars.getDocExt(doc.Key);
                    if (s != null && (PublicFuncsNvars.getDocExt(doc.Key).ToLower() == "doc" || PublicFuncsNvars.getDocExt(doc.Key).ToLower() == "docx"))
                        PublicFuncsNvars.updateRefferencesInWordDoc(doc.Key, refferences, newRefferences);
                }
            }
            progressBar2.Maximum -= 1;
            progressBar2.CreateGraphics().DrawString("100%", new Font("Arial", 9.75f, FontStyle.Bold), Brushes.Black, new PointF(progressBar2.Width / 2 - 10,
                progressBar2.Height / 2 - 7));

            foreach (Folder f in checkedFolders)
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("DELETE FROM dbo.tm_mesimot WHERE ms_mshimh=@folderId", conn);
                comm.Parameters.AddWithValue("@folderId", f.id);
                conn.Open();
                comm.ExecuteNonQuery();
            }

            PublicFuncsNvars.getFolders();
            comboBox5.SelectedIndex = -1;
            comboBox5.SelectedIndex = 0;

            Cursor.Current = Cursors.Default;
        }

        private short getMaxClassification(List<Folder> checkedFolders)
        {
            Classification c = Classification.unknown;
            foreach (Folder f in checkedFolders)
                if (f.classification > c)
                    c = f.classification;
            return (short)c;
        }

        private bool shortDescInCheckedItems(List<Folder> checkedFolders, string shortDesc)
        {
            foreach (Folder f in checkedFolders)
                if (f.shortDescription == shortDesc)
                    return true;
            return false;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            foldersView.Clear();
            switch (comboBox5.Text)
            {
                case "הכל":
                    foldersView.AddRange(PublicFuncsNvars.folders);
                    break;
                case "לשכה":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x=>x.branch==Branch.office));
                    break;
                case "פרוייקטים":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.projects));
                    break;
                case "תקציבים":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.budgets));
                    break;
                case "מחשוב":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.computers));
                    break;
                case "ארגון":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.organization));
                    break;
                case "ייצור":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.manufacturing));
                    break;
                case "פיתוח":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.development));
                    break;
                case "סייף":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.sayaf));
                    break;
                case "חלקה":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.chelka));
                    break;
                case "אחר":
                    foldersView.AddRange(PublicFuncsNvars.folders.Where(x => x.branch == Branch.other));
                    break;
            }
            checkedListBox1.DataSource = null;
            checkedListBox1.DataSource = foldersView;
            checkedListBox1.DisplayMember = "description";
            checkedListBox1.ValueMember = "id";
            checkedListBox1.Select();
        }

        delegate void invokeMethod();

        void updateNotes()
        {
            if (checkedListBox1.CheckedItems.Count > 0)
            {
                string s = "מורכב מאיחוד התיקים הבאים:" + Environment.NewLine;
                for (int i = 0; i < checkedListBox1.CheckedItems.Count - 1; i++)
                {
                    Folder f = (Folder)checkedListBox1.CheckedItems[i];
                    s += "תיק מספר: " + f.id + " - " + f.description + "," + Environment.NewLine;
                }
                Folder fl = (Folder)checkedListBox1.CheckedItems[checkedListBox1.CheckedItems.Count - 1];
                s += "תיק מספר: " + fl.id + " - " + fl.description;
                textBox12.Text = s;
            }
            else
                textBox12.Clear();
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke(new invokeMethod(updateNotes));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel3.Visible = false;
            textBox9.Clear();
            textBox10.Clear();
            comboBox4.SelectedItem = null;
            textBox12.Clear();
            textBox11.Clear();
            panel3.Visible = false;
            panel1.Visible = true;
        }
    }
}
