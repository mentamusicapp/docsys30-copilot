using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentsModule
{
    public partial class SearchPatternScreen : Form
    {
        string conStr = Global.ConStr;
        bool newOrUpdate;
        string name;
        List<User> users;
        bool okSender = false, okRef = false, okBro = false, okDir = false, okFromPro = false, okToPro = false;
        int row1 = 0, row2 = 0, row3 = 0;
        string strTyped = "";
        Dictionary<int, string> projects;
        List<Folder> directories;

        public SearchPatternScreen(bool newOrUpdate, string name)
        {
            this.newOrUpdate = newOrUpdate;
            this.name = name;
            users = PublicFuncsNvars.users;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conStr);
            SqlCommand comm = new SqlCommand();
            comm.Connection = conn;
            if(newOrUpdate)
            {
                comm.CommandText = "INSERT INTO dbo.docSearchPatterns (userCode, pattName, fromID, toID, fromDate"+
                    ", toDate, subject, senderCode, senderFirstName, senderLastName, senderRole, referencedCode, referencedRole," +
                    " directoryCode, branch, isPublished, isActive, docContent, docOrDirect, isForAct, fromProject, toProject) VALUES(@userCode, @pattName,"+
                    " @fromID, @toID, @fromDate, @toDate, @subject, @senderCode, @senderFirstName, @senderLastName, @senderRole, @referencedCode, @referencedRole," +
                    " @directoryCode, @branch, @isPublished, @isActive, @docContent, @docOrDirect, @isForAct, @fromProject, @toProject)";
            }
            else
            {
                comm.CommandText = "UPDATE dbo.docSearchPatterns SET pattName=@pattName, fromID=@fromID, toID=@toID, fromDate=" +
                    "@fromDate, toDate=@toDate, subject=@subject, senderCode=@senderCode, senderFirstName=@senderFirstName, senderLastName=@senderLastName, senderRole=@senderRole," +
                    " referencedCode=@referencedCode, referencedRole=@referencedRole,"+
                    " directoryCode=@directoryCode, branch=@branch, isPublished=@isPublished, isActive=@isActive,"+
                    " docContent=@docContent, docOrDirect=@docOrDirect, isForAct=@isForAct, fromProject=@fromProject, toProject=@toProject" +
                    " WHERE userCode=@userCode AND pattName=@pn";
                comm.Parameters.AddWithValue("@pn", name);
            }
            comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
            comm.Parameters.AddWithValue("@pattName", textBox20.Text);
            if (!checkBox3.Checked)
            {
                comm.Parameters.AddWithValue("@fromID", int.Parse(textBox1.Text));
                comm.Parameters.AddWithValue("@toID", int.Parse(textBox2.Text));
            }
            else
            {
                comm.Parameters.AddWithValue("@fromID", 0);
                comm.Parameters.AddWithValue("@toID", 0);
            }
            if (!checkBox2.Checked)
            {
                comm.Parameters.AddWithValue("@fromDate", dateTimePicker2.Value.ToShortDateString());
                comm.Parameters.AddWithValue("@toDate", dateTimePicker1.Value.ToShortDateString());
            }
            else
            {
                comm.Parameters.AddWithValue("@fromDate", "");
                comm.Parameters.AddWithValue("@toDate", "");
            }
            comm.Parameters.AddWithValue("@subject", textBox5.Text);
            int res7;
            if (int.TryParse(textBox7.Text, out res7))
                comm.Parameters.AddWithValue("@senderCode", res7);
            else
                comm.Parameters.AddWithValue("@senderCode", 0);
            comm.Parameters.AddWithValue("@senderFirstName", textBox3.Text);
            comm.Parameters.AddWithValue("@senderLastName", textBox21.Text);
            comm.Parameters.AddWithValue("@senderRole", textBox8.Text);
            int res10;
            if (int.TryParse(textBox10.Text, out res10))
                comm.Parameters.AddWithValue("@referencedCode", res10);
            else
                comm.Parameters.AddWithValue("@referencedCode", 0);
            comm.Parameters.AddWithValue("@referencedRole", textBox9.Text);
            
            int res12;
            if (int.TryParse(textBox12.Text, out res12))
                comm.Parameters.AddWithValue("@directoryCode", res12);
            else
                comm.Parameters.AddWithValue("@directoryCode", 0);
            comm.Parameters.AddWithValue("@branch", comboBox3.Text);
            comm.Parameters.AddWithValue("@isPublished", comboBox4.Text);
            comm.Parameters.AddWithValue("@isActive", checkBox1.Checked);
            comm.Parameters.AddWithValue("@docContent", textBox17.Text);
            comm.Parameters.AddWithValue("@docOrDirect", comboBox2.Text);
            comm.Parameters.AddWithValue("@isForAct", comboBox1.Text);
            int res14, res13;
            
            if(int.TryParse(textBox14.Text, out res14))
                comm.Parameters.AddWithValue("@fromProject", res14);
            else
                comm.Parameters.AddWithValue("@fromProject", 0);
            if (int.TryParse(textBox14.Text, out res13))
                comm.Parameters.AddWithValue("@toProject", res13);
            else
                comm.Parameters.AddWithValue("@toProject", 0);
            
            try
            {
                conn.Open();
                comm.ExecuteNonQuery();
                name = textBox20.Text;
                MessageBox.Show("התבנית נשמרה בהצלחה");
            }
            catch(Exception ex)
            {
                PublicFuncsNvars.saveLogError(FindForm().Name, ex.ToString(), ex.Message);
                MessageBox.Show("לא ניתן להשלים את הפעולה, בדוק אם שם תבנית זהה כבר קיים במערכת עבורך.");
            }
            finally
            {
                conn.Close();
            }
            Program.dm.ds.reloadPatterns();
            this.Close();
        }

        private void SearchPatternScreen_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            dateTimePicker2.Value = DateTime.Today.AddMonths(-1);
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
            
            users = PublicFuncsNvars.users;
            foreach (User u in users)
                dataGridView3.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
            textBox7.Text = PublicFuncsNvars.curUser.userCode.ToString();
            comboBox3.Text = PublicFuncsNvars.getBranchString(PublicFuncsNvars.curUser.branch);

            directories = PublicFuncsNvars.folders;
            foreach (Folder d in directories)
                dataGridView2.Rows.Add(d.id, d.shortDescription, d.description, PublicFuncsNvars.getBranchString(d.branch));

            projects = PublicFuncsNvars.projects;
            foreach (KeyValuePair<int, string> p in projects)
                dataGridView4.Rows.Add(p.Key.ToString(), p.Value);
            if(!newOrUpdate)
            {
                textBox20.Text = name;
                SqlConnection conn = new SqlConnection(conStr);
                SqlCommand comm = new SqlCommand("SELECT fromID, toID, fromDate" +
                    ", toDate, subject, senderCode, senderFirstName, senderRole, referencedCode, referencedRole,"+
                    " directoryCode, branch, isPublished, isActive, docContent, docOrDirect, isForAct, fromProject, toProject, senderLastName FROM dbo.docSearchPatterns WHERE" +
                    " userCode=@userCode AND pattName=@pattName", conn);
                comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin).userCode);
                comm.Parameters.AddWithValue("@pattName", name);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if(sdr.Read())
                {
                    textBox1.Text = sdr.GetInt32(0).ToString();
                    textBox2.Text = sdr.GetInt32(1).ToString();
                    if (textBox1.Text.Equals("0") && textBox2.Text.Equals("0"))
                        checkBox3.Checked = true;
                    if (!sdr.GetString(2).Trim().Equals(""))
                    {
                        checkBox2.Checked = false;
                        string[] fromDate = sdr.GetString(2).Trim().Split('/');
                        dateTimePicker2.Value = new DateTime(int.Parse(fromDate[2]), int.Parse(fromDate[1]), int.Parse(fromDate[0]));
                        string[] toDate = sdr.GetString(3).Trim().Split('/');
                        dateTimePicker1.Value = new DateTime(int.Parse(toDate[2]), int.Parse(toDate[1]), int.Parse(toDate[0]));
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
                        textBox21.Text = sdr.GetString(19).Trim();
                    }
                    if (rc != 0)
                        textBox10.Text = rc.ToString();
                    else
                        textBox9.Text = sdr.GetString(9).Trim();
                    if (dc != 0)
                        textBox12.Text = dc.ToString();
                    comboBox3.Text = sdr.GetString(11).Trim();
                    comboBox4.Text = sdr.GetString(12).Trim();
                    checkBox1.Checked = sdr.GetBoolean(13);
                    textBox17.Text = sdr.GetString(14).Trim();
                    comboBox2.Text = sdr.GetString(15).Trim();
                    comboBox1.Text = sdr.GetString(16).Trim();
                    textBox14.Text = sdr.GetInt32(17).ToString();
                    textBox13.Text = sdr.GetInt32(18).ToString();
                }
                conn.Close();
            }
            else
            {
                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
                comboBox4.SelectedIndex = 0;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Text.Equals(""))
            {
                textBox8.Text = "";
                textBox3.Text = "";
                textBox21.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox7.Text.Equals("קוד"))
            {
                textBox8.Text = "תפקיד";
                textBox3.Text = "שם פרטי";
                textBox21.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            else
            {
                int res;
                if (int.TryParse(textBox7.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            textBox8.Text = u.job;
                            textBox3.Text = u.firstName;
                            textBox21.Text = u.lastName;
                            comboBox3.Text =PublicFuncsNvars.getBranchString(u.branch);
                            break;
                        }
                    }
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox7.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text.Equals(""))
                textBox9.Text = "";
            else if (textBox10.Text.Equals("קוד"))
            {
                textBox9.Text = "תפקיד";
                comboBox1.Text = "הכל";
            }
            else
            {
                int res;
                if (int.TryParse(textBox10.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            textBox9.Text = u.job;
                            break;
                        }
                    }
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox10.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            textBox11.TextChanged -= textBox11_TextChanged;
            TextBox tb = null;
            PublicFuncsNvars.directoryByCode(ref textBox12, ref textBox11, ref textBox4, ref tb, "קוד", "שם מקוצר",
                "SELECT shm_mshimh, shm_mkotzr FROM dbo.tm_mesimot WHERE ms_mshimh=@id AND shm_mkotzr<>''", "@id", typeof(int));
            textBox11.TextChanged += textBox11_TextChanged;
            if (!textBox12.Text.Equals("קוד") && !textBox12.Text.Equals(""))
            {
                int index = 0;
                dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().StartsWith(textBox12.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            textBox12.TextChanged -= textBox12_TextChanged;
            TextBox tb = null;
            PublicFuncsNvars.directoryByCode(ref textBox11, ref textBox12, ref textBox4, ref tb, "שם מקוצר", "קוד",
                "SELECT shm_mshimh, ms_mshimh FROM dbo.tm_mesimot WHERE shm_mkotzr=@shortName", "@shortName", typeof(string));
            textBox12.TextChanged += textBox12_TextChanged;
            if (!textBox11.Text.Equals("שם מקוצר") && !textBox11.Text.Equals(""))
            {
                int index = 0;
                dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString().StartsWith(textBox11.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView2.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                dateTimePicker1.ResetText();
                dateTimePicker2.Value = DateTime.Today.AddMonths(-1);
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
            else
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                textBox1.Clear();
                textBox2.Clear();
                textBox1.Enabled = false;
                textBox2.Enabled = false;
            }
            else
            {
                textBox1.Enabled = true;
                textBox2.Enabled = true;
            }
        }

        private void SearchPatternScreen_FormClosed(object sender, FormClosedEventArgs e)
        {
            Program.dm.ds.sps = null;
            Program.dm.ds.Show();
        }

        private void dataGridView3_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row3 = e.RowIndex;
                dataGridView3.Rows[row3].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            string value = dataGridView3.Rows[dataGridView3.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            if (okSender)
                textBox7.Text = value;
            else if (okRef)
                textBox10.Text = value;
        }

        private void dataGridView3_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridView3.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
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

        private void button3_Click(object sender, EventArgs e)
        {
            if (okSender)
                textBox7.Text = PublicFuncsNvars.curUser.userCode.ToString();
            else if (okRef)
                textBox10.Text = "קוד";
            dataGridView3.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView3.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
        }

        private void selectingSender()
        {
            dataGridView3.Visible = true;
            button3.Visible = true;
            button5.Visible = true;
            dataGridView2.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            okSender = true;
            okRef = false;
            okBro = false;
            makeProjectsTableInVisible();
        }

        private void textBox7_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void makeProjectsTableInVisible()
        {
            Control[] controls = { dataGridView4, button10, button11 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void textBox7_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            int res;
            if (!int.TryParse(textBox7.Text, out res))
            {
                textBox7.Text = "";
            }
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            if (textBox7.Text.Equals(""))
                textBox7.Text = "קוד";
        }

        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

        private void selectingRecipient()
        {
            dataGridView3.Visible = true;
            button3.Visible = true;
            button5.Visible = true;
            dataGridView2.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            okSender = false;
            okRef = true;
            okBro = false;
            makeProjectsTableInVisible();
        }

        private void textBox10_Click(object sender, EventArgs e)
        {
            selectingRecipient();
        }

        private void textBox19_Click(object sender, EventArgs e)
        {
            dataGridView3.Visible = true;
            button3.Visible = true;
            button5.Visible = true;
            dataGridView2.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
            okSender = false;
            okRef = false;
            okBro = true;
            dataGridView3.Focus();
            makeProjectsTableInVisible();
        }

        private void textBox12_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = true;
            button2.Visible = true;
            button4.Visible = true;
            dataGridView3.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
            okDir = true;
            makeProjectsTableInVisible();
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = true;
            button2.Visible = true;
            button4.Visible = true;
            dataGridView3.Visible = false;
            button3.Visible = false;
            button5.Visible = false;
            okDir = true;
            dataGridView2.Focus();
        }

        private void textBox14_Click(object sender, EventArgs e)
        {
            strTyped = "";
            dataGridView4.Visible = true;
            button10.Visible = true;
            button11.Visible = true;
            makeUsersTableVisibleOrInvisible(false);
            makeDirectoriesTableInVisible();
            okFromPro = true;
            okToPro = false;
        }

        private void textBox13_Click(object sender, EventArgs e)
        {
            strTyped = "";
            dataGridView4.Visible = true;
            button10.Visible = true;
            button11.Visible = true;
            makeUsersTableVisibleOrInvisible(false);
            makeDirectoriesTableInVisible();
            okFromPro = false;
            okToPro = true;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int res;
                if (int.TryParse(textBox14.Text, out res))
                    textBox15.Text = projects[res];
                else
                    textBox15.Text = "שם פרויקט";
            }
            catch { }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int res;
                if (int.TryParse(textBox13.Text, out res))
                    textBox16.Text = projects[res];
                else
                    textBox16.Text = "שם פרויקט";
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox12.Text = "קוד";
            dataGridView2.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            button2.Visible = false;
            button4.Visible = false;
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

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (okFromPro)
                textBox14.Text = dataGridView4.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
            else if (okToPro)
                textBox13.Text = dataGridView4.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }

        private void dataGridView4_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridView4.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridView4.Rows)
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

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView4_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row1 = e.RowIndex;
                dataGridView4.Rows[row1].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (okDir)
                textBox12.Text = dataGridView2.Rows[dataGridView2.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
        }

        private void dataGridView2_Leave(object sender, EventArgs e)
        {

        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridView2.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    row.Cells[col].Selected = true;
                    break;
                }
            }
        }

        private void dataGridView2_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void dataGridView2_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                row2 = e.RowIndex;
                dataGridView2.Rows[row2].Cells[e.ColumnIndex].Selected = true;
            }
        }

        private void makeDirectoriesTableInVisible()
        {
            Control[] controls = { dataGridView2, button2, button4 };
            PublicFuncsNvars.changeControlsVisiblity(false, controls.ToList());
        }

        private void makeUsersTableVisibleOrInvisible(bool b)
        {
            Control[] controls = { dataGridView3, button3, button5 };
            PublicFuncsNvars.changeControlsVisiblity(b, controls.ToList());
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

        private void textBox11_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            dataGridView2.Visible = true;
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

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker1.MinDate = dateTimePicker2.Value;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MaxDate = dateTimePicker1.Value;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (textBox8.Text.Equals(""))
            {
                textBox7.Text = "";
                textBox3.Text = "";
                textBox21.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox8.Text.Equals("תפקיד"))
            {
                textBox7.Text = "קוד";
                textBox3.Text = "שם פרטי";
                textBox21.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[1], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString().StartsWith(textBox8.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
            {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox21.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox3.Text.Equals("שם פרטי"))
            {
                textBox7.Text = "קוד";
                textBox8.Text = "תפקיד";
                textBox21.Text = "שם משפחה";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[1], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString().StartsWith(textBox3.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            if (textBox21.Text.Equals(""))
            {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox3.Text = "";
                comboBox3.Text = "הכל";

            }
            else if (textBox21.Text.Equals("שם"))
            {
                textBox7.Text = "קוד";
                textBox8.Text = "תפקיד";
                textBox3.Text = "שם פרטי";
                comboBox3.Text = "הכל";
            }
            else
            {
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[2], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[2].Value != null && row.Cells[2].Value.ToString().StartsWith(textBox21.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
                textBox10.Text = "";
            else if (textBox9.Text.Equals("תפקיד"))
            {
                textBox10.Text = "קוד";
                comboBox1.Text = "הכל";
            }
            else
            {
                int index = 0;
                dataGridView3.Sort(dataGridView3.Columns[3], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells[3].Value != null && row.Cells[3].Value.ToString().StartsWith(textBox9.Text))
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView3.FirstDisplayedScrollingRowIndex = index;
            }
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            textBox8.Text = "";
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            if (textBox8.Text.Equals(""))
                textBox8.Text = "תפקיד";
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            textBox3.Text = "";
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
                textBox3.Text = "שם פרטי";
        }

        private void textBox21_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            textBox21.Text = "";
        }

        private void textBox21_Leave(object sender, EventArgs e)
        {
            if (textBox21.Text.Equals(""))
                textBox21.Text = "שם משפחה";
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            strTyped = "";
            textBox9.Text = "";
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
                textBox9.Text = "תפקיד";
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

        private void textBox8_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox21_Click(object sender, EventArgs e)
        {
            selectingSender();
        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            selectingRecipient();
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }
    }
}
