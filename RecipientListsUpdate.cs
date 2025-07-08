using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentsModule
{
    public partial class RecipientListsUpdate : Form
    {
        RecipientList rl;
        bool isJustOne;
        bool okAddUser = false;
        string strTyped = "";

        public RecipientListsUpdate(RecipientList rl, bool isJustOne)
        {
            this.rl = rl;
            this.isJustOne = isJustOne;
            InitializeComponent();
        }

        private void ShowRecipientList_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            foreach (Recipient r in PublicFuncsNvars.interDist)
                dataGridViewInterDist.Rows.Add(false, r.getId(), r.getRole(), r.getEmail());
            dataGridViewListContents.CellValueChanged += dataGridViewListContents_CellValueChanged;
            comboBox2.SelectedIndex = 0;
            textBox2.Text = PublicFuncsNvars.curUser.userCode.ToString();
            dataGridViewUsers.SelectionChanged -= dataGridViewUsers_SelectionChanged;
            foreach (User u in PublicFuncsNvars.users.Where(X=>X.isActive).ToList())
                dataGridViewUsers.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
            dataGridViewUsers.SelectionChanged += dataGridViewUsers_SelectionChanged;
            if (isJustOne)
            {
                button8.Visible = false;
                textBox1.Text = rl.name;
                textBox2.Text = rl.owner.ToString();
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                comboBox1.Text = PublicFuncsNvars.getBranchString(rl.branch);
                comboBox3.Text = rl.getLevelString();
                comboBox1.Enabled = false;
                comboBox3.Enabled = false;
                button2.PerformClick();
            }
            DocumentsMenu.PathTemplate(this.button1, 30);
            DocumentsMenu.PathTemplate(this.button2, 30);
            DocumentsMenu.PathTemplate(this.button3, 30);
            DocumentsMenu.PathTemplate(this.button4, 30);
            DocumentsMenu.PathTemplate(this.button5, 30);
            DocumentsMenu.PathTemplate(this.button6, 30);
            DocumentsMenu.PathTemplate(this.button7, 30);
            DocumentsMenu.PathTemplate(this.button8, 30);
            DocumentsMenu.PathTemplate(this.button9, 30);
            DocumentsMenu.PathTemplate(this.button10, 30);
            DocumentsMenu.PathTemplate(this.button11, 30);
            DocumentsMenu.PathTemplate(this.button12, 30);
            DocumentsMenu.PathTemplate(this.button13, 30);
            DocumentsMenu.PathTemplate(this.button14, 30);
            DocumentsMenu.PathTemplate(this.button25, 30);
            DocumentsMenu.PathTemplate(this.button26, 30);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridViewRecipientLists.Rows.Clear();
            int res=0;
            if (textBox2.Text == "" || textBox2.Text == "קוד" || int.TryParse(textBox2.Text, out res))
            {
                foreach (KeyValuePair<KeyValuePair<short, short>, RecipientList> recKVP in PublicFuncsNvars.recipientsLists)
                {
                    RecipientList recl = recKVP.Value;
                    if (recl.owner == PublicFuncsNvars.curUser.userCode ||
                       (recl.branch == PublicFuncsNvars.curUser.branch && recl.level == RecipientListsLevel.branch) ||
                        recl.level == RecipientListsLevel.unit||PublicFuncsNvars.curUser.roleType==RoleType.computers)
                    {
                        if ((textBox2.Text == "" || textBox2.Text == "קוד" || recl.owner == res) && recl.name.Contains(textBox1.Text) &&
                             (recl.getLevelString() == comboBox3.Text||comboBox3.Text==""||comboBox3.Text=="הכל")
                             && (PublicFuncsNvars.getBranchString(recl.branch) == comboBox1.Text||comboBox1.Text==""||comboBox1.Text=="הכל"))
                        {
                            dataGridViewRecipientLists.Rows.Add(recKVP.Key.Key, recKVP.Key.Value, recl.name, recl.getLevelString(), recl.getOwnerName());
                        }
                    }
                }
                if (dataGridViewRecipientLists.Rows.Count > 0)
                    dataGridViewRecipientLists.Rows[0].Selected = true;
                else
                    MessageBox.Show("אין רשימות תפוצה המתאימות לחיפוש", "חיפוש רשימות תפוצה",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void dataGridViewRecipientLists_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewRecipientLists.SelectedRows.Count > 0)
            {
                if (dataGridViewRecipientLists.SelectedRows[0].Index >= 0)
                {
                    if (!isJustOne)
                    {
                        DataGridViewRow row = dataGridViewRecipientLists.SelectedRows[0];
                        KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                        rl = PublicFuncsNvars.recipientsLists[kvp];
                    }
                    dataGridViewListContents.Rows.Clear();
                    dataGridViewListContents.CellValueChanged -= dataGridViewListContents_CellValueChanged;
                    foreach (Recipient r in rl.getRecipients())
                    {
                        int rowIndex = dataGridViewListContents.Rows.Add(r.getId(), r.getId() != 99999 ? PublicFuncsNvars.getUserNameByUserCode(r.getId()) : r.getRole(), r.getRole());
                        dataGridViewListContents.Rows[rowIndex].Cells[3].Value = r.getIFA() ? "לפעולה" : "לידיעה";
                    }
                    dataGridViewListContents.CellValueChanged += dataGridViewListContents_CellValueChanged;
                    if (dataGridViewListContents.Rows.Count > 0)
                        dataGridViewListContents.Rows[0].Selected = true;
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            int res;
            if(textBox2.Text=="")
            {
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                comboBox1.SelectedItem = null;
            }
            else if(textBox2.Text=="קוד")
            {
                textBox3.Text = "שם פרטי";
                textBox4.Text = "שם משפחה";
                textBox5.Text = "תפקיד";
                comboBox1.SelectedItem = null;
            }
            else if(int.TryParse(textBox2.Text, out res))
            {
                if(PublicFuncsNvars.userCodeExists(res))
                {
                    User u = PublicFuncsNvars.getUserByCode(res);
                    textBox3.Text = u.firstName;
                    textBox4.Text = u.lastName;
                    textBox5.Text = u.job;
                    comboBox1.Text = PublicFuncsNvars.getBranchString(u.branch);
                }
            }

        }

        private void dataGridViewUsers_SelectionChanged(object sender, EventArgs e)
        {
            if (!okAddUser)
                textBox2.Text = dataGridViewUsers.SelectedCells[0].OwningRow.Cells[0].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            okAddUser = true;
            panel6.Visible = false;
            panel4.Visible = false;
            panel1.Visible = true;
            comboBox2.Visible = true;
            button10.Visible = true;
            button5.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            okAddUser = false;
            button3.Visible = true;
            button9.Visible = true;
            button2.Visible = false;
            button1.Visible = false;
            button4.Visible = false;
            button11.Visible = false;
            button12.Visible = false;
            button5.Visible = false;
            button10.Visible = false;
            comboBox2.Visible = false;
            textBox2.Text = PublicFuncsNvars.curUser.userCode.ToString();
            textBox2.Enabled = false;
            comboBox1.Enabled = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            okAddUser = false;
            panel6.Visible = false;
            panel4.Visible = false;
            panel1.Visible = false;
            comboBox2.Visible = false;
            button10.Visible = false;
            button5.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridViewRecipientLists.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dataGridViewRecipientLists.SelectedRows[0];
                KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                if (PublicFuncsNvars.isAuthorizedUser(PublicFuncsNvars.recipientsLists[kvp].owner, PublicFuncsNvars.curUser))
                {
                    DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם למחוק לצמיתות את רשימת התפוצה \"" + rl.name + "\"?", "מחיקת רשימת תפוצה",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                      MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    if (res == DialogResult.Yes)
                    {
                        SqlConnection conn = new SqlConnection(Global.ConStr);
                        SqlCommand comm = new SqlCommand("DELETE FROM dbo.tm_tfuza WHERE cod_lst_tpotzh=@listCode AND cod_sys=@sysCode", conn);
                        comm.Parameters.AddWithValue("@listCode", (short)row.Cells["RListIDColumn"].Value);
                        comm.Parameters.AddWithValue("@sysCode", (short)row.Cells["sysCodeColumn"].Value);
                        conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        comm.CommandText = "DELETE FROM dbo.tm_tfuz_res WHERE cod_lst_tpotzh=@listCode AND cod_sys=@sysCode";
                        conn.Open();
                        comm.ExecuteNonQuery();
                        conn.Close();
                        PublicFuncsNvars.recipientsLists.Remove(kvp);
                        dataGridViewRecipientLists.Rows.Remove(row);
                        dataGridViewListContents.Rows.Clear();
                        MessageBox.Show("הרשימה \"" + rl.name + "\" נמחקה בהצלחה", "מחיקת רשימת תפוצה",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
                else
                {
                    MessageBox.Show("אין לך הרשאות למחוק רשימת תפוצה זו", "מחיקת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
            else
            {
                MessageBox.Show("על מנת למחוק רשימת תפוצה עליכם לסמן את הרשימה הרצויה", "מחיקת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridViewRecipientLists.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dataGridViewRecipientLists.SelectedRows[0];
                KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                if (PublicFuncsNvars.isAuthorizedUser(PublicFuncsNvars.recipientsLists[kvp].owner, PublicFuncsNvars.curUser))
                {
                    okAddUser = true;
                    button1.Visible = true;
                    button4.Visible = true;
                    button11.Visible = true;
                    button12.Visible = true;
                    button13.Visible = true;
                    button25.Visible = true;
                    button9.Visible = false;
                    button2.Visible = false;
                    button12.Visible = true;
                    comboBox1.Enabled = false;
                    button3.Visible = false;
                    textBox1.Text = row.Cells["RListNameColumn"].Value.ToString();
                    comboBox3.Text = row.Cells["RListLevelColumn"].Value.ToString();
                    textBox2.Text = PublicFuncsNvars.recipientsLists[kvp].owner.ToString();
                }
                else
                {
                    MessageBox.Show("אין לך הרשאות לעדכן רשימת תפוצה זו", "עדכון רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
            else
            {
                MessageBox.Show("על מנת לעדכן רשימת תפוצה עליכם לסמן את הרשימה הרצויה", "מחיקת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            button13.Visible = false;
            button25.Visible = false;
            button1.Visible = false;
            button4.Visible = false;
            button11.Visible = false;
            button12.Visible = false;
            button2.Visible = true;
            button9.Visible = false;
            button5.Visible = false;
            comboBox2.Visible = false;
            button10.Visible = false;
            comboBox1.Enabled = true;
            okAddUser=false;
            textBox2.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(textBox1.Text=="")
            {
                MessageBox.Show("אין ליצור רשימה ללא שם", "יצירת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else if(comboBox1.Text==""||comboBox1.Text=="הכל")
            {
                MessageBox.Show("אין ליצור רשימה ללא ענף", "יצירת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else if (comboBox3.Text == "" || comboBox3.Text == "הכל")
            {
                MessageBox.Show("אין ליצור רשימה ללא רמת רשימה", "יצירת רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else
            {
                int res2;
                if (int.TryParse(textBox2.Text, out res2))
                {
                    if (PublicFuncsNvars.userCodeExists(res2))
                    {
                        DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם ליצור את רשימת התפוצה \"" + textBox1.Text + "\"?", "יצירת רשימת תפוצה",
                                                          MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                          MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        if (res == DialogResult.Yes)
                        {
                            SqlConnection conn = new SqlConnection(Global.ConStr);
                            SqlCommand comm = new SqlCommand("INSERT INTO dbo.tm_tfuza (cod_sys, cod_lst_tpotzh, nam_lst_tpotzh, anf_tpotzh, creating_user, personalObranchOunit)"
                                                            + Environment.NewLine + "output inserted.cod_lst_tpotzh" + Environment.NewLine
                                                            + "VALUES(@sysCode, (SELECT MAX(cod_lst_tpotzh) FROM dbo.tm_tfuza WHERE cod_sys=@sysCode)+1,"
                                                            + "@listName, @branch, @creatingUser, @personalObranchOunit)", conn);
                            comm.Parameters.AddWithValue("@sysCode", 2);
                            comm.Parameters.AddWithValue("@listName", textBox1.Text);
                            char c = PublicFuncsNvars.getBranchByString(comboBox1.Text);
                            comm.Parameters.AddWithValue("@branch", c);
                            comm.Parameters.AddWithValue("@creatingUser", res2);
                            short level = (short)PublicFuncsNvars.getRecipientListLevelByString(comboBox3.Text);
                            comm.Parameters.AddWithValue("@personalObranchOunit", level);
                            conn.Open();
                            short listNum = (short)comm.ExecuteScalar();
                            conn.Close();
                            PublicFuncsNvars.recipientsLists.Add(new KeyValuePair<short, short>(2, listNum), new RecipientList(listNum, res2, level, c, textBox1.Text,
                                                                 new List<Recipient>()));
                            int rowNum = dataGridViewRecipientLists.Rows.Add((short)2, listNum, textBox1.Text, comboBox3.Text, PublicFuncsNvars.getUserNameByUserCode(res2));
                            dataGridViewRecipientLists.Rows[rowNum].Selected = true;
                            MessageBox.Show("רשימת התפוצה \"" + textBox1.Text + "\" נוצרה בהצלחה", "יצירת רשימת תפוצה",
                                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                            button3.Visible = false;
                            button9.Visible = false;
                            button2.Visible = true;
                            button12.Visible = false;
                            textBox1.Clear();
                            textBox2.Clear();
                            comboBox3.SelectedItem = null;
                            comboBox1.SelectedItem = null;
                            comboBox1.Enabled = true;
                            textBox2.Enabled = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("אין משתמש כזה במערכת", "יצירת רשימת תפוצה",
                                         MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                         MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
                else
                    MessageBox.Show("מספר משתמש לא יכול להכיל תווים שאינם ספרות, ולא יכול להיות ריק", "יצירת רשימת תפוצה",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            button12.Visible = false;
            button3.Visible = false;
            button2.Visible = true;
            textBox2.Enabled = true;
            comboBox1.Enabled = true;
            textBox1.Clear();
            textBox2.Clear();
            comboBox3.SelectedItem = null;
            comboBox1.SelectedItem = null;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridViewRecipientLists.SelectedRows.Count > 0 && dataGridViewListContents.SelectedCells.Count > 0)
            {
                DataGridViewRow rowLists = dataGridViewRecipientLists.SelectedRows[0];
                DataGridViewRow rowContents = dataGridViewListContents.SelectedCells[0].OwningRow;
                DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם להסיר את המשתמש " + rowContents.Cells[2].Value + " מרשימת התפוצה \"" + rl.name + "\"?", "הסרת משתמש מרשימת תפוצה",
                                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                  MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                if (res == DialogResult.Yes)
                {
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("DELETE FROM dbo.tm_tfuz_res WHERE cod_lst_tpotzh=@listCode AND cod_sys=@sysCode AND onum=@recipientNum", conn);
                    comm.Parameters.AddWithValue("@listCode", (short)rowLists.Cells["RListIDColumn"].Value);
                    comm.Parameters.AddWithValue("@sysCode", (short)rowLists.Cells["sysCodeColumn"].Value);
                    short nid = rl.getRecipients().Where(x => x.getId() == (int)rowContents.Cells[0].Value).ToList()[0].getNID();
                    comm.Parameters.AddWithValue("@recipientNum", rl.getRecipients().Where(x => x.getId() == (int)rowContents.Cells[0].Value).ToList()[0].getNID());
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    dataGridViewListContents.Rows.Remove(rowContents);
                    dataGridViewListContents.Rows[0].Selected = true;
                    rl.removeRecipient(nid);
                    MessageBox.Show("המשתמש \"" + rowContents.Cells[2].Value + "\" הוסר בהצלחה", "הסרה מרשימת תפוצה",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
            else
            {
                MessageBox.Show("על מנת להסיר משתמש עליכם לסמן את המשתמש בטבלת תוכן הרשימה, ואת הרשימה הרצויה", "הסרה מרשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridViewRecipientLists.SelectedRows.Count > 0 && dataGridViewUsers.SelectedCells.Count > 0)
            {
                DataGridViewRow rowLists = dataGridViewRecipientLists.SelectedRows[0];
                DataGridViewRow rowUsers = dataGridViewUsers.SelectedCells[0].OwningRow;
                DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם להוסיף את המשתמש " + rowUsers.Cells[3].Value + " לרשימת התפוצה \"" + rl.name + "\"?", "הוספת משתמש לרשימת תפוצה",
                                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                  MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                if (res == DialogResult.Yes)
                {
                    int userCode = (int)rowUsers.Cells[0].Value;
                    if (addingUser(rowUsers.Cells[1].Value + " " + rowUsers.Cells[2].Value, userCode,true, rowUsers.Cells[3].Value.ToString(),
                        PublicFuncsNvars.getUserEmail(userCode)))
                    {
                        MessageBox.Show("המשתמש \"" + rowUsers.Cells[3].Value.ToString() + "\" נוסף בהצלחה", "הוספה לרשימת תפוצה",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
            }
            else
            {
                MessageBox.Show("על מנת להוסיף משתמש עליכם לסמן את המשתמש בטבלת המשתמשים, ואת הרשימה הרצויה", "הוספה לרשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void RecipientListsUpdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData.HasFlag(Keys.K) && ModifierKeys.HasFlag(Keys.Control))
            {
                if (okAddUser)
                {
                    if (dataGridViewRecipientLists.SelectedRows.Count > 0)
                    {
                        this.BringToFront();
                        List<Tuple<string, string, string, bool>> rec = PublicFuncsNvars.getCtrlKRecipients();
                        DataGridViewRow rowLists = dataGridViewRecipientLists.SelectedRows[0];
                        foreach (Tuple<string, string, string, bool> r in rec)
                        {
                            DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם להוסיף את המשתמש " + r.Item1 + " לרשימת התפוצה \"" + rl.name + "\"?", "הוספת משתמש לרשימת תפוצה",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                      MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                            if (res == DialogResult.Yes)
                            {
                                addingUser(r.Item1, 99999, r.Item4, r.Item3, r.Item2);
                            }
                            MessageBox.Show("המשתמשים נוספו בהצלחה", "הוספה לרשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        }
                    }
                    else
                    {
                        MessageBox.Show("על מנת להוסיף משתמש עליכם לסמן את הרשימה הרצויה", "הוספה לרשימת תפוצה",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
            }
        }

        private bool addingUser(string name, int userCode, bool ifa, string role, string email)
        {
            foreach (Recipient rec in rl.getRecipients())
                if ((userCode != 99999 && userCode == rec.getId()) || (userCode == 99999 && role == rec.getRole()))
                {
                    MessageBox.Show("משתמש זה כבר נמצא ברשימת תפוצה זו", "הוספה לרשימת תפוצה",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    return false;
                }

            DataGridViewRow rowLists = dataGridViewRecipientLists.SelectedRows[0];
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT isnull(MAX(onum),0) FROM dbo.tm_tfuz_res WHERE cod_sys=@sysCode AND cod_lst_tpotzh=@listCode", conn);
            comm.Parameters.AddWithValue("@listCode", (short)rowLists.Cells["RListIDColumn"].Value);
            comm.Parameters.AddWithValue("@sysCode", (short)rowLists.Cells["sysCodeColumn"].Value);
            conn.Open();
            short nid = (short)((short)comm.ExecuteScalar()+(short)1);
            conn.Close();
            comm.CommandText="INSERT INTO dbo.tm_tfuz_res (cod_sys, cod_lst_tpotzh, onum , cod_mcotb, is_actn_ydyah, tiur_tafkid, ktovet_mail)"
                                            + Environment.NewLine + "output inserted.onum" + Environment.NewLine
                                            + "VALUES(@sysCode, @listCode, @nid, @userCode, @ifa, @role, @email)";
            comm.Parameters.AddWithValue("@nid", nid);
            comm.Parameters.AddWithValue("@userCode", userCode);
            comm.Parameters.AddWithValue("@ifa", !ifa);
            comm.Parameters.AddWithValue("@role", role);
            comm.Parameters.AddWithValue("@email", email);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            rl.addRecipient(new Recipient(userCode, nid, role, ifa, true, email));
            int rowNum = dataGridViewListContents.Rows.Add(userCode, name, role);
            dataGridViewListContents.CellValueChanged -= dataGridViewListContents_CellValueChanged;
            dataGridViewListContents.Rows[rowNum].Cells[3].Value = ifa ? "לפעולה" : "לידיעה";
            dataGridViewListContents.CellValueChanged += dataGridViewListContents_CellValueChanged;
            dataGridViewListContents.Rows[rowNum].Selected = true;
            return true;

        }

        private void dataGridViewUsers_KeyPress(object sender, KeyPressEventArgs e)
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

        private void dataGridViewUsers_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridViewUsers_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

        private void RecipientListsUpdate_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!isJustOne)
                Program.rlu = null;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("אין לעדכן רשימה ללא שם", "עדכון רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else if (comboBox1.Text == "" || comboBox1.Text == "הכל")
            {
                MessageBox.Show("אין לעדכן רשימה ללא ענף", "עדכון רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else if (comboBox3.Text == "" || comboBox3.Text == "הכל")
            {
                MessageBox.Show("אין לעדכן רשימה ללא רמת רשימה", "עדכון רשימת תפוצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            else
            {
                int res2;
                if (int.TryParse(textBox2.Text, out res2))
                {
                    if (PublicFuncsNvars.userCodeExists(res2))
                    {
                        DataGridViewRow row = dataGridViewRecipientLists.SelectedRows[0];
                        KeyValuePair<short, short> kvp = new KeyValuePair<short, short>((short)row.Cells["sysCodeColumn"].Value, (short)row.Cells["RListIDColumn"].Value);
                        DialogResult res = MessageBox.Show("האם אתם בטוחים שברצונכם לעדכן את רשימת התפוצה \"" + textBox1.Text + "\"?", "עדכון רשימת תפוצה",
                                                          MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                          MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        if (res == DialogResult.Yes)
                        {
                            SqlConnection conn = new SqlConnection(Global.ConStr);
                            SqlCommand comm = new SqlCommand("UPDATE dbo.tm_tfuza SET nam_lst_tpotzh=@listName, anf_tpotzh=@branch, creating_user=@creatingUser, "
                                                            + "personalObranchOunit=@personalObranchOunit WHERE cod_sys=@sysCode AND cod_lst_tpotzh=@listId", conn);
                            comm.Parameters.AddWithValue("@sysCode", kvp.Key);
                            comm.Parameters.AddWithValue("@listId", kvp.Value);
                            comm.Parameters.AddWithValue("@listName", textBox1.Text);
                            char c = PublicFuncsNvars.getBranchByString(comboBox1.Text);
                            comm.Parameters.AddWithValue("@branch", c);
                            comm.Parameters.AddWithValue("@creatingUser", res2);
                            short level = (short)PublicFuncsNvars.getRecipientListLevelByString(comboBox3.Text);
                            comm.Parameters.AddWithValue("@personalObranchOunit", level);
                            conn.Open();
                            comm.ExecuteNonQuery();
                            conn.Close();
                            row.Cells[2].Value=PublicFuncsNvars.recipientsLists[kvp].name = textBox1.Text;
                            PublicFuncsNvars.recipientsLists[kvp].branch = (Branch)c;
                            PublicFuncsNvars.recipientsLists[kvp].owner = res2;
                            PublicFuncsNvars.recipientsLists[kvp].level = (RecipientListsLevel)level;
                            row.Cells[3].Value = PublicFuncsNvars.recipientsLists[kvp].getLevelString();
                            row.Cells[4].Value = PublicFuncsNvars.recipientsLists[kvp].getOwnerName();
                            MessageBox.Show("רשימת התפוצה \"" + textBox1.Text + "\" עודכנה בהצלחה", "עדכון רשימת תפוצה",
                                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                            MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        }
                    }
                    else
                    {
                        MessageBox.Show("אין משתמש כזה במערכת", "עדכון רשימת תפוצה",
                                         MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                         MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }
                else
                    MessageBox.Show("מספר משתמש לא יכול להכיל תווים שאינם ספרות, ולא יכול להיות ריק", "עדכון רשימת תפוצה",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }

        private void dataGridViewListContents_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                if (dataGridViewRecipientLists.SelectedRows.Count > 0 && dataGridViewListContents.SelectedCells.Count > 0)
                {
                    DataGridViewRow rowLists = dataGridViewRecipientLists.SelectedRows[0];
                    DataGridViewRow rowContents = dataGridViewListContents.SelectedCells[0].OwningRow;
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("UPDATE dbo.tm_tfuz_res SET is_actn_ydyah=@ifa WHERE cod_lst_tpotzh=@listCode AND cod_sys=@sysCode AND onum=@recipientNum", conn);
                    comm.Parameters.AddWithValue("@listCode", (short)rowLists.Cells["RListIDColumn"].Value);
                    comm.Parameters.AddWithValue("@sysCode", (short)rowLists.Cells["sysCodeColumn"].Value);
                    Recipient r = rl.getRecipients().Where(x => x.getId() == (int)rowContents.Cells[0].Value).ToList()[0];
                    short nid = r.getNID();
                    comm.Parameters.AddWithValue("@recipientNum", nid);
                    bool ifa = rowContents.Cells[e.ColumnIndex].Value.ToString() == "לידיעה";
                    comm.Parameters.AddWithValue("@ifa", ifa);
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    r.setIFA(!ifa);
                    MessageBox.Show("המשתמש \"" + rowContents.Cells[2].Value + "\" עבר ל" + rowContents.Cells[e.ColumnIndex].Value.ToString(), "שינוי סוג כיתוב",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
                else
                {
                    MessageBox.Show("על מנת לעדכן משתמש עליכם לסמן את המשתמש בטבלת תוכן הרשימה, ואת הרשימה הרצויה", "עדכון מכותב",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
        }

        private void dataGridViewListContents_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridViewListContents.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            if (textBox2.Enabled)
            {
                dataGridViewUsers.SelectionChanged -= dataGridViewUsers_SelectionChanged;
                panel1.Visible = true;
                dataGridViewUsers.SelectionChanged += dataGridViewUsers_SelectionChanged;

            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Enabled && !dataGridViewUsers.Focused)
                panel1.Visible = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            panel6.Visible = true;
            panel4.Visible = false;
            panel1.Visible = false;
            comboBox2.Visible = false;
            button10.Visible = true;
            button5.Visible = false;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            panel6.Visible = false;
            panel4.Visible = true;
            panel1.Visible = false;
            comboBox2.Visible = true;
            button10.Visible = true;
            button5.Visible = false;
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
                addingUser(interIds, 99999, comboBox2.Text == "לפעולה", interIds, "");
                panel4.Visible = false;
                button10.Visible = false;
                comboBox2.Visible = false;
            }
            else
            {
                MessageBox.Show("נא לבחור לפחות ת.פ. אחד", "בחירת ת.פ. שגויה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            addingUser(textBox15.Text, 99999, comboBox4.Text == "לפעולה", textBox15.Text, "");
            panel6.Visible = false;
            button10.Visible = false;
            textBox15.Clear();
        }
    }
}
