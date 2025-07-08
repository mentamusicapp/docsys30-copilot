using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace DocumentsModule
{
    public partial class UsersUpdate : Form
    {
        private Dictionary<int, string> patterns;
        string strTyped = "";
        public UsersUpdate()
        {
            InitializeComponent();
        }

        private void UsersUpdate_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.roleTypes' table. You can move, or remove it, as needed.
            this.roleTypesTableAdapter.Fill(this.mantakDBDataSetDocuments.roleTypes);
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.tm_kubiot' table. You can move, or remove it, as needed.
            this.tm_kubiotTableAdapter.Fill(this.mantakDBDataSetDocuments.tm_kubiot);
            patterns = new Dictionary<int, string>();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT onum_word_template, dscr_template FROM dbo.tm_templ_bhi", conn); // WHERE file_data<>0x00000000
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                patterns.Add(sdr.GetInt16(0), sdr.GetString(1).Trim());
                dataGridView5.Rows.Add(sdr.GetInt16(0), sdr.GetString(1).Trim());
            }
            conn.Close();

            foreach (User u in PublicFuncsNvars.users)
                dataGridViewUsersForAuthorizations.Rows.Add(u.userCode, u.firstName, u.lastName, u.job);
            DocumentsMenu.PathTemplate(this.button1, 55);
            DocumentsMenu.PathTemplate(this.button16, 55);
            DocumentsMenu.PathTemplate(this.button3, 30);
            DocumentsMenu.PathTemplate(this.button2, 30);
            DocumentsMenu.PathTemplate(this.button11, 30);
            DocumentsMenu.PathTemplate(this.button4, 30);
            DocumentsMenu.PathTemplate(this.button5, 30);
            DocumentsMenu.PathTemplate(this.button6, 30);
            DocumentsMenu.PathTemplate(this.button7, 30);
            DocumentsMenu.PathTemplate(this.button8, 30);
            DocumentsMenu.PathTemplate(this.button9, 30);
            DocumentsMenu.PathTemplate(this.button12, 30);
            DocumentsMenu.PathTemplate(this.button32, 30);
            DocumentsMenu.PathTemplate(this.button10, 25);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PublicFuncsNvars.getUsers();
            usersDataGridView.Rows.Clear();
            int res2 = 0;
            bool tb2Empty = textBox2.Text.Equals(""), cbEmpty = comboBox3.Text.Equals("") || comboBox3.Text.Equals("הכל");
            int.TryParse(textBox2.Text, out res2);
            Branch b = (Branch)PublicFuncsNvars.getBranchByString(comboBox3.Text);

            var usersToDisplay =
                from user in PublicFuncsNvars.users
                where (user.userCode == res2 || tb2Empty) && user.userLogin.Contains(textBox1.Text) && user.firstName.Contains(textBox3.Text) &&
                       user.lastName.Contains(textBox27.Text) && user.job.Contains(textBox9.Text) && (cbEmpty || user.branch == b)
                orderby user.userCode ascending
                select user;

            foreach (User u in usersToDisplay)
                usersDataGridView.Rows.Add(u.userLogin, u.userCode, u.firstName, u.lastName, u.job, PublicFuncsNvars.getBranchString(u.branch));
            if (usersDataGridView.Rows.Count == 0)
            {
                MessageBox.Show("לא נמצאו משתמשים",
                                "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox9.Clear();
            textBox27.Clear();
            comboBox3.SelectedItem = null;
            usersDataGridView.Rows.Clear();
        }

        private void UsersUpdate_FormClosed(object sender, FormClosedEventArgs e)
        {
            Services.uu = null;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (PublicFuncsNvars.curUser.roleType == RoleType.computers)
            {
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox8.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox13.Clear();
                textBox14.Clear();
                textBox15.Clear();
                textBox16.Clear();
                textBox17.Clear();
                textBox18.Clear();
                textBox19.Clear();
                textBox20.Clear();
                textBox21.Clear();
                textBox22.Clear();
                textBox24.Clear();
                textBox25.Clear();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                comboBox4.SelectedItem = null;
                panel1.Visible = false;
                panel2.Visible = true;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                button5.BringToFront();
                textBox7.ReadOnly = false;
            }
            else
                MessageBox.Show(".רק צוות מחשוב יכול לפתוח משתמשים", "משתמש חדש", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int res7 = -1, res13=-1, res25=-1,res8;
            short res12 = -1;

            if ((!textBox8.Text.ToUpper().StartsWith("W") && !textBox8.Text.ToUpper().StartsWith("U")) || !int.TryParse(textBox8.Text.Substring(1), out res8) ||
                    textBox8.Text.Length != 6 || PublicFuncsNvars.userLoginExists(textBox8.Text) )
                    {
                        MessageBox.Show(".משתמש משהב\"ט זה כבר קיים עבור משתמש אחר או שאינו חוקי, אנא בחרו משתמש משהב\"ט אחר" + Environment.NewLine
                                + "משתמש משהב\"ט חוקי מתחיל באות U ואחריה 5 ספרות.",
                                "משתמש משהב\"ט קיים או לא חוקי", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                     }
            else if (!int.TryParse(textBox7.Text, out res7) || PublicFuncsNvars.userCodeExists(res7))
            {
                MessageBox.Show("יוזר ביחידה זה כבר קיים ביחידה עבור משתמש אחר או שאינו ערך חוקי, אנא בחרו יוזר אחר"+ Environment.NewLine
                                + "יוזר ביחידה מכיל רק ספרות.",
                                "יוזר קיים ביחידה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (textBox5.Text == "" || textBox4.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox4.Text == "" || textBox6.Text == "" ||
                     textBox7.Text == "" || textBox8.Text == "" || textBox10.Text == "" || textBox12.Text == "" || textBox13.Text == "" || textBox14.Text == "" ||
                     textBox15.Text == "" || textBox25.Text == "" || textBox26.Text == "")
            {
                MessageBox.Show("אין להשאיר את שדות ריקים",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (!short.TryParse(textBox12.Text, out res12) || !cubesContains(res12) || !int.TryParse(textBox13.Text, out res13) || textBox13.Text.Length > 9 ||
                     !int.TryParse(textBox25.Text, out res25) || !PublicFuncsNvars.userCodeExists(res25))
            {
                MessageBox.Show("בשדות 'קוביה', 'ת.ז.' ו-'יוזר מפקד' יש להכניס אך ורק ערכים מספריים חוקיים",
                                "ערכים לא חוקיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if(!PublicFuncsNvars.validEmail(textBox15.Text))
            {
                MessageBox.Show("כתובת דוא\"ל לא חוקית",
                                "ערכים לא חוקיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("INSERT INTO dbo.tmtafkidu(kod_tpkid, taor_tpkid, sog_gop_ihidh, anp_hbrh_ihidh, bal_tpkid_bmshimot_, drgh, "
                                                +"shm_mshphh, shm_prti, shm_mshtmsh_MAGIC, kod_template_WORD, kod_sioog_bthoni, sog_tarich_lhdpsh, kod_kobih, "
                                                +"User_Login, sog_mshtmsh, kod_drgh, kod_tpki, kod_ihidh, kod_mdor, mdpst_1__Capture, tlpon_abodh, pks, tlpon_bit, "
                                                +"pks_bit, chtobt, air, mikod, hatimh, anp_mshtmsh_bmntk, atz_mshimot_bm, is_lhpail_bntib_ahron, chrtst_batz_bm, "
                                                +"is_lakob_ahr_chrtst, modol_hpalh_bm, mhot_hnhih_bm, ntib_bhirt_kbtzim_bm, is_lakob_ahr_ntib_kbtzim, "
                                                +"archh_shosh_mshimh_bm, iozm_hnhih_bm, sholh_baitor_mchtb_bm, bm_lhtch_tikim, is_lshloh_doh_hnhiot_sttos_iomi_, "
                                                +"tz, is_midor_nsph__aishi, is_midor_nsph_anpi, doal, getsMailOnAnyUserBirthday, commanderId, roleTypeCode, isActive, signature) "
                                                +"VALUES (@userId, @jobDescription, 1, @branch, 0, 0, @lastName, @firstName, @userId, @pattern, @classification, "
                                                +"2, @cube, @login, '', 14, 0, 0, 0, 0, @workPhone, '', '', '', '', '', 0, @signature, @branch, '', 0, '', 0, '', "
                                                +"'', '', 0, '', '', 0, '', 0, @idNum, 0, 0, @email, 0, @commanderId, @roleTypeCode, 1, @signatureImage)", conn);
                comm.Parameters.AddWithValue("@userid", res7);
                comm.Parameters.AddWithValue("@jobDescription", textBox4.Text);
                comm.Parameters.AddWithValue("@lastName", textBox26.Text);
                comm.Parameters.AddWithValue("@firstName", textBox5.Text);
                comm.Parameters.AddWithValue("@pattern", textBox10.Text);
                comm.Parameters.AddWithValue("@classification", PublicFuncsNvars.getClassificationCode(comboBox2.Text));
                comm.Parameters.AddWithValue("@cube", res12);
                comm.Parameters.AddWithValue("@login", textBox8.Text.ToUpper());
                comm.Parameters.AddWithValue("@workPhone", textBox14.Text);
                comm.Parameters.AddWithValue("@signature", textBox6.Text);
                char branchChar = PublicFuncsNvars.getBranchByString(comboBox1.Text);
                short branchShort;
                if (!short.TryParse(branchChar.ToString(), out branchShort))
                    branchShort = 1;
                comm.Parameters.AddWithValue("@branch", branchShort);
                comm.Parameters.AddWithValue("@idNum", res13);
                comm.Parameters.AddWithValue("@email", textBox15.Text);
                comm.Parameters.AddWithValue("@commanderId", res25);
                short roleTypeShort=(short)comboBox4.SelectedValue;
                comm.Parameters.AddWithValue("@roleTypeCode", roleTypeShort);
                byte[] sigArr = (byte[])(new ImageConverter()).ConvertTo(pictureBox1.Image, typeof(byte[]));
                comm.Parameters.AddWithValue("@signatureImage", sigArr);

                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();

                PublicFuncsNvars.users.Add(new User(res7, res7, textBox5.Text, textBox26.Text, textBox15.Text, textBox8.Text, textBox4.Text,
                                                    branchChar, branchShort, res25, roleTypeShort, true, checkBox12.Checked));


                panel2.Visible = false;
                panel1.Visible = true;

                MessageBox.Show("משתמש " + textBox7.Text + ": " + textBox26.Text + " " + textBox5.Text + " נוצר בהצלחה",
                                "אישור יצירת משתמש", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                usersDataGridView.Rows.Clear();
            }
        }

        private bool cubesContains(short res12)
        {
            MantakDBDataSetDocuments.tm_kubiotDataTable dt = tm_kubiotTableAdapter.GetData();
            foreach(DataRow row in dt.Rows)
            {
                if (row.Field<short>(0)==res12)
                    return true;
            }
            return false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel1.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            checkBox10.Checked = false;
            checkBox10.Checked = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            List<string> headers = new List<string>();

            if (checkBox1.Checked)
                headers.Add(checkBox1.Text);
            if (checkBox2.Checked)
                headers.Add(checkBox2.Text);
            if (checkBox3.Checked)
                headers.Add(checkBox3.Text);
            if (checkBox4.Checked)
                headers.Add(checkBox4.Text);
            if (checkBox5.Checked)
                headers.Add(checkBox5.Text);
            if (checkBox6.Checked)
                headers.Add(checkBox6.Text);
            if (checkBox7.Checked)
                headers.Add(checkBox7.Text);
            if (checkBox8.Checked)
                headers.Add(checkBox8.Text);
            if (checkBox9.Checked)
                headers.Add(checkBox9.Text);

            
            List<string[]> values = new List<string[]>();


            IEnumerable<User> users = PublicFuncsNvars.users;

            users =
                from user in PublicFuncsNvars.users
                where user.userLogin.StartsWith("U") && user.isActive
                orderby user.lastName ascending
                select user;

            foreach (User u in users)
            {
                string[] valRow = new string[headers.Count];
                int i = 0;
                foreach (Control c in panel5.Controls)
                {
                    if (c is CheckBox && ((CheckBox)c).Checked&& c!=checkBox10)
                    {
                        string prop = getUserPropertyName(((CheckBox)c).Text);
                        var type1 = u.GetType();

                        PropertyInfo pi = type1.GetProperty(prop, BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);


                        dynamic value = pi.GetValue(u, null);

                        if (value != null)
                        {
                            if (value is Branch)
                                valRow[i] = PublicFuncsNvars.getBranchString(value);
                            else if (value is RoleType)
                            {
                                MantakDBDataSetDocuments.roleTypesDataTable dt = roleTypesTableAdapter.GetData();
                                DataRow row = dt.Rows.Find((short)value);
                                valRow[i] = row.Field<string>(1);
                            }
                            else if (prop == "commanderCode")
                                valRow[i] = PublicFuncsNvars.users.Where(x => x.userCode == u.commanderCode).ToList()[0].getFullName();
                            else
                                valRow[i] = value.ToString();
                        }
                        i++;
                        if (i == valRow.Length)
                            break;
                    }
                }
                values.Add(valRow);
            }

            PublicFuncsNvars.exportToXL("users-list", "רשימת משתמשים", headers.ToArray(), values);
            panel5Invisible();
        }

        private string getUserPropertyName(string s)
        {
            switch(s)
            {
                case "משתמש משהב\"ט":
                    return "userLogin";
                case "יוזר ביחידה":
                    return "userCode";
                case "שם פרטי":
                    return "firstName";
                case "שם משפחה":
                    return "lastName";
                case "תפקיד":
                    return "job";
                case "ענף":
                    return "branch";
                case "אי-מייל":
                    return "email";
                case "מפקד":
                    return "commanderCode";
                case "סוג משתמש":
                    return "roleType";
                default:
                    return null;
            }
        }

        private void usersDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridViewAuthorizations.Rows.Clear();
                User u = PublicFuncsNvars.users.Where(x => x.userCode == (int)usersDataGridView.Rows[e.RowIndex].Cells[1].Value).ToList()[0];
                textBox4.Text = usersDataGridView.Rows[e.RowIndex].Cells[4].Value.ToString();
                textBox26.Text = usersDataGridView.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox5.Text = usersDataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox7.Text = usersDataGridView.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox8.Text = usersDataGridView.Rows[e.RowIndex].Cells[0].Value.ToString();
                pictureBox1.Image = PublicFuncsNvars.getSignatureImage(u.userCode);

                foreach (KeyValuePair<int, bool> au in u.getAutoAuthorizedUsers())
                {
                    User tu = PublicFuncsNvars.getUserByCode(au.Key);
                    string name = tu.getFullName(), role = tu.job;
                    int rowIndex = dataGridViewAuthorizations.Rows.Add(au.Key, name, role);
                    dataGridViewAuthorizations.Rows[rowIndex].Cells[3].Value = au.Value ? "לעריכה" : "לצפיה";
                }

                SqlConnection conn=new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT kod_template_WORD, kod_kobih, tlpon_abodh, hatimh, tz, kod_sioog_bthoni FROM dbo.tmtafkidu WHERE kod_tpkid=@userId", conn);
                comm.Parameters.AddWithValue("@userId", u.userCode);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                sdr.Read();
                textBox6.Text=sdr.GetString(3).ToString();
                textBox10.Text=sdr.GetInt16(0).ToString();
                textBox11.Clear();
                textBox12.Text=sdr.GetInt16(1).ToString();
                textBox13.Text=sdr.GetInt32(4).ToString();
                textBox14.Text=sdr.GetString(2).ToString();
                textBox15.Text=u.email;
                textBox25.Text = u.commanderCode.ToString();
                comboBox1.SelectedIndex = comboBox1.Items.IndexOf(PublicFuncsNvars.getBranchString(u.branch));
                comboBox2.SelectedIndex = comboBox2.Items.IndexOf(PublicFuncsNvars.getClassificationByEnum(PublicFuncsNvars.getClassification(sdr.GetInt16(5))));
                comboBox4.SelectedValue = u.roleType;
                listBox1.SelectedValue = textBox12.Text;
                setFirstPatternRow();
                panel1.Visible = false;
                panel2.Visible = true;
                panel5.Visible = false;
                textBox7.ReadOnly = true;
                checkBox11.Checked = u.isActive;
                checkBox12.Checked = u.allowedToOpenFolders;
                button6.BringToFront();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int res7 = int.Parse(textBox7.Text), res13 = -1, res25 = -1, res8;
            short res12 = -1;
            DialogResult res = DialogResult.Cancel;

            //if ((!int.TryParse(textBox8.Text.Substring(1), out res8) ||
            //      textBox8.Text.Length != 6 || (PublicFuncsNvars.userLoginExists(textBox8.Text) && !PublicFuncsNvars.userloginMatchesUserCode(textBox8.Text, res7))))

            
            // Checks if the user code is empty or first char between A and Z and min length = 6 and user not exists
            if (textBox8.Text.Length <= 5  || (textBox8.Text != "" && (textBox8.Text[0].ToString().ToUpper().ToCharArray()[0] < 'A' || textBox8.Text[0].ToString().ToUpper().ToCharArray()[0] > 'Z')))// || textBox8.Text != "" && PublicFuncsNvars.userLoginExists(textBox8.Text))
            {
                MessageBox.Show("משתמש זה כבר קיים עבור משתמש אחר או שאינו חוקי, אנא בחרו משתמש אחר",
                                "משתמש קיים או לא חוקי", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (textBox5.Text == "" || textBox4.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox4.Text == "" || textBox6.Text == "" ||
                     textBox7.Text == "" || textBox10.Text == "" || textBox12.Text == "" || textBox13.Text == "" || textBox14.Text == "" ||
                     textBox15.Text == "" || textBox25.Text == "" || textBox26.Text == "")
            {
                MessageBox.Show("אין להשאיר את שדות ריקים",
                                "שדות ריקים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (!short.TryParse(textBox12.Text, out res12) || !cubesContains(res12) || !int.TryParse(textBox13.Text, out res13) || textBox13.Text.Length > 9 ||
                     !int.TryParse(textBox25.Text, out res25) || !PublicFuncsNvars.userCodeExists(res25))
            {
                MessageBox.Show("בשדות 'קוביה', 'ת.ז.' ו-'יוזר מפקד' יש להכניס אך ורק ערכים מספריים חוקיים",
                                "ערכים לא חוקיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else if (!PublicFuncsNvars.validEmail(textBox15.Text))
            {
                MessageBox.Show("כתובת דוא\"ל לא חוקית",
                                "ערכים לא חוקיים", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
            else
            {
                if (textBox8.Text == "")
                {
                    res = MessageBox.Show("לא הוכנס משתמש משהב\"ט. היוזר יהפוך ללא פעיל. האם להמשיך?",
                                    "משתמש משהב\"ט ריק", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
                if (res != DialogResult.No)
                {
                    if (res == DialogResult.Yes)
                        checkBox11.Checked = false;

                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand comm = new SqlCommand("UPDATE dbo.tmtafkidu SET taor_tpkid=@jobDescription, shm_mshphh=@lastName, "
                        + "shm_prti=@firstName, kod_template_WORD=@pattern, kod_sioog_bthoni=@classification, kod_kobih=@cube, "
                        + "User_Login=@login, tlpon_abodh=@workPhone, hatimh=@signature, anp_mshtmsh_bmntk=@branch, tz=@idNum, "
                        + "doal=@email, commanderId=@commanderId, roleTypeCode=@roleTypeCode, isActive=@isActive, isAllowedToOpenFolders=@isAllowedToOpenFolders, "
                        + "signature=@signatureImage WHERE kod_tpkid=@id", conn);
                    comm.Parameters.AddWithValue("@id", res7);
                    comm.Parameters.AddWithValue("@jobDescription", textBox4.Text);
                    comm.Parameters.AddWithValue("@lastName", textBox26.Text);
                    comm.Parameters.AddWithValue("@firstName", textBox5.Text);
                    comm.Parameters.AddWithValue("@pattern", textBox10.Text);
                    comm.Parameters.AddWithValue("@classification", PublicFuncsNvars.getClassificationCode(comboBox2.Text));
                    comm.Parameters.AddWithValue("@cube", textBox12.Text);
                    comm.Parameters.AddWithValue("@login", textBox8.Text);
                    comm.Parameters.AddWithValue("@workPhone", textBox14.Text);
                    comm.Parameters.AddWithValue("@signature", textBox6.Text);
                    comm.Parameters.AddWithValue("@branch", int.Parse(PublicFuncsNvars.getBranchByString(comboBox1.Text).ToString()));
                    comm.Parameters.AddWithValue("@idNum", int.Parse(textBox13.Text));
                    comm.Parameters.AddWithValue("@email", textBox15.Text);
                    comm.Parameters.AddWithValue("@commanderId", int.Parse(textBox25.Text));
                    comm.Parameters.AddWithValue("@roleTypeCode", (short)comboBox4.SelectedValue);
                    comm.Parameters.AddWithValue("@isActive", checkBox11.Checked);
                    comm.Parameters.AddWithValue("@isAllowedToOpenFolders", checkBox12.Checked);
                    byte[] sigArr = (byte[])(new ImageConverter()).ConvertTo(pictureBox1.Image, typeof(byte[]));
                    comm.Parameters.AddWithValue("@signatureImage", sigArr);

                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();


                    var u = PublicFuncsNvars.users.Find(x => x.userCode == int.Parse(textBox7.Text));
                    if (checkBox11.Checked)
                    {
                        u.job = textBox4.Text;
                        u.lastName = textBox26.Text;
                        u.firstName = textBox5.Text;
                        u.userLogin = textBox8.Text;
                        u.branch = (Branch)PublicFuncsNvars.getBranchByString(comboBox1.Text);
                        u.email = textBox15.Text;
                        u.commanderCode = int.Parse(textBox25.Text);
                        u.roleType = (RoleType)int.Parse(comboBox4.SelectedValue.ToString());
                    }
                    else
                        u.isActive = false;

                    panel2.Visible = false;
                    panel1.Visible = true;

                    MessageBox.Show("משתמש " + textBox7.Text + ": " + textBox26.Text + " " + textBox5.Text + " עודכן בהצלחה",
                                    "אישור עדכון משתמש", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                    usersDataGridView.Rows.Clear();
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
            {
                var cube = mantakDBDataSetDocuments.tm_kubiot.Rows[listBox1.SelectedIndex].ItemArray;
                textBox16.Text = cube[1].ToString();
                textBox17.Text = cube[2].ToString();
                textBox18.Text = cube[3].ToString();
                textBox19.Text = cube[4].ToString();
                textBox20.Text = cube[5].ToString();
                textBox21.Text = cube[6].ToString();
                textBox22.Text = cube[7].ToString();
                textBox23.Text = cube[8].ToString();
                textBox24.Text = cube[9].ToString();
                textBox12.Text = listBox1.SelectedValue.ToString();
            }
        }

        private void textBox12_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            panel4.Visible = true;
            panel3.Visible = false;
            panel7.Visible = false;
            listBox1.SelectedValue = textBox12.Text == "" ? "1" : textBox12.Text;
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
            panel3.Visible = true;
            panel4.Visible = false;
            panel7.Visible = false;
            setFirstPatternRow();
        }

        private void setFirstPatternRow()
        {
            if (!textBox10.Text.Equals(""))
            {
                int index = 0;
                dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == textBox10.Text)
                    {
                        index = row.Index;
                        break;
                    }
                }
                dataGridView5.FirstDisplayedScrollingRowIndex = index;
                dataGridView5.Rows[dataGridView5.FirstDisplayedScrollingRowIndex].Selected = true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel7.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel7.Visible = true;
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox10.Text = dataGridView5.SelectedRows[0].Cells[0].Value.ToString();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            int res;
            if (textBox10.Text!=""&& int.TryParse(textBox10.Text, out res)&&patterns.Keys.Contains(res))
                textBox11.Text = patterns[res];
            else
            {
                textBox10.Text = "";
                textBox11.Text = "";
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox10.Checked)
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
            }
        }

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (!((CheckBox)sender).Checked)
            {
                checkBox10.CheckedChanged -= checkBox10_CheckedChanged;
                checkBox10.Checked = false;
                checkBox10.CheckedChanged += checkBox10_CheckedChanged;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel5Invisible();
        }

        private void panel5Invisible()
        {
            checkBox10.Checked = true;
            checkBox10.Checked = false;
            panel5.Visible = false;
            panel1.Visible = true;
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            PublicFuncsNvars.textBox_Enter();
        }

        private void dataGridViewUsersForAuthorizations_KeyPress(object sender, KeyPressEventArgs e)
        {
            strTyped += e.KeyChar;
            int col = dataGridViewUsersForAuthorizations.SelectedCells[0].ColumnIndex;
            foreach (DataGridViewRow row in dataGridViewUsersForAuthorizations.Rows)
            {
                if (row.Cells[col].Value != null && row.Cells[col].Value.ToString().StartsWith(strTyped))
                {
                    dataGridViewUsersForAuthorizations.ClearSelection();
                    row.Cells[col].Selected = true;
                    dataGridViewUsersForAuthorizations.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }

        private void dataGridViewUsersForAuthorizations_KeyUp(object sender, KeyEventArgs e)
        {
            eraseStrTyped(e.KeyData);
        }

        private void dataGridViewUsersForAuthorizations_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strTyped = "";
        }

        private void eraseStrTyped(Keys keyData)
        {
            if (Keys.Right == keyData || Keys.Left == keyData || Keys.Up == keyData || Keys.Down == keyData || Keys.PageUp == keyData ||
                Keys.PageDown == keyData || Keys.Home == keyData || Keys.End == keyData || Keys.Tab == keyData)
                strTyped = "";
        }

        private void button32_Click(object sender, EventArgs e)
        {
            panel6.Visible = true;
            comboBox6.Visible = true;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridViewUsersForAuthorizations.SelectedCells.Count > 0)
            {
                List<DataGridViewRow> rowCollection = new List<DataGridViewRow>();
                foreach (DataGridViewCell cell in dataGridViewUsersForAuthorizations.SelectedCells)
                {
                    if (!rowCollection.Contains(cell.OwningRow))
                        rowCollection.Add(cell.OwningRow);
                }
                foreach (DataGridViewRow row in rowCollection)
                {
                    int userCode = int.Parse(row.Cells[0].Value.ToString());
                    var u = PublicFuncsNvars.users.Find(x => x.userCode == int.Parse(textBox7.Text));
                    if (u.addAuthorization(userCode, comboBox6.Text == "לעריכה"))
                    {
                        dataGridViewAuthorizations.Rows.Add(userCode, row.Cells[1].Value.ToString() + " " + row.Cells[2].Value.ToString(),
                                                        row.Cells[3].Value.ToString(), comboBox6.Text);
                    }
                }
                comboBox6.Visible = false;
                panel6.Visible = false;
            }
            strTyped = "";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            var u = PublicFuncsNvars.users.Find(x => x.userCode == int.Parse(textBox7.Text));
            u.removeAuthorization(int.Parse(dataGridViewAuthorizations.SelectedRows[0].Cells[0].Value.ToString()));
            dataGridViewAuthorizations.Rows.Remove(dataGridViewAuthorizations.SelectedRows[0]);
        }

        private void pictureBox1_Click(object sender, EventArgs e)//חתימה בתמונה
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Filter = "תמונות חתימה|*.png";
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            DialogResult res = ofd.ShowDialog();
            if(DialogResult.OK==res)
                pictureBox1.Image = Image.FromFile(ofd.FileName);
        }
    }
}
