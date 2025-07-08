using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace DocumentsModule
{
    public partial class ExecSqlScript : Form
    {
        public ExecSqlScript()
        {
            InitializeComponent();
            this.Icon = Global.AppIcon;
        }

        private void ExecSqlScript_Load(object sender, EventArgs e)
        {
            DocumentsMenu.PathTemplate(this.BtnLoadSql, 30);
            DocumentsMenu.PathTemplate(this.BtnRunSql, 30);

            comboBox1.Items.AddRange(new object[]
            {
                Encoding.Default.WebName,
                Encoding.UTF8.WebName,
                Encoding.ASCII.WebName,
                Encoding.Unicode.WebName,
                Encoding.UTF32.WebName
            });
            comboBox1.SelectedIndex = 0;
        }

        private void BtnLoadSql_Click(object sender, EventArgs e)
        {
            if( textBox1.Text != "")
            {
                if (File.Exists(textBox1.Text))
                {
                    string encodingName = comboBox1.SelectedItem.ToString();
                    string script = File.ReadAllText(textBox1.Text, Encoding.GetEncoding(encodingName));
                    textBox2.Text = script;
                    BtnRunSql.Visible = true;
                }
                else
                {
                    MessageBox.Show("הקובץ לא קיים.");
                }
            }
        }

        private void BtnRunSql_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            try
            {
                conn.Open();
                IEnumerable<string> commandStrings = Regex.Split(textBox2.Text, @"^\s*GO\s*$", RegexOptions.Multiline| RegexOptions.IgnoreCase);
                foreach(string commandString in commandStrings)
                {
                    if (commandString.Trim() != "")
                        new SqlCommand(commandString, conn).ExecuteNonQuery();
                }
                MessageBox.Show("Data Base updated succesfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                conn.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog f = new OpenFileDialog())
            {
                f.Filter = "SQL files(*.sql)|*.sql|All files (*.*)|*.*";
                f.Title = "תבחר SQL קובץ.";
                if (f.ShowDialog() == DialogResult.OK)
                    textBox1.Text = f.FileName;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnLoadSql_Click(sender, e);
        }
    }
}
