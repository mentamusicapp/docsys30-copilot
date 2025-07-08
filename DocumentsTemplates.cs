using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;


namespace DocumentsModule
{
    public partial class DocumentsTemplates : Form
    {
        public DocumentsTemplates()
        {
            InitializeComponent();
            dataGridView2.CellEndEdit += dataGridView2_CellEndEdit;
            DocumentsMenu.PathTemplate(this.btnUpdateSql, 20);
        }

        private void DocumentsTemplates_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;

            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT onum_word_template ONum ,dscr_template Teur, nam_word_template Nativ ,lctn_kobyh_template Mikum ,file_extension Ext FROM tm_templ_bhi", conn);
            conn.Open();
            using (SqlDataReader reader = comm.ExecuteReader())
            {
                DataTable dt = new DataTable();
                dt.Load(reader);
                dt.Columns.Add("Modified", typeof(bool));
                dataGridView2.DataSource = dt;
                dataGridView2.Refresh();
                dataGridView2.Visible = true;
            }
            conn.Close();
            dataGridView2.Columns[0].ReadOnly = true;
            dataGridView2.Refresh();
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.Rows[e.RowIndex].Cells["modified"].Value = true;
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {}

        private void btnUpdateSql_Click(object sender, EventArgs e) 
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE tm_templ_bhi "+
            "SET dscr_template = @Teur, nam_word_template = @Nativ, lctn_kobyh_template = @Mikum, file_extension = @Ext "+
            "where onum_word_template = @ONum", conn);
            foreach(DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Cells["modified"].Value!=null)
                {
                    try
                    {
                        if ((bool)row.Cells["modified"].Value == true)
                        {
                            comm.Parameters.AddWithValue("@ONum", row.Cells[0].Value);
                            comm.Parameters.AddWithValue("@Teur", row.Cells[1].Value);
                            comm.Parameters.AddWithValue("@Nativ", row.Cells[2].Value);
                            comm.Parameters.AddWithValue("@Mikum", row.Cells[3].Value);
                            comm.Parameters.AddWithValue("@Ext", row.Cells[4].Value);
                            try
                            {
                                conn.Open();
                                comm.ExecuteNonQuery();
                                conn.Close();
                                row.Cells["modified"].Value = false;
                                MessageBox.Show("עודכן!");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("עדכון נכשל: " + ex.Message);
                            }
                        }
                    }
                    catch{}
                }
            }
        }
    }
}
