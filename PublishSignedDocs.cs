using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentsModule
{
    public partial class PublishSignedDocs : Form
    {
        public PublishSignedDocs()
        {
            InitializeComponent();
        }

        private void PublishSignedDocs_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            User u = PublicFuncsNvars.curUser;
            bool isPerson = false;
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.documents' table. You can move, or remove it, as needed.
            if ((u.roleType == RoleType.clerk && u.permissionsBranch == Branch.office) || u.roleType == RoleType.computers)
                this.documentsSignTableAdapter.FillToPublish(this.mantakDBDataSetDocuments.documentsSign, false, true, true);
            else if (u.roleType == RoleType.clerk)
                this.documentsSignTableAdapter.FillToPublishBranch(this.mantakDBDataSetDocuments.documentsSign, false, true, short.Parse(((char)u.permissionsBranch).ToString()), true);
            else
            {
                this.documentsSignTableAdapter.FillToPublishPerson(this.mantakDBDataSetDocuments.documentsSign, false, true, u.userCode, true);
                isPerson = true;
            }


            if (isPerson)
            {
                foreach (MantakDBDataSetDocuments.documentsSignRow row in mantakDBDataSetDocuments.documentsSign.OrderByDescending(x => x.tarich_hamichtav))
                {
                    int filingUser;
                    if (!int.TryParse(row.user_metaiek, out filingUser) || !PublicFuncsNvars.users.Any(x=>x.userCode==filingUser))
                        filingUser = PublicFuncsNvars.curUser.userCode;
                    try
                    {
                        if (row.tarich_hamichtav != "00000000")
                            dataGridViewDocs.Rows.Add(row.shotef_mismach, row.hanadon.Trim(), PublicFuncsNvars.users.First(x => x.userCode == filingUser).job,
                                DateTime.ParseExact(row.tarich_hamichtav, "yyyyMMdd", CultureInfo.CurrentCulture), "בטל חתימה", "הפצה");
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            else
            {
                senderColumn.HeaderText = "יוצר מסמך";
                foreach (MantakDBDataSetDocuments.documentsSignRow row in mantakDBDataSetDocuments.documentsSign.OrderByDescending(x => x.tarich_hamichtav))
                {
                    if (row.tarich_hamichtav != "00000000")
                        dataGridViewDocs.Rows.Add(row.shotef_mismach, row.hanadon.Trim(), row.teur_tafkid_sholeah.Trim(),
                            DateTime.ParseExact(row.tarich_hamichtav, "yyyyMMdd", CultureInfo.CurrentCulture), "בטל חתימה", "הפצה");
                }
            }
        }

        private void PublishSignedDocs_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Program.psd = null;
            }
            catch { }
        }

        private void dataGridViewDocs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                switch (e.ColumnIndex)
                {
                    case 4:
                        abortSignDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value);
                        dataGridViewDocs.Rows.RemoveAt(e.RowIndex);
                        break;
                    case 5:
                        PublicFuncsNvars.beginToPublishDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value,
                                    dataGridViewDocs.Rows[e.RowIndex].Cells["nameColumn"].Value.ToString());
                        dataGridViewDocs.Rows.RemoveAt(e.RowIndex);
                        break;
                    default:
                        break;
                }
            }
        }

        private void abortSignDoc(int docId)
        {
            if (PublicFuncsNvars.alreadySigned(docId))
            {
                DialogResult result = MessageBox.Show("האם אתם בטוחים שברצונכם לבטל את החתימה?", "חתימה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (DialogResult.Yes == result)
                {
                    PublicFuncsNvars.abortSignDoc(docId);
                    MessageBox.Show("החתימה בוטלה.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
            }
            else
                MessageBox.Show("לא ניתן לבטל חתימה על מסמך לא חתום.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void dataGridViewDocs_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                switch (e.ColumnIndex)
                {
                    case 4:
                        break;
                    default:
                        int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                        PublicFuncsNvars.openDocumentHandlingForm(id);
                        break;
                }
        }
    }
}
