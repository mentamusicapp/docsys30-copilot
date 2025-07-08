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
    public partial class TransferToSign : Form
    {
        public TransferToSign()
        {
            InitializeComponent();
        }

        private void TransferToSign_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.documents' table. You can move, or remove it, as needed.
            this.documentsSignTableAdapter.FillTransferToSign(this.mantakDBDataSetDocuments.documentsSign, false, PublicFuncsNvars.curUser.userCode.ToString(), true);
            foreach (MantakDBDataSetDocuments.documentsSignRow row in mantakDBDataSetDocuments.documentsSign.OrderByDescending(x => x.tarich_hamichtav))
            {
                dataGridViewDocs.Rows.Add(row.shotef_mismach, row.hanadon.Trim(), row.teur_tafkid_sholeah.Trim(),
                    DateTime.ParseExact(row.tarich_hamichtav.Trim(), "yyyyMMdd", CultureInfo.CurrentCulture), "העברה לחתימה");
            }
        }

        private void dataGridViewDocs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                switch (e.ColumnIndex)
                {
                    case 4:
                        transferDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value, e.RowIndex);
                        break;
                    default:
                        break;
                }
        }

        private void transferDoc(int docId, int index)
        {
            DialogResult result = MessageBox.Show("האם את/ה בטוח/ה שהמסמך מוכן לחתימה?", "העברה לחתימה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            if (DialogResult.Yes == result)
            {
                documentsSignTableAdapter.transferToSign(true, docId);
                dataGridViewDocs.Rows.RemoveAt(index);
                sendMailToSender(docId);
            }
        }
        private void sendMailToSender(int docId)
        {
            string senderEmail = PublicFuncsNvars.getUserEmail((int)(documentsSignTableAdapter.getSenderByDocId(docId)));
            PublicFuncsNvars.sendMail(PublicFuncsNvars.curUser.email + ";" + PublicFuncsNvars.curUser.job, senderEmail, null, null, "העברת שוטף לחתימה",
                "שוטף מספר " + docId + " הועבר אליך לחתימה ע\"י " + PublicFuncsNvars.curUser.getFullName()+"", new List<Tuple<byte[], string>>());
        }

        private void dataGridViewDocs_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                switch (e.ColumnIndex)
                {
                    case 4:
                    case 5:
                        break;
                    default:
                        int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                        PublicFuncsNvars.openDocumentHandlingForm(id);
                        break;
                }
        }

        private void TransferToSign_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Program.tts = null;
            }
            catch { }
            
        }
    }
}
