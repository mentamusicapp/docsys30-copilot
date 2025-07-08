using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentsModule
{
    public partial class ForMeToSign : Form
    {
        public ForMeToSign()
        {
            InitializeComponent();
        }

        private void ForMeToSign_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            // TODO: This line of code loads data into the 'mantakDBDataSetDocuments.documents' table. You can move, or remove it, as needed.
            this.documentsSignTableAdapter.FillForMeToSign(this.mantakDBDataSetDocuments.documentsSign, PublicFuncsNvars.curUser.userCode, true, false, false, true);
            foreach (MantakDBDataSetDocuments.documentsSignRow row in mantakDBDataSetDocuments.documentsSign.OrderByDescending(x => x.tarich_hamichtav))
            {
                int filingUser;
                if(!int.TryParse(row.user_metaiek, out filingUser))
                    filingUser=PublicFuncsNvars.curUser.userCode;
                dataGridViewDocs.Rows.Add(row.shotef_mismach, row.hanadon.Trim(), PublicFuncsNvars.users.Where(x => x.userCode == filingUser).First().job.Trim(),
                    DateTime.ParseExact(row.tarich_hamichtav, "yyyyMMdd", CultureInfo.CurrentCulture), "חתימה", "הפצה", "החזר להמשך עבודה");

            }
        }

        private void dataGridViewDocs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex>=0)
                switch (e.ColumnIndex)
                {
                    case 4:
                        if (dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == "חתימה")
                        {
                            if (signDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value))
                                dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "בטל חתימה";
                        }
                        else if (dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == "בטל חתימה")
                        {
                            if (abortSignDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value))
                                dataGridViewDocs.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "חתימה";
                        }
                        else
                            MessageBox.Show("כותרת הכפתור היא לא \"חתימה\" ולא \"בטל חתימה\". ישנה בעיה בכפתור זה, אנא פנו לצוות מחשוב.", "הפצה",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign |
                                MessageBoxOptions.RtlReading);
                        break;
                    case 5:
                        if (PublicFuncsNvars.alreadySigned((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value))
                        {
                            PublicFuncsNvars.beginToPublishDoc((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value,
                                dataGridViewDocs.Rows[e.RowIndex].Cells["nameColumn"].Value.ToString());
                            dataGridViewDocs.Rows.RemoveAt(e.RowIndex);
                        }
                        else
                            MessageBox.Show("לא ניתן להפיץ מסמך לא חתום.", "הפצה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        break;
                    case 6:
                        if (!PublicFuncsNvars.alreadySigned((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value))
                        {
                            MantakDBDataSetDocuments.documentsSignRow row=mantakDBDataSetDocuments.documentsSign.FindByshotef_mismach(
                                (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value);
                            int res;
                            if (int.TryParse(row.user_metaiek, out res) && (int)documentsSignTableAdapter.getSenderByDocId(row.shotef_mismach) != res)
                            {
                                documentsSignTableAdapter.transferToSign(false, (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value);
                                dataGridViewDocs.Rows.RemoveAt(e.RowIndex);
                                sendMailToCreator((int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value);
                            }
                            else
                                MessageBox.Show("לא ניתן להחזיר לעבודה מסמך שאת/ה יצרת.", "החזרה לעבודה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        }
                        else
                            MessageBox.Show("לא ניתן להחזיר לעבודה מסמך חתום.", "החזרה לעבודה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        break;
                    default:
                        break;
                }
        }

        private bool abortSignDoc(int docId)
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
                    return true;
                }
                else
                    return false;
            }
            else
            {
                MessageBox.Show("לא ניתן לבטל חתימה על מסמך לא חתום.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return false;
            }
        }

        private void sendMailToCreator(int docId)
        {
            string creatorEmail= PublicFuncsNvars.getUserEmail((int)(documentsSignTableAdapter.getCreatorByDocId(docId)));
            PublicFuncsNvars.sendMail(PublicFuncsNvars.curUser.email + ";" + PublicFuncsNvars.curUser.job, creatorEmail, null, null, "החזרת שוטף להמשך עבודה",
                "שוטף מספר " + docId + " הוחזר אליך להמשך עבודה ע\"י " + PublicFuncsNvars.curUser.getFullName()+"", new List<Tuple<byte[], string>>());
        }

        private bool signDoc(int docId)
        {
            if (!PublicFuncsNvars.alreadySigned(docId))
            {
                if (PublicFuncsNvars.docHasRecipients(docId))
                {
                    DialogResult result = MessageBox.Show("האם אתם בטוחים שהמסמך מוכן לחתימה?", "חתימה", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    if (DialogResult.Yes == result)
                    {
                        bool signed = PublicFuncsNvars.signDoc(docId);
                        if (signed)
                        {
                            MessageBox.Show("המסמך נחתם בהצלחה.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            return true;
                        }
                        else
                            return false;
                        
                    }
                    else
                        return false;
                }
                else
                {
                    MessageBox.Show("לא ניתן לחתום על מסמך ללא מכותבים.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return false;
                }
            }
            else
            {
                MessageBox.Show("לא ניתן לחתום על מסמך פעמיים.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return false;
            }
        }

        private void dataGridViewDocs_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                switch (e.ColumnIndex)
                {
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                        break;
                    default:
                        int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                        PublicFuncsNvars.openDocumentHandlingForm(id);
                        break;
                }
        }

        private void ForMeToSign_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Program.fmts = null;
            }
            catch { }
            
        }
    }
}
