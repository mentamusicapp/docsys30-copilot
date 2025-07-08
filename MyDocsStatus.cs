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
    public partial class MyDocsStatus : Form
    {
        public MyDocsStatus()
        {
            InitializeComponent();
        }

        private void MyDocsStatus_Load(object sender, EventArgs e)
        {

            this.Icon = Global.AppIcon;
            int userCode=PublicFuncsNvars.curUser.userCode;
            MantakDBDataSetDocuments.documentsSignDataTable transferTable = (new MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter()).
                GetDataTransferToSign(false, userCode.ToString(), true);
            MantakDBDataSetDocuments.documentsSignDataTable signTable = (new MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter()).
                GetDataForMeToSign(userCode, true, false, false, true);
            MantakDBDataSetDocuments.documentsSignDataTable publishTable = (new MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter()).
                GetDataToPublishPerson(false, true, userCode, true);

            addToDocs(transferTable, "בעריכה");
            addToDocs(signTable, "ממתין לחתימה");
            addToDocs(publishTable, "ממתין להפצה");
        }

        private void addToDocs(MantakDBDataSetDocuments.documentsSignDataTable table, string status)
        {
            Color c;
            if (status == "בעריכה")
                c = Color.SkyBlue;
            else if (status == "ממתין לחתימה")
                c = Color.LimeGreen;
            else
                c = Color.Olive;
            foreach (MantakDBDataSetDocuments.documentsSignRow row in table.Rows)
            {
                int index = dataGridViewDocs.Rows.Add(row.shotef_mismach, row.hanadon.Trim(), row.teur_tafkid_sholeah.Trim(),
                    DateTime.ParseExact(row.tarich_hamichtav, "yyyyMMdd", CultureInfo.CurrentCulture), status);
                dataGridViewDocs.Rows[index].DefaultCellStyle.BackColor = c;
            }
        }

        private void MyDocsStatus_FormClosed(object sender, FormClosedEventArgs e)
        {
            Program.mds = null;
        }

        private void dataGridViewDocs_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int id = (int)dataGridViewDocs.Rows[e.RowIndex].Cells["docIdColumn"].Value;
                PublicFuncsNvars.openDocumentHandlingForm(id);
            }
        }
    }
}
