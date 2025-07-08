using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Concurrent;

namespace DocumentsModule
{
    public partial class Services : Form
    {
        //internal static DocumentsSearch ds = null;
        internal static FoldersUpdate fu = null;
        internal static UsersUpdate uu = null;
        internal static RecipientListsUpdate rlu = null;
        internal static ForMeToSign fmts = null;
        internal static TransferToSign tts = null;
        internal static PublishSignedDocs psd = null;
        internal static MyDocsStatus mds = null;
        //internal Services srvs = null;
        //internal DragDropForm ddf = null;
        public Services()
        {
            InitializeComponent();
            this.FormClosed += Services_FormClosed;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            //יערה שינתה ב-30.10.22 על מנת לעבור להשתמש בכל התכונה בקונייקשן סטרינג המוגר בפרופרטיס
            //SqlConnection conn = new SqlConnection("Data Source=modsql6p;Initial Catalog=MantakDB;Integrated Security=True");
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=0 WHERE whoOpenedForEdit<>0", conn);
            //   SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=0 WHERE shotef_mismach=520858 OR shotef_mismach=520857", conn);
            conn.Open();
            comm.CommandTimeout = 0;
            comm.ExecuteNonQuery();
            int RowCount = 0;
            comm = new SqlCommand("SELECT @@ROWCOUNT", conn);
            //comm.ExecuteNonQuery();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                RowCount = sdr.GetInt32(0);

            }
            conn.Close();
            MessageBox.Show(RowCount + " " + "מסמכים שוחררו בהצלחה", "מסמכים שוחררו", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        private void button2_Click(object sender, EventArgs e)
        { }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox("הכנס קוד משתמש ברשת", "התחבר בשם אחר", "");
            PublicFuncsNvars.curUser = PublicFuncsNvars.getUserFromLogIn(input.Trim());
            if (PublicFuncsNvars.curUser == null)
            {
                MessageBox.Show("היוזר לא קיים, אנא נסה משתמש אחר");
                PublicFuncsNvars.curUser = PublicFuncsNvars.getUserFromLogIn(PublicFuncsNvars.userLogin);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DocumentsExport de = new DocumentsExport();
            de.Activate();
            de.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DocumenstImport dI = new DocumenstImport();
            dI.Activate();
            dI.ShowDialog();
        }

        private void Services_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
        }

        private void Services_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Program.srvs = null;
            }
            catch { }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DocumentsTemplates dt = new DocumentsTemplates();
            dt.Activate();
            dt.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ExecSqlScript ess = new ExecSqlScript();
            ess.Activate();
            ess.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            EnviromentVariables ev = new EnviromentVariables();
            ev.Activate();
            ev.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UpdateINI ui = new UpdateINI();
            ui.Activate();
            ui.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //UsersUpdate uu = new UsersUpdate();
            if (uu == null)
            {
                uu = new UsersUpdate();
                uu.Activate();
            }
            uu.Show();
            uu.BringToFront();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Program.rlu == null)
            {
                Program.rlu = new RecipientListsUpdate(null, false);
                Program.rlu.Activate();
            }
            Program.rlu.Show();
            Program.rlu.BringToFront();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (Program.fu == null)
            {
                Program.fu = new FoldersUpdate();
                Program.fu.Activate();
            }
            Program.fu.Show();
            Program.fu.BringToFront();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SaveTxt st = new SaveTxt();
            st.Activate();
            st.ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog f = new OpenFileDialog())
            {
                f.Filter = "SQL files(*.sql)|*.sql|All files (*.*)|*.*";
                f.Title = "תבחר SQL קובץ.";
                if (f.ShowDialog() == DialogResult.OK)
                {
                    string[] commands = File.ReadAllLines(f.FileName);
                    ConcurrentQueue<string> logQueue = new ConcurrentQueue<string>();
                    Parallel.For(0,commands.Length, i =>
                     {
                         string commandText = commands[i]?.Trim();
                         int lineNumber = i + 1;
                         if (i % 100 == 0) Console.WriteLine("================LINE NUMBER" + i+"=======================");
                         if (string.IsNullOrWhiteSpace(commandText))
                             return;
                         try
                         {
                             using (SqlConnection connection = new SqlConnection(Global.ConStr))
                             {
                                 connection.Open();
                                 using (SqlCommand command = new SqlCommand(commandText, connection))
                                 {
                                     command.ExecuteNonQuery();
                                 }

                             }
                         }

                         catch (Exception ee)
                         {
                             
                            // lock (logfile)
                           //  {
                                 logQueue.Enqueue($@"Line #{lineNumber}: {ee.Message} {Environment.NewLine}");
                                 //File.AppendAllText(logfile, $@"Line #{lineNumber}: {ee.Message} {Environment.NewLine}");
                            // }
                         }
                     });
                    string logfile = $@"C:\temp\DocExp\{Path.GetFileNameWithoutExtension(f.FileName)}.log";
                    File.AppendAllLines(logfile, logQueue);
                    Console.WriteLine("FINSIHED!!!!");
                    //SqlConnection conn = new SqlConnection(Global.ConStr);
                    //List<string> lines = File.ReadAllLines(f.FileName).ToList();

                    //for (int i = 0; i < lines.Count(); i++)
                    //{
                    //    var x = lines[i];
                    //    try {
                           
                    //        SqlCommand comm = new SqlCommand(lines[i], conn);
                    //        conn.Open();
                    //        comm.CommandTimeout = 0;
                    //        comm.ExecuteNonQuery();
                    //        conn.Close();
                    //    }

                    //    catch (Exception ee)
                    //    {
                    //        File.AppendAllText($@"C:\temp\DocExp\{Path.GetFileNameWithoutExtension(f.FileName)}.log",$@"Line #{i+1}: {ee.Message} {Environment.NewLine}");
                    //        // MessageBox.Show(ee.Message);
                    //        conn.Close();
                    //        continue;
                    //    }
                    //}
                }
                   
            }
        }
    }
}
//            DocumentsModule.MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter dsta =
//              new DocumentsModule.MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter();
//            DocumentsModule.MantakDBDataSetDocuments.documentsSignDataTable dsdt = dsta.GetDataAllTransferedToSign(false, "", "20180102", true);
//            DocumentsModule.MantakDBDataSetDocumentsTableAdapters.tmtafkiduTableAdapter usersAdapter =
//                new DocumentsModule.MantakDBDataSetDocumentsTableAdapters.tmtafkiduTableAdapter();
//            Dictionary<int, string> docsToSend = new Dictionary<int, string>();
//            int user = 0;
//            DocumentsModule.MantakDBDataSetDocuments.tmtafkiduRow userRow;
//            List<DocumentsModule.MantakDBDataSetDocuments.documentsSignRow> dsdtRows = dsdt.Where(x => includedInPilot(x.user_metaiek)).ToList();
//            foreach (DocumentsModule.MantakDBDataSetDocuments.documentsSignRow row in dsdtRows)
//            {
//                int tempUser;
//                if (!int.TryParse(row.user_metaiek.Trim(), out tempUser))
//                    tempUser = 0;
//                if (user != tempUser)
//                {
//                    if (user != 0)
//                    {
//                        userRow = usersAdapter.GetAUser(user).First();
//                        string body = "";
//                        foreach (KeyValuePair<int, string> doc in docsToSend)
//                            body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                        DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                            "יש לך " + docsToSend.Count + " מסמכים שלא העברת לחתימה.", body, null);
//                    }
//                    user = tempUser;
//                    docsToSend.Clear();
//                }
//                docsToSend.Add(row.shotef_mismach, row.hanadon);
//            }
//            if (docsToSend.Any())
//            {
//                userRow = usersAdapter.GetAUser(user).First();
//                string body = "";
//                foreach (KeyValuePair<int, string> doc in docsToSend)
//                    body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                    "יש לך " + docsToSend.Count + " מסמכים שלא העברת לחתימה.", body, null);
//            }

//            DocumentsModule.MantakDBDataSetDocumentsTableAdapters.documentsForUpdatesTableAdapter dta =
//                new DocumentsModule.MantakDBDataSetDocumentsTableAdapters.documentsForUpdatesTableAdapter();
//            DocumentsModule.MantakDBDataSetDocuments.documentsForUpdatesDataTable ddt = dta.GetDataAllToSign(true, false);
//            List<DocumentsModule.MantakDBDataSetDocuments.documentsForUpdatesRow> ddtRows = ddt.Where(x => includedInPilot(x.kod_sholeah.ToString())).ToList();
//            user = 0;
//            foreach (DocumentsModule.MantakDBDataSetDocuments.documentsForUpdatesRow row in ddtRows)
//            {
//                int tempUser = row.kod_sholeah;
//                if (user != tempUser)
//                {
//                    if (user != 0)
//                    {
//                        userRow = usersAdapter.GetAUser(user).First();
//                        string body = "";
//                        foreach (KeyValuePair<int, string> doc in docsToSend)
//                            body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                        DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                            "יש לך " + docsToSend.Count + " מסמכים המחכים לחתימתך.", body, null);
//                    }
//                    user = tempUser;
//                    docsToSend.Clear();
//                }
//                docsToSend.Add(row.shotef_mismach, row.hanadon);
//            }
//            if (docsToSend.Any())
//            {
//                userRow = usersAdapter.GetAUser(user).First();
//                string body = "";
//                foreach (KeyValuePair<int, string> doc in docsToSend)
//                    body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                    "יש לך " + docsToSend.Count + " מסמכים המחכים לחתימתך.", body, null);
//            }

//            ddt = dta.GetDataAllToPublish(true, false);
//            ddtRows = ddt.Where(x => includedInPilot(x.kod_sholeah.ToString())).ToList();
//            user = 0;
//            foreach (DocumentsModule.MantakDBDataSetDocuments.documentsForUpdatesRow row in ddtRows)
//            {
//                int tempUser = row.kod_sholeah;
//                if (user != tempUser)
//                {
//                    if (user != 0)
//                    {
//                        userRow = usersAdapter.GetAUser(user).First();
//                        string body = "";
//                        foreach (KeyValuePair<int, string> doc in docsToSend)
//                            body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                        DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                            "יש לך " + docsToSend.Count + " מסמכים חתומים שלא הופצו.", body, null);
//                    }
//                    user = tempUser;
//                    docsToSend.Clear();
//                }
//                docsToSend.Add(row.shotef_mismach, row.hanadon);
//            }
//            if (docsToSend.Any())
//            {
//                userRow = usersAdapter.GetAUser(user).First();
//                string body = "";
//                foreach (KeyValuePair<int, string> doc in docsToSend)
//                    body += "שוטף " + doc.Key + ": " + doc.Value + Environment.NewLine;
//                DocumentsModule.PublicFuncsNvars.sendMail("mntkmihshuv@modnet.il;מנת\"ק - תקלות", userRow.doal.Trim(), null, null,
//                    "יש לך " + docsToSend.Count + " מסמכים חתומים שלא הופצו.", body, null);
//            }
//        }

//        private static bool includedInPilot(string user)
//        {
//            int intUser;
//            if (int.TryParse(user, out intUser))
//            {
//                DocumentsModule.MantakDBDataSetDocumentsTableAdapters.tmtafkiduTableAdapter usersAdapter =
//                new DocumentsModule.MantakDBDataSetDocumentsTableAdapters.tmtafkiduTableAdapter();
//                if (usersAdapter.GetAUser(intUser).Any())
//                {
//                    char first = user[0];
//                    if (!(first == '5' || first == '1' || first == '7' || first == '2' || intUser == 404))
//                        return true;
//                }
//            }
//            return false;
//        }
//    }
//}
