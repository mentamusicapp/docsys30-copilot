using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Management;
using System.ComponentModel;
using System.Windows.Forms.Integration;

namespace DocumentsModule
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        internal static DocumentsMenu dm;
        //Thread newDocThread = null;
        internal static FoldersUpdate fu = null;
        internal static UsersUpdate uu = null;
        internal static RecipientListsUpdate rlu = null;
        internal static ForMeToSign fmts = null;
        internal static TransferToSign tts = null;
        internal static PublishSignedDocs psd = null;
        internal static MyDocsStatus mds = null;
        internal static Services srvs = null;
        internal static DragDropForm ddf = null;

        internal static Form wpfForm;
        internal static string folderPath = Application.StartupPath + "\\" + PublicFuncsNvars.userLogin + "\\temp";
        internal static string archiveFolder = Application.StartupPath + "\\" + PublicFuncsNvars.userLogin + "\\archive\\";
        //public bool afterSave = true;
        internal static bool haveArgs = false;
        //public bool afterDelete = false;
        [STAThread]
        static void Main()
        {
            
            if (!File.Exists(Global.IniFileName))
            {
                MessageBox.Show("קובץ " + Global.IniFileName + " לא קיים");
                Environment.Exit(0);
            }
            Global.INIvalues = File.ReadLines(Global.IniFileName).Where(line => (!String.IsNullOrWhiteSpace(line) && !line.StartsWith("#"))).Select(line => line.Split(new char[] { '=' }, 2, 0)).ToDictionary(parts => parts[0].Trim(), parts => parts.Length > 1 ? parts[1].Trim() : null);
        
            Global.P_SQL_SRV = Global.INIvalues["SQL_SRV"];
            Global.P_SQL_DB= Global.INIvalues["SQL_DB"];
            Global.P_MAX_CLASS = Global.INIvalues["MAX_CLASS"];
            Global.P_SQL_USR = Global.INIvalues["SQL_USR"];
            Global.P_SQL_PSW = Global.INIvalues["SQL_PSW"];  // UpdateINI.DecryptString(Global.INIvalues["SQL_PSW"]);
            Global.P_APP = Global.INIvalues["APP"];
            Global.P_LCL = Global.INIvalues["LCL"];

            folderPath = Global.P_LCL + "\\" + PublicFuncsNvars.userLogin + "\\temp";
            archiveFolder = Global.P_LCL + "\\" + PublicFuncsNvars.userLogin + "\\archive\\";
            if (Global.P_SQL_USR == "" || Global.P_SQL_PSW == "")//בדיקה האם להתחבר לפי משתמש.A.W 30/05/2024 
                Global.ConStr = "Data Source=" + Global.P_SQL_SRV + ";Initial Catalog=" + Global.P_SQL_DB + ";Integrated Security=True";//בניהולית ובסגולה מתחברים ככה
            else
                Global.ConStr = " Server=" + Global.P_SQL_SRV + ";Database = " + Global.P_SQL_DB + "; User ID=" + Global.P_SQL_USR + "; Password=" + Global.P_SQL_PSW + ";";//בארמי מתחברים ככה.
            // Persist Security Info=False;Encrypt=False;";
            //Data Source=Global.P_SQL_SRV;Initial Catalog=Global.P_SQL_DB

            Global.ReadOnly= Global.INIvalues["READ_ONLY"];

            string[] args = Environment.GetCommandLineArgs();

            if (args.Length == 3) haveArgs = true;
                // if (IsProcessOpen(Application.ProductName)) return;
            if (!haveArgs && IsProcessOpen("DocumentsModule"))
            {
                MessageBox.Show("התוכנית כבר רצה, אין אפשרות לפתוח אותה שוב", "Program Already Running", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
           // DocumentsModule
            //  if (IsProcessOpen()) return;
            System.IO.Directory.CreateDirectory(@"c:\temp\");
            System.IO.Directory.CreateDirectory(folderPath);
            System.IO.Directory.CreateDirectory(folderPath + "\\print");
            try
            {
                PublicFuncsNvars.getUsers();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.GetType().ToString() + ": " + e.Message + "; StackTrace: " + e.StackTrace, "DB Error");
                PublicFuncsNvars.saveLogError("Program", e.ToString(), e.Message);
            }
            PublicFuncsNvars.setArithabortOn();
            PublicFuncsNvars.getInterDist();
            PublicFuncsNvars.getFolders();
            PublicFuncsNvars.getProjects();
            PublicFuncsNvars.getRecipientsLists();
            PublicFuncsNvars.initializeCurrentUser();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //openWtf();

            if (args.Length == 3)
            {
                if (args[1] == "-s")
                {
                    //MessageBox.Show("compartir");
                    string ShotefToOpen = args[2];
                    dm = new DocumentsMenu(args[2]);

                }
            }
            else
            {
                //dm = new DocumentsMenu("533615");
                wpfForm= openWpf();
                //dm = new DocumentsMenu();
            }
            try
            {
                /*if (dm != null)
                    Application.Run(dm);
                else */if (wpfForm != null)
                    Application.Run(wpfForm);
                if (!haveArgs && System.IO.Directory.Exists(folderPath))
                    System.IO.Directory.Delete(folderPath, true);
            }
            catch (Exception e)
            {
                PublicFuncsNvars.saveLogError("Program", e.ToString(), e.Message);
            }
            finally
            {
                PublicFuncsNvars.releaseAllHeldDocs();
            }
        }


        //    private static bool IsProcessOpen()
        //    {

        //        if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
        //        {
        //            MessageBox.Show("התוכנית כבר רצה, אין אפשרות לפתוח אותה שוב", "Program Already Running", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return true;
        //        }


        //        return false;

        //    }
        //}

        private static bool IsProcessOpen(string name)
        {
            int counter = 0;
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.ToUpper().Contains(name.ToUpper()) && !clsProcess.ProcessName.ToUpper().Contains("VSHOST"))
                 //   if (clsProcess.ProcessName.ToUpper().CompareTo(name.ToUpper()) == 0)
                {
                    try
                    {
                       // MessageBox.Show(clsProcess.ProcessName);
                        var hndl = clsProcess.Handle;
                    }
                    catch (Win32Exception) { continue; }

                    if (++counter > 1)
                        return true;
                }

            return false;
        }



        private static Form openWpf()
        {
            View.UserControls.MainWindow wpfmain = new View.UserControls.MainWindow();
            wpfmain.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
            wpfmain.Arrange(new System.Windows.Rect(0, 0, wpfmain.DesiredSize.Width, wpfmain.DesiredSize.Height));

            ElementHost elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = wpfmain
            };

            wpfForm = new Form
            {
                AutoSize = true,
                StartPosition = FormStartPosition.CenterScreen
            };
            wpfForm.Controls.Add(elementHost);
            wpfForm.ClientSize = new Size(1200, 1000);
            wpfForm.WindowState = FormWindowState.Maximized;
            wpfForm.Icon = Global.AppIcon;
            wpfForm.Text = "מערכת מסמכים" + Global.Version;
            wpfForm.FormClosing += WpfForm_FormClosing;
            return wpfForm;
        }

        private static void WpfForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            /*if (newDocThread != null && MyGlobals.dragFlag == true)
            {

                if (newDocThread.IsAlive) newDocThread.Abort();
                MyGlobals.dragFlag = false;
                newDocThread = null;

            }

            if (newDocThread == null || !newDocThread.IsAlive)
            {*/
                Thread.Sleep(1000);
                if (PublicFuncsNvars.openDocs.Count > 0 || PublicFuncsNvars.openAtts.Count > 0 || PublicFuncsNvars.openVerDocs.Count > 0)
                {
                    MessageBox.Show("קיימים מסמכים פתוחים בתוכנה." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר קיימים מסמכים פתוחים."
                        + Environment.NewLine + "אנא סגרו את כל המסמכים הפתוחים ונסו שוב.", "אין אפשרות לסגירה",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                
                else if (srvs != null)
                {
                    MessageBox.Show("מסך שירות פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך שירות פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך שירות ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (fmts != null)
                {
                    MessageBox.Show("מסך מסמכים לחתימתי פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך מסמכים לחתימתי פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך מסמכים לחתימתי ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (tts != null)
                {
                    MessageBox.Show("מסך העברה לחתימה פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך העברה לחתימה פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך העברה לחתימה ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (mds != null)
                {
                    MessageBox.Show("מסך מסמכים שלי פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך מסמכים שלי פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך מסמכים שלי ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (psd != null)
                {
                    MessageBox.Show("מסך מסמכים להפצה פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך מסמכים להפצה פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך מסמכים להפצה ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (fu != null)
                {
                    MessageBox.Show("מסך עדכון תיקים פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך עדכון תיקים פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך עדכון תיקים ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (rlu != null)
                {
                    MessageBox.Show("מסך עדכון רשימות תפוצה פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך עדכון רשימות תפוצה פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך עדכון רשימות תפוצה ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (uu != null)
                {
                    MessageBox.Show("מסך עדכון משתמשים פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר מסך עדכון משתמשים פתוח."
                            + Environment.NewLine + "אנא סגרו את מסך עדכון משתמשים ונסו שוב.", "אין אפשרות לסגירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
            /*}

            else
            {
                MessageBox.Show("נא לסגור את שאר חלונות התוכנה לפני סגירת החלון הראשי.", "אין אפשרות לסגירה",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                e.Cancel = true;
            }*/
            //throw new NotImplementedException();
        }
    }
   
}
