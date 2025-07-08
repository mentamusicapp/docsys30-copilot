using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Windows.Forms.Integration;

namespace DocumentsModule
{
    public partial class DocumentsMenu : Form
    {
        Thread newDocThread=null;
        internal DocumentsSearch ds = null;
        internal FoldersUpdate fu = null;
        internal UsersUpdate uu = null;
        internal RecipientListsUpdate rlu = null;
        internal ForMeToSign fmts = null;
        internal TransferToSign tts = null;
        internal PublishSignedDocs psd = null;
        internal MyDocsStatus mds = null;
        internal Services srvs = null;
        internal DragDropForm ddf = null;
        private string v;
        public DocumentsMenu()
        {
            InitializeComponent();
            this.Text = this.Text + Global.Version;

            PathTemplate(this.button1, 55);
            PathTemplate(this.button2, 55);
            PathTemplate(this.button6, 55);
            PathTemplate(this.button7, 55);
            PathTemplate(this.button8, 55);
            PathTemplate(this.button9,55);
            this.button10.FlatStyle = FlatStyle.Flat;
            button10.FlatAppearance.BorderSize = 0;
        }

        internal static void PathTemplate(Button button,int radius)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.Paint += (sender, e) =>
            {
                GraphicsPath path = new GraphicsPath();
                Rectangle bounds = button.ClientRectangle;
                path.AddArc(bounds.X, bounds.Y, radius, radius, 180, 90);
                path.AddArc(bounds.Right - radius, 0, radius, radius, 270, 90);
                path.AddArc(bounds.Right - radius, bounds.Bottom - radius, radius, radius, 0, 90);
                path.AddArc(bounds.X, bounds.Bottom - radius, radius, radius, 90, 90);
                path.CloseFigure();
                button.Region = new Region(path);
                using (Pen pen = new Pen(Color.White, 5))
                {
                    button.FlatAppearance.BorderSize = 0;
                    e.Graphics.DrawPath(pen, path);
                }
            };
        }

        public DocumentsMenu(string shotef)
        {
            MyGlobals.haveArgs = true;
            InitializeComponent();
            this.Text = this.Text + Global.Version;
            int id = int.Parse(shotef);
            var ds = new DocumentsSearch(shotef);
            this.Close();
        
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ds == null)
            {
                ds = new DocumentsSearch();
                ds.Activate();
            }
            ds.Show();
            ds.BringToFront();
        }

        private void DocumentsMenu_Load(object sender, EventArgs e)
        {
            this.Icon= Global.AppIcon;

        }

        private void button2_Click(object sender, EventArgs e) 
        {

            if (newDocThread != null && MyGlobals.dragFlag == true)
            {
                if (newDocThread.IsAlive) newDocThread.Abort();
                MyGlobals.dragFlag = false;
                newDocThread = null;

            }
            
                
            
            if (newDocThread == null || !newDocThread.IsAlive)
            {
                newDocThread = new Thread(openNewDocForm);
                newDocThread.SetApartmentState(ApartmentState.STA);
                newDocThread.Start();
                //newDocThread = null;//אהבה הוסיפה כדי שיהיה אפשר לפתוח כמה פעמים מסמך חדש.
                
            }
            else
            {
                MessageBox.Show("ניתן לפתוח עד מסמך חדש אחד בו זמנית.", "אין אפשרות לפתיחה", MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            }
        }

        private void openNewDocForm(object state)
        {
            NewDocument nd = new NewDocument();
            nd.Activate();
            nd.ShowDialog();
        }

        private void DocumentsMenu_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (newDocThread != null && MyGlobals.dragFlag == true)
            {
               
                if (newDocThread.IsAlive) newDocThread.Abort();
                MyGlobals.dragFlag = false;
                newDocThread = null;

            }

            if (newDocThread == null || !newDocThread.IsAlive)
            {
                Thread.Sleep(1000);
                if (PublicFuncsNvars.openDocs.Count > 0 || PublicFuncsNvars.openAtts.Count > 0|| PublicFuncsNvars.openVerDocs.Count > 0)
                {
                    MessageBox.Show("קיימים מסמכים פתוחים בתוכנה." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר קיימים מסמכים פתוחים."
                        + Environment.NewLine + "אנא סגרו את כל המסמכים הפתוחים ונסו שוב.", "אין אפשרות לסגירה",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    e.Cancel = true;
                }
                else if (ds != null)
                {
                    MessageBox.Show("חיפוש מסמכים פתוח." + Environment.NewLine + "לא ניתן לסגור את התוכנה כאשר חיפוש מסמכים פתוח."
                            + Environment.NewLine + "אנא סגרו את חיפוש מסמכים ונסו שוב.", "אין אפשרות לסגירה",
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
            }
            
            else
            {
                MessageBox.Show("נא לסגור את שאר חלונות התוכנה לפני סגירת החלון הראשי.", "אין אפשרות לסגירה",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                e.Cancel = true;
            }
        }
        

        private void button6_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (fmts == null)
            {
                fmts = new ForMeToSign();
                fmts.Activate();
            }
            fmts.Show();
            fmts.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (tts == null)
            {
                tts = new TransferToSign();
                tts.Activate();
            }
            tts.Show();
            tts.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (psd == null)
            {
                psd = new PublishSignedDocs();
                psd.Activate();
            }
            psd.Show();
            psd.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (mds == null)
            {
                mds = new MyDocsStatus();
                mds.Activate();
            }
            mds.Show();
            mds.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void btnScanHelp_Click(object sender, EventArgs e)
        {
            Process p = new Process();
            p.StartInfo.FileName = "ScanConfig.pdf";
            p.Start();
        }

        private void button10_Click(object sender, EventArgs e)
        {

            //MIT
            srvs = null;
            string Input = Interaction.InputBox("נא להכניס סיסמת מנהל", "סיסמא");
            if (Input == "") return;
            if (Input != "MIT")
            {
                MessageBox.Show("סיסמא שגויה","", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }
            
            Cursor.Current = Cursors.WaitCursor;
            if (srvs == null)
            {
                srvs = new Services();
                srvs.Activate();
            }
            srvs.Show();
            srvs.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void DocumentsMenu_Enter(object sender, EventArgs e)
        {
            
        }

        private void DocumentsMenu_Activated(object sender, EventArgs e)
        {
            loggedUser.Text = PublicFuncsNvars.getCurUser(); 
        }

        private void btnTemplates_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (ddf == null)
            {
                ddf = new DragDropForm();
                ddf.Activate();
            }

            ddf.Show();
            ddf.BringToFront();
            Cursor.Current = Cursors.Default;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }
    }
}
