using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using WinForms= System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using Microsoft.VisualBasic;

namespace DocumentsModule.View.UserControls
{
    /// <summary>
    /// Interaction logic for ToolBar.xaml
    /// </summary>
    public partial class ToolBar : UserControl
    {
        /*Thread newDocThread = null;
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
        public static int Top = 1000;*/
        public ToolBar()
        {
            InitializeComponent();
           // btnMenu.ContextMenu = (ContextMenu)this.Resources["contextMenu"];
            //Tokzaot.SelectedIndex = 1;
            
        }

        private void btnMenu_Click(object sender, RoutedEventArgs e)
        {
            //btnMenu.ContextMenu.IsOpen = true;
        }

        /*private void NewDocument_Click(object sender, RoutedEventArgs e)
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
                WinForms.MessageBox.Show("ניתן לפתוח עד מסמך חדש אחד בו זמנית.", "אין אפשרות לפתיחה", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Exclamation,
                    WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);

            }
        }*/

        private void openNewDocForm(object state)
        {
            NewDocument nd = new NewDocument();
            nd.Activate();
            nd.ShowDialog();
        }

        /*private void PasstoSign_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (tts == null)
            {
                tts = new TransferToSign();
                tts.Activate();
            }
            tts.Show();
            tts.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }
        private void ToSign_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (fmts == null)
            {
                fmts = new ForMeToSign();
                fmts.Activate();
            }
            fmts.Show();
            fmts.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }
        private void ToPublish_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (psd == null)
            {
                psd = new PublishSignedDocs();
                psd.Activate();
            }
            psd.Show();
            psd.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }
        private void Services_Click(object sender, RoutedEventArgs e)
        {
            //MIT
            srvs = null;
            string Input = Interaction.InputBox("נא להכניס סיסמת מנהל", "סיסמא");
            if (Input == "") return;
            if (Input != "MIT")
            {
                WinForms.MessageBox.Show("סיסמא שגויה", "", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Exclamation, WinForms.MessageBoxDefaultButton.Button1,
                    WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                return;
            }

            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (srvs == null)
            {
                srvs = new Services();
                srvs.Activate();
            }
            srvs.Show();
            srvs.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }*/
        

        private void btnSettings_Click(object sender, RoutedEventArgs e)
        {
            //MIT
            Program.srvs = null;
            string Input = Interaction.InputBox("נא להכניס סיסמת מנהל", "סיסמא");
            if (Input == "") return;
            if (Input != "MIT")
            {
                WinForms.MessageBox.Show("סיסמא שגויה", "", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Exclamation, WinForms.MessageBoxDefaultButton.Button1,
                    WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                return;
            }

            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (Program.srvs == null)
            {
                Program.srvs = new Services();
                Program.srvs.Activate();
            }

           // PublicFuncsNvars.getFolders();
            Program.srvs.Show();
            Program.srvs.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }

        private void Person_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //PublicFuncsNvars.getCurUser()
            //MessageBox.Show(PublicFuncsNvars.userLogin,"משתמש נוכחי");
            MessageBox.Show(PublicFuncsNvars.getCurUser(), "משתמש נוכחי");
        }
    }
}
