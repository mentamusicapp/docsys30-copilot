using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using WinForms = System.Windows.Forms;
using System.Threading;


namespace DocumentsModule.View.UserControls
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Controls.UserControl
    {
        /*Thread newDocThread = null;
        internal FoldersUpdate fu = null;
        internal UsersUpdate uu = null;
        internal RecipientListsUpdate rlu = null;
        internal ForMeToSign fmts = null;
        internal TransferToSign tts = null;
        internal PublishSignedDocs psd = null;
        internal MyDocsStatus mds = null;
        internal Services srvs = null;
        internal DragDropForm ddf = null;*/
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ToolBar_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void DataGridDocs_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void DataGridDocs_Loaded(object sender, RoutedEventArgs e)
        {

        }
        public static void UpdateLabelresult(int numberofRows)
        {
            //MainWindow.Results.Content= "נמצאו " + numberofRows + " מסמכים";
        }

        private void MoveToSign_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (Program.tts == null)
            {
                Program.tts = new TransferToSign();
                Program.tts.Activate();
            }
            Program.tts.Show();
            Program.tts.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }

        private void NewDoc_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            NewDocument nd = new NewDocument();
            nd.Activate();
            nd.ShowDialog();
        }

        private void DocumentsSign_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (Program.fmts == null)
            {
                Program.fmts = new ForMeToSign();
                Program.fmts.Activate();
            }
            Program.fmts.Show();
            Program.fmts.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }

        private void Publish_Click(object sender, RoutedEventArgs e)
        {
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            if (Program.psd == null)
            {
                Program.psd = new PublishSignedDocs();
                Program.psd.Activate();
            }
            Program.psd.Show();
            Program.psd.BringToFront();
            WinForms.Cursor.Current = WinForms.Cursors.Default;
        }

        private void Home_Click(object sender, RoutedEventArgs e)
        {
            toolBarUC.searchUC.DefaultSearch();
        }
    }
}
