using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text.RegularExpressions;

namespace DocumentsModule
{
    public partial class DragDropForm : Form
    {
        Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
        public DragDropForm()
        {
            InitializeComponent();
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            DocumentsMenu.PathTemplate(this.button1, 35);
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {

                Explorer oExplorer = app.ActiveExplorer();
                Selection oSelection = oExplorer.Selection;
                if (oSelection.Count != 1) return;
                MailItem mi = (MailItem)oSelection[1];
                tb_Sender.Text = mi.SenderName;
                tb_Date.Text = mi.CreationTime.ToShortDateString();
                tb_subject.Text = mi.Subject;
                //   tb_body.Text = mi.Body;
                string body = Regex.Replace(mi.Body, @"[\r\n]+", "\r\n");
                tb_body.Text = body;
                if (mi.Subject.Contains("@בלמ\"ס")) combo_sivug.SelectedIndex = 0;
                else if (mi.Subject.Contains("@שמור")) combo_sivug.SelectedIndex = 1;
                else if (mi.Subject.Contains("@סודי@")) combo_sivug.SelectedIndex = 3;
                else if (mi.Subject.Contains("@סודי ביותר")) combo_sivug.SelectedIndex = 4;

                string size = mi.Size.ToString();
                if (mi.Size >= (1 << 30))
                    size = string.Format("{0}Gb", mi.Size >> 30);

                else if (mi.Size >= (1 << 20))
                    size = string.Format("{0}Mb", mi.Size >> 20);

                else if (mi.Size >= (1 << 10))
                    size = string.Format("{0}Kb", mi.Size >> 10);
                tb_Size.Text = size;

                foreach (Attachment item in mi.Attachments)
                {
                    List_Attachements.Items.Add(item.FileName);
                }
            }

            else
            {

                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Any())
                {
                    FileInfo fi = new FileInfo(files[0]);
                    string size = fi.Length.ToString();
                    if (fi.Length >= (1 << 30))
                        size = string.Format("{0}Gb", fi.Length >> 30);

                    else if (fi.Length >= (1 << 20))
                        size = string.Format("{0}Mb", fi.Length >> 20);

                    else if (fi.Length >= (1 << 10))
                        size = string.Format("{0}Kb", fi.Length >> 10);


                    tb_Size.Text = size;
                    tb_Name.Text = fi.Name;
                    tb_Date.Text = File.GetLastWriteTime(files[0]).ToShortDateString();

                }

            }
      
        }

        private void panel1_DragOver(object sender, DragEventArgs e)
        {
            ClearAll();
            e.Effect = DragDropEffects.Copy;
        }

        private void ClearAll()
        {
            tb_anaf.Clear();
            tb_Date.Clear();
            tb_Name.Clear();
            tb_Sender.Clear();
            tb_Size.Clear();
            tb_subject.Clear();
            tb_Tik.Clear();
            tb_HanhayaNumber.Clear();
            List_Attachements.Items.Clear();

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
