using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace DocumentsModule
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void SendEmailToContacts()
        {
            string subjectEmail = "my subject";
            string body = "my body";
            //Microsoft.Office.Interop.Outlook.MAPIFolder sentContacts = (Microsoft.Office.Interop.Outlook.MAPIFolder) this

            CreateEmailItem(subjectEmail, "DANIEL_NAVE@modnet.il", body);
        }

        private void CreateEmailItem(string subjectEmail,string toEmail,string bodyEmail)
        {
            Microsoft.Office.Interop.Outlook.MailItem eMail = new Microsoft.Office.Interop.Outlook.MailItem();
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Send();
        }
    }
}
