using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Security.Cryptography;

namespace DocumentsModule
{
    public partial class UpdateINI : Form
    {
        public UpdateINI()
        {
            InitializeComponent();
            this.Icon = Global.AppIcon;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string encryptedString=EncryptString(textBox1.Text);
            textBox2.Text = encryptedString;
            button3.Visible=true;
        }
        private string EncryptString(string text)
        {
            byte[] iv = new byte[16];
            byte[] array;
            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(Global.Key);
                aes.IV = iv;
                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter streamWriter = new StreamWriter((Stream)cryptoStream))
                        {
                            streamWriter.Write(text);
                        }
                        array = memoryStream.ToArray();
                    }
                }
            }

            return Convert.ToBase64String(array);
        }

        public static string DecryptString(string text)
        {
            byte[] iv = new byte[16];
            byte[] buffer = Convert.FromBase64String(text);
            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(Global.Key);
                aes.IV = iv;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
                using (MemoryStream memoryStream = new MemoryStream(buffer))
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader streamReader = new StreamReader((Stream)cryptoStream))
                        {
                            return streamReader.ReadToEnd();
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string decriptString=DecryptString(textBox1.Text);
            textBox2.Text = decriptString;
            button3.Visible = false;
        }

        private void UpdateINI_Load(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Global.INIvalues["SQL_PSW"] = textBox2.Text;
            Global.INIvalues.Remove("[SYSTEM]");
            using (StreamWriter writer = new StreamWriter(Global.IniFileName))
            {
                writer.WriteLine("[SYSTEM]");
                foreach (string key in Global.INIvalues.Keys)
                {
                    writer.WriteLine($"{key}={Global.INIvalues[key]}");
                }
            }
        }
    }
}
