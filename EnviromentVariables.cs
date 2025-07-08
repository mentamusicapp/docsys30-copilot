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
    public partial class EnviromentVariables : Form
    {
        public EnviromentVariables()
        {
            InitializeComponent();
        }

        private void EnviromentVariables_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
            foreach(KeyValuePair<string,string> key in Global.INIvalues)
            {
                textBox1.AppendText($"{key.Key}={key.Value}{Environment.NewLine}");
            }
        }
    }
}
