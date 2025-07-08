using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;

namespace DocumentsModule
{
    public partial class Signature : Form
    {
        private bool isDrawing = false;
        private Point lastPoint;
        private Bitmap rawingBitmap;
        private Graphics graphics;

        public Signature()
        {
            InitializeComponent();
            InitializeDrawingBitmap();
            this.pictureBox1.MouseDown += new MouseEventHandler(pictureBox1_MouseDown);
            this.pictureBox1.MouseMove += new MouseEventHandler(pictureBox1_MouseMove);
            this.pictureBox1.MouseUp += new MouseEventHandler(pictureBox1_MouseUp);

        }

        private void InitializeDrawingBitmap()
        {
            rawingBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            graphics = Graphics.FromImage(rawingBitmap);
            graphics.Clear(Color.White);
            pictureBox1.Image = rawingBitmap;
        }
        private void pictureBox1_MouseDown(object ender, MouseEventArgs e)
        {
            isDrawing = true;
            lastPoint = e.Location;
        }
        private void pictureBox1_MouseMove(object ender, MouseEventArgs e)
        {
            if (isDrawing)
            {
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    graphics.DrawLine(pen, lastPoint, e.Location);
                }
                lastPoint = e.Location;
                pictureBox1.Invalidate();
            }
        }
        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            isDrawing = false;
        }

        private void BtnUpload_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Fil s|*. pg;*.jpeg;*.png;*bmp";
                openFileDialog.Title = " תבחר תמונה ";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    using (Image loadedImage = Image.FromFile(filePath))
                    {
                        using (Graphics g = Graphics.FromImage(rawingBitmap))
                        {
                            g.DrawImage(loadedImage, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
                        }
                        pictureBox1.Invalidate();
                    }
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT signature FROM dbo.tmtafkidu WHERE kod_tpkid=@userCode AND signature IS NOT NULL", conn);
            comm.Parameters.AddWithValue("@userCode", PublicFuncsNvars.curUser.userCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (!sdr.Read())
            {
                conn.Close();
                MessageBox.Show(" קיימת בעיה בקובץ החתימה שלך ." + Environment.NewLine + " המסמך לא נחתם .", " חתימה ",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                //return false; 
            }
            string picPath = Program.folderPath + "\\" + PublicFuncsNvars.curUser.userCode.ToString() + ".png";
            File.WriteAllBytes(picPath, sdr.GetSqlBytes(0).Buffer);
            conn.Close();
            using (Image loadedImage = Image.FromFile(picPath))
            {
                using (Graphics g = Graphics.FromImage(rawingBitmap))
                {
                    g.DrawImage(loadedImage, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
                }
                pictureBox1.Invalidate();
            }

        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            graphics.Clear(Color.White);
            pictureBox1.Invalidate();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            string picPath = Program.folderPath + "\\" + PublicFuncsNvars.curUser.userCode.ToString() + ".png";
            rawingBitmap.Save(picPath, System.Drawing.Imaging.ImageFormat.Png);
            this.Close();
        }

        private void Signature_Load(object sender, EventArgs e)
        {
            this.Icon = Global.AppIcon;
            DocumentsMenu.PathTemplate(BtnClear, 20);
            DocumentsMenu.PathTemplate(BtnOk, 20);
            DocumentsMenu.PathTemplate(BtnUpload, 20);
            DocumentsMenu.PathTemplate(button2, 20);
        }
    }
}
