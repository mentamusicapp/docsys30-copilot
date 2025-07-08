using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentsModule
{
    /*
     * this is a Label Class with an addition of outlining
     */
    public partial class CustomLabel : Label
    {
        internal Color outLineColor { get; set; }
        internal float outLineWidth { get; set; }
        public CustomLabel()
        {
            outLineColor = Color.Black;
            outLineWidth = 2f;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            pe.Graphics.FillRectangle(new SolidBrush(BackColor), ClientRectangle);
            using (GraphicsPath gp = new GraphicsPath())
            using (Pen outline = new Pen(outLineColor, outLineWidth) { LineJoin = LineJoin.Round })
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Far;
                sf.FormatFlags = StringFormatFlags.DirectionRightToLeft;
                using (Brush foreBrush = new SolidBrush(ForeColor))
                {
                    gp.AddString(Text, Font.FontFamily, (int)Font.Style, Font.Size, ClientRectangle, sf);
                    pe.Graphics.ScaleTransform(1.3f, 1.35f);
                    pe.Graphics.SmoothingMode = SmoothingMode.HighQuality;
                    pe.Graphics.DrawPath(outline, gp);
                    pe.Graphics.FillPath(foreBrush, gp);
                }
            }
        }
    }
}
