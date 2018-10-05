using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Cronos_Data
{
    public class SPanel : Panel
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.FillRoundedRectangle(new SolidBrush(Color.FromArgb(78, 122, 159)), 10, 10, Width - 40, Height - 60, 10);
            SolidBrush brush = new SolidBrush(
                Color.FromArgb(78, 122, 159)
                );
            g.FillRoundedRectangle(brush, 12, 12, Width - 44, Height - 64, 10);
            g.DrawRoundedRectangle(new Pen(ControlPaint.Light(Color.FromArgb(78, 122, 159), 0.00f)), 12, 12, Width - 44, Height - 64, 10);
            g.FillRoundedRectangle(new SolidBrush(Color.FromArgb(78, 122, 159)), 12, 12 + ((Height - 64) / 2), Width - 44, (Height - 64) / 2, 10);
        }
    }
}
