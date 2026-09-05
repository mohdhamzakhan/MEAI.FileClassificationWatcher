using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace MEAI.FileClassificationWatcher
{
    // A single custom-drawn icon (navy shield + gold padlock) shared by the tray icon and
    // every popup dialog's title bar. No external .ico asset needed — it's rendered once
    // via GDI+ and cached, so the taskbar and every window use the exact same mark instead
    // of Program.cs's SystemIcons.Shield (Windows' generic UAC shield) for the tray and
    // WinForms' own default icon for dialogs that never set one.
    public static class BrandIcon
    {
        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr handle);

        private static readonly Color BrandNavy = Color.FromArgb(0x1B, 0x26, 0x3B);
        private static readonly Color BrandNavyLight = Color.FromArgb(0x2E, 0x3F, 0x5C);
        private static readonly Color BrandGold = Color.FromArgb(0xD4, 0xAF, 0x37);

        private static Icon? _cached;

        // Icon.FromHandle wraps a raw GDI handle that has to be destroyed manually via
        // DestroyIcon or it leaks — cloning immediately gives a fully-managed Icon that
        // owns its own copy, so the original handle can be released right away. Cached
        // after the first call since every caller wants the identical icon anyway.
        public static Icon Create()
        {
            if (_cached != null) return _cached;

            using var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var b = new Rectangle(0, 0, 31, 31);

                PointF[] shieldPoints =
                {
                    new(b.Width * 0.06f, b.Height * 0.10f),
                    new(b.Width * 0.94f, b.Height * 0.10f),
                    new(b.Width * 0.94f, b.Height * 0.52f),
                    new(b.Width * 0.50f, b.Height * 1.00f),
                    new(b.Width * 0.06f, b.Height * 0.52f),
                };

                using (var path = new GraphicsPath())
                {
                    path.AddPolygon(shieldPoints);
                    using var fill = new LinearGradientBrush(b, BrandNavyLight, BrandNavy, LinearGradientMode.Vertical);
                    g.FillPath(fill, path);
                    using var pen = new Pen(Color.White, 1.2f);
                    g.DrawPath(pen, path);
                }

                float bodyW = b.Width * 0.34f, bodyH = b.Height * 0.28f;
                float bodyX = (b.Width - bodyW) / 2f, bodyY = b.Height * 0.50f;
                var bodyRect = new RectangleF(bodyX, bodyY, bodyW, bodyH);

                float shackleW = bodyW * 0.66f, shackleH = bodyH * 0.85f;
                float shackleX = (b.Width - shackleW) / 2f, shackleY = bodyY - shackleH * 0.62f;
                using (var shacklePen = new Pen(BrandGold, 2f))
                {
                    g.DrawArc(shacklePen, shackleX, shackleY, shackleW, shackleH, 180, 180);
                }

                using (var bodyPath = new GraphicsPath())
                {
                    bodyPath.AddRoundedRect(bodyRect, 2f);
                    using var bodyBrush = new SolidBrush(BrandGold);
                    g.FillPath(bodyBrush, bodyPath);
                }
            }

            var hIcon = bmp.GetHicon();
            try
            {
                _cached = (Icon)Icon.FromHandle(hIcon).Clone();
            }
            finally
            {
                DestroyIcon(hIcon);
            }
            return _cached;
        }
    }

    internal static class GraphicsPathExtensions
    {
        public static void AddRoundedRect(this GraphicsPath path, RectangleF rect, float radius)
        {
            float d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
        }
    }
}