using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace PmwDeviceScan
{
    internal static class AppInfo
    {
        public const string Version = "2.0";
    }

    // Colours, fonts, glyphs and drawing helpers shared by the scan window and Settings.
    // Sizes are given in pixels as the design states them; fonts convert px to points.
    internal static class Ui
    {
        public static readonly Color Brand = Color.FromArgb(0x00, 0x78, 0xD4);
        public static readonly Color BrandHover = Color.FromArgb(0x10, 0x6E, 0xBE);
        public static readonly Color BrandPress = Color.FromArgb(0x00, 0x5A, 0x9E);
        public static readonly Color Ink = Color.FromArgb(0x1A, 0x1A, 0x1A);
        public static readonly Color Soft = Color.FromArgb(0x61, 0x61, 0x61);
        public static readonly Color Muted = Color.FromArgb(0xA0, 0xA0, 0xA0);
        public static readonly Color Waiting = Color.FromArgb(0x8A, 0x8A, 0x8A);
        public static readonly Color Line = Color.FromArgb(0xEB, 0xEB, 0xEB);
        public static readonly Color Border = Color.FromArgb(0xD1, 0xD1, 0xD1);
        public static readonly Color CardFill = Color.FromArgb(0xFB, 0xFB, 0xFB);
        public static readonly Color CardLine = Color.FromArgb(0xE5, 0xE5, 0xE5);
        public static readonly Color Tile = Color.FromArgb(0xF3, 0xF3, 0xF3);
        public static readonly Color Nav = Color.FromArgb(0xE9, 0xE9, 0xE9);
        public static readonly Color Footer = Color.FromArgb(0xF9, 0xF9, 0xF9);
        public static readonly Color Hover = Color.FromArgb(0xF0, 0xF0, 0xF0);
        public static readonly Color Press = Color.FromArgb(0xE5, 0xE5, 0xE5);
        public static readonly Color Picked = Color.FromArgb(0xE8, 0xF1, 0xFB);
        public static readonly Color Green = Color.FromArgb(0x0F, 0x7B, 0x0F);
        public static readonly Color Red = Color.FromArgb(0xC4, 0x2B, 0x1C);
        public static readonly Color Amber = Color.FromArgb(0x9D, 0x5D, 0x00);
        public static readonly Color Disabled = Color.FromArgb(0xC9, 0xD3, 0xDF);
        public static readonly Color Pink = Color.FromArgb(0xFD, 0xE7, 0xE9);
        public static readonly Color PinkLine = Color.FromArgb(0xF1, 0xBB, 0xBC);

        public const TextFormatFlags LeftLine = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
            | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis;
        public const TextFormatFlags RightLine = TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
            | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis;
        public const TextFormatFlags CenterLine = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
            | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;

        // Declared first: the family lookups below read it.
        private static readonly HashSet<string> Families = LoadFamilies();
        private static readonly string BodyFamily = Has("Segoe UI Variable Text") ? "Segoe UI Variable Text" : "Segoe UI";
        private static readonly bool HasSemibold = Has("Segoe UI Semibold");
        private static readonly string GlyphFamily = Has("Segoe Fluent Icons") ? "Segoe Fluent Icons"
            : Has("Segoe MDL2 Assets") ? "Segoe MDL2 Assets" : null;
        private static readonly Dictionary<string, Font> Fonts = new Dictionary<string, Font>();

        private static HashSet<string> LoadFamilies()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var installed = new InstalledFontCollection())
            {
                foreach (FontFamily f in installed.Families) set.Add(f.Name);
            }
            return set;
        }

        private static bool Has(string family)
        {
            return Families.Contains(family);
        }

        private static Font Get(string family, float px, FontStyle style)
        {
            string key = family + "|" + px.ToString("0.##") + "|" + (int)style;
            Font f;
            if (!Fonts.TryGetValue(key, out f))
            {
                f = new Font(family, px * 0.75f, style, GraphicsUnit.Point);
                Fonts[key] = f;
            }
            return f;
        }

        public static Font Body(float px) { return Get(BodyFamily, px, FontStyle.Regular); }

        public static Font Semi(float px)
        {
            return HasSemibold ? Get("Segoe UI Semibold", px, FontStyle.Regular) : Get(BodyFamily, px, FontStyle.Bold);
        }

        public static Font Mono() { return Get("Consolas", 12f, FontStyle.Regular); }

        public static void Draw(Graphics g, string text, Font f, Color c, Rectangle r, TextFormatFlags flags)
        {
            if (string.IsNullOrEmpty(text)) return;
            TextRenderer.DrawText(g, text, f, r, c, flags);
        }

        // Height a wrapped paragraph needs at this width.
        public static int Lines(string text, Font f, int width)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            return TextRenderer.MeasureText(text, f, new Size(width, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Height;
        }

        // Width of one line of text.
        public static int Width(string text, Font f)
        {
            return TextRenderer.MeasureText(text ?? "", f, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
        }

        public static Label AddLabel(Control parent, string text, int x, int y, int w, int h, Font f, Color c)
        {
            var l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.Size = new Size(w, h);
            l.Font = f;
            l.ForeColor = c;
            l.BackColor = parent.BackColor;
            parent.Controls.Add(l);
            return l;
        }

        public static GraphicsPath Rounded(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            if (radius <= 0)
            {
                p.AddRectangle(r);
                return p;
            }
            int d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        public static void FillRound(Graphics g, Rectangle r, int radius, Color c)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath p = Rounded(r, radius))
            using (var b = new SolidBrush(c))
            {
                g.FillPath(b, p);
            }
        }

        public static void StrokeRound(Graphics g, Rectangle r, int radius, Color c)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath p = Rounded(r, radius))
            using (var pen = new Pen(c))
            {
                g.DrawPath(pen, p);
            }
        }

        // Segoe Fluent Icons / Segoe MDL2 Assets codepoints. Names not listed here are drawn as plain text,
        // so "#" and "1" work as they are.
        private static int CodeOf(string name)
        {
            switch (name)
            {
                case "monitor": return 0xE7F4;
                case "folder": return 0xE8B7;
                case "gear": return 0xE713;
                case "pin": return 0xE718;
                case "tag": return 0xE8EC;
                case "key": return 0xE8D7;
                case "apps": return 0xE71D;
                case "check": return 0xE73E;
                case "processor": return 0xE950;
                case "office": return 0xE8A5;
                default: return 0;
            }
        }

        // Windows 7 has neither icon font: a plain letter stands in, so nothing shows as a box.
        private static string FallbackOf(string name)
        {
            switch (name)
            {
                case "monitor": return "PC";
                case "folder": return "F";
                case "gear": return "⚙";
                case "pin": return "P";
                case "tag": return "T";
                case "key": return "K";
                case "apps": return "≡";
                case "check": return "✓";
                case "processor": return "CPU";
                case "office": return "W";
                default: return name;
            }
        }

        public static void DrawGlyph(Graphics g, string name, Rectangle r, Color c, int px)
        {
            int code = CodeOf(name);
            if (code == 0)
            {
                Draw(g, name, Body(px), c, r, CenterLine);
                return;
            }
            if (GlyphFamily != null)
            {
                Draw(g, char.ConvertFromUtf32(code), Get(GlyphFamily, px, FontStyle.Regular), c, r, CenterLine);
                return;
            }
            bool symbol = name == "gear" || name == "apps" || name == "check";
            Font f = symbol ? Get("Segoe UI Symbol", px, FontStyle.Regular) : Body(px);
            Draw(g, FallbackOf(name), f, c, r, CenterLine);
        }
    }

    // A control that paints itself completely.
    internal abstract class Painted : Control
    {
        protected Painted()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint
                | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.White;
        }

        protected Color ParentColor()
        {
            return Parent != null ? Parent.BackColor : Color.White;
        }
    }

    // A rounded box, used for cards and the pink error box. Its children sit on its fill.
    internal sealed class Card : Painted
    {
        private Color _fill = Color.White;
        public Color Edge = Color.Empty;
        public int Radius = 6;

        public Color Fill
        {
            get { return _fill; }
            set
            {
                _fill = value;
                BackColor = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(ParentColor());
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            Ui.FillRound(g, r, Radius, _fill);
            if (Edge != Color.Empty) Ui.StrokeRound(g, r, Radius, Edge);
        }
    }

    // A glyph on its own, optionally on a rounded tile.
    internal sealed class GlyphView : Painted
    {
        public string Glyph = "";
        public Color Tint = Ui.Brand;
        public Color Fill = Color.Empty;
        public int Radius;
        public int Px = 16;

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            if (Fill != Color.Empty) Ui.FillRound(g, new Rectangle(0, 0, Width - 1, Height - 1), Radius, Fill);
            Ui.DrawGlyph(g, Glyph, new Rectangle(0, 0, Width, Height), Tint, Px);
        }
    }

    // A borderless icon button (folder in the header, gear in the footer).
    internal sealed class IconButton : Painted
    {
        public string Glyph = "";
        public Color Tint = Ui.Ink;
        private bool _hover;

        public IconButton()
        {
            Size = new Size(36, 36);
            Cursor = Cursors.Hand;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            if (Enabled && _hover) Ui.FillRound(g, new Rectangle(0, 0, Width - 1, Height - 1), 6, Ui.Press);
            Ui.DrawGlyph(g, Glyph, new Rectangle(0, 0, Width, Height), Enabled ? Tint : Ui.Muted, 16);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            Invalidate();
            base.OnEnabledChanged(e);
        }
    }

    // The window's top strip: a plain title (computer and model, or place) or a result (title, badge).
    internal sealed class HeaderBar : Painted
    {
        private bool _result;
        private string _line1 = "";
        private string _line2 = "";
        private string _title = "";
        private Color _titleColor = Ui.Green;
        private bool _ok;

        public HeaderBar()
        {
            Height = 60;
        }

        public void SetPlain(string line1, string line2)
        {
            _result = false;
            _line1 = line1 ?? "";
            _line2 = line2 ?? "";
            Invalidate();
        }

        public void SetResult(string title, Color color, string detail, bool ok)
        {
            _result = true;
            _title = title;
            _titleColor = color;
            _line2 = detail ?? "";
            _ok = ok;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            if (!_result)
            {
                Rectangle tile = new Rectangle(16, 12, 36, 36);
                Ui.FillRound(g, tile, 8, Ui.Brand);
                Ui.DrawGlyph(g, "monitor", tile, Color.White, 20);
                int tw = Width - 64 - 56;
                Ui.Draw(g, _line1, Ui.Semi(14), Ui.Ink, new Rectangle(64, 11, tw, 20), Ui.LeftLine);
                Ui.Draw(g, _line2, Ui.Body(12), Ui.Soft, new Rectangle(64, 31, tw, 18), Ui.LeftLine);
            }
            else
            {
                int tw = Width - 16 - 36 - 24;
                Ui.Draw(g, _title, Ui.Semi(18), _titleColor, new Rectangle(16, 8, tw, 24), Ui.LeftLine);
                Ui.Draw(g, _line2, Ui.Body(12), Ui.Soft, new Rectangle(16, 32, tw, 18), Ui.LeftLine);
                Rectangle badge = new Rectangle(Width - 16 - 36, 12, 36, 36);
                Ui.FillRound(g, badge, 18, _titleColor);
                if (_ok) Ui.DrawGlyph(g, "check", badge, Color.White, 20);
                else Ui.Draw(g, "!", Ui.Semi(20), Color.White, badge, Ui.CenterLine);
            }
            using (var pen = new Pen(Ui.Line)) g.DrawLine(pen, 0, Height - 1, Width, Height - 1);
        }
    }

    // The window's bottom strip: a status line (with a spinner while scanning and an optional
    // "Keep open" link) and the gear button.
    internal sealed class FooterBar : Painted
    {
        public string Status = "";
        public bool Spinning;
        public int Angle;
        public bool ShowKeep;
        public readonly IconButton Gear;
        public event EventHandler KeepOpenClicked;
        private Rectangle _keep = Rectangle.Empty;

        public FooterBar()
        {
            Height = 48;
            BackColor = Ui.Footer;
            Gear = new IconButton();
            Gear.Glyph = "gear";
            Gear.BackColor = Ui.Footer;
            Gear.Location = new Point(356, 6);
            Controls.Add(Gear);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            using (var pen = new Pen(Ui.Line)) g.DrawLine(pen, 0, 0, Width, 0);

            int x = 16;
            if (Spinning)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(Ui.Brand, 2f)) g.DrawArc(pen, new Rectangle(16, 16, 16, 16), Angle, 270f);
                x = 40;
            }
            int limit = Gear.Visible ? Gear.Left - 8 : Width - 16;
            string shown = ShowKeep ? Status + " · " : Status;
            Ui.Draw(g, shown, Ui.Body(12), Ui.Soft, new Rectangle(x, 0, Math.Max(0, limit - x), Height), Ui.LeftLine);

            _keep = Rectangle.Empty;
            if (ShowKeep)
            {
                int kx = x + Ui.Width(shown, Ui.Body(12));
                int kw = Ui.Width("Keep open", Ui.Body(12));
                if (kx + kw <= limit)
                {
                    _keep = new Rectangle(kx, 0, kw, Height);
                    Ui.Draw(g, "Keep open", Ui.Body(12), Ui.Brand, _keep, Ui.LeftLine);
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            Cursor = (!_keep.IsEmpty && _keep.Contains(e.Location)) ? Cursors.Hand : Cursors.Default;
            base.OnMouseMove(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (!_keep.IsEmpty && _keep.Contains(e.Location))
            {
                EventHandler h = KeepOpenClicked;
                if (h != null) h(this, EventArgs.Empty);
            }
            base.OnMouseClick(e);
        }
    }

    // One row of the activity list (scanning) or of the result list (completed).
    // Step rows: name, a sub line, a state on the right and a marquee bar while running.
    // Value rows: a small label over a value.
    internal sealed class ActivityRow : Painted
    {
        public string Glyph = "apps";
        public string Title = "";
        public string Sub = "";
        public Color SubColor = Ui.Soft;
        public bool TitleBold;
        public string State = "";
        public Color StateColor = Ui.Soft;
        public bool Progress;
        public int Phase;
        public bool ValueMode;

        public ActivityRow()
        {
            Height = 60;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            int tileY = (Height - 32) / 2;
            Rectangle tile = new Rectangle(16, tileY, 32, 32);
            Ui.FillRound(g, tile, 6, Ui.Tile);
            Ui.DrawGlyph(g, Glyph, tile, Ui.Brand, 16);

            int tx = 60;
            int right = Width - 16;
            if (ValueMode)
            {
                Ui.Draw(g, Title, Ui.Body(12), Ui.Soft, new Rectangle(tx, 7, right - tx, 16), Ui.LeftLine);
                Ui.Draw(g, Sub, Ui.Body(14), SubColor, new Rectangle(tx, 23, right - tx, 20), Ui.LeftLine);
                return;
            }

            int stateW = State.Length > 0 ? 84 : 0;
            int textW = right - tx - (stateW > 0 ? stateW + 8 : 0);
            Ui.Draw(g, Title, TitleBold ? Ui.Semi(14) : Ui.Body(14), Ui.Ink, new Rectangle(tx, 9, textW, 20), Ui.LeftLine);
            Ui.Draw(g, Sub, Ui.Body(12), SubColor, new Rectangle(tx, 30, right - tx, 16), Ui.LeftLine);
            if (stateW > 0)
                Ui.Draw(g, State, Ui.Body(12), StateColor, new Rectangle(right - stateW, 9, stateW, 20), Ui.RightLine);

            if (Progress)
            {
                Rectangle track = new Rectangle(tx, Height - 12, right - tx, 3);
                Ui.FillRound(g, track, 1, Ui.CardLine);
                int seg = track.Width * 45 / 100;
                float t = (Phase % 100) / 100f;
                int x = track.X - seg + (int)(t * (track.Width + seg));
                Rectangle bar = Rectangle.Intersect(track, new Rectangle(x, track.Y, seg, track.Height));
                if (!bar.IsEmpty) Ui.FillRound(g, bar, 1, Ui.Brand);
            }
        }
    }

    // The circular countdown on Scanned before: the arc shrinks as the seconds run out.
    internal sealed class CountRing : Painted
    {
        private int _seconds = 3;
        public int Total = 3;

        public CountRing()
        {
            Size = new Size(64, 64);
        }

        public int Seconds
        {
            get { return _seconds; }
            set
            {
                _seconds = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle ring = new Rectangle(2, 2, Width - 4, Height - 4);
            using (var track = new Pen(Ui.CardLine, 4f)) g.DrawEllipse(track, ring);
            float sweep = Total > 0 ? 360f * _seconds / Total : 0f;
            if (sweep > 360f) sweep = 360f;
            if (sweep > 0f)
            {
                using (var arc = new Pen(Ui.Brand, 4f)) g.DrawArc(arc, ring, -90f, sweep);
            }
            Ui.Draw(g, _seconds.ToString(), Ui.Semi(22), Ui.Brand, new Rectangle(0, 0, Width, Height), Ui.CenterLine);
        }
    }

    // One place on the drive: pin, name and how many reports it has. Clicking it fills the fields.
    internal sealed class PlaceRow : Painted
    {
        public string Place = "";
        public string Count = "";
        private bool _picked;
        private bool _hover;

        public PlaceRow()
        {
            Height = 44;
            Cursor = Cursors.Hand;
        }

        public bool Picked
        {
            get { return _picked; }
            set
            {
                _picked = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            Rectangle body = new Rectangle(0, 0, Width - 1, Height - 1);
            if (_picked) Ui.FillRound(g, body, 6, Ui.Picked);
            else if (_hover) Ui.FillRound(g, body, 6, Ui.Hover);
            Ui.DrawGlyph(g, "pin", new Rectangle(8, 13, 18, 18), Ui.Brand, 16);
            Ui.Draw(g, Place, Ui.Body(14), Ui.Ink, new Rectangle(34, 0, Width - 34 - 96, Height), Ui.LeftLine);
            Ui.Draw(g, Count, Ui.Body(12), Ui.Soft, new Rectangle(Width - 96, 0, 88, Height), Ui.RightLine);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }
    }

    // A left-column entry of the Settings window, with the blue pill when selected.
    internal sealed class NavItem : Painted
    {
        private bool _selected;
        private bool _hover;

        public NavItem()
        {
            Height = 36;
            BackColor = Ui.Tile;
            Cursor = Cursors.Hand;
        }

        public bool Selected
        {
            get { return _selected; }
            set
            {
                _selected = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            Rectangle body = new Rectangle(0, 0, Width - 1, Height - 1);
            if (_selected) Ui.FillRound(g, body, 4, Ui.Nav);
            else if (_hover) Ui.FillRound(g, body, 4, Ui.Hover);
            if (_selected) Ui.FillRound(g, new Rectangle(0, 10, 3, 16), 2, Ui.Brand);
            Ui.Draw(g, Text, Ui.Body(14), Ui.Ink, new Rectangle(12, 0, Width - 24, Height), Ui.LeftLine);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }
    }

    // The design's button: 36 px tall, 4 px corners. Primary is filled blue; secondary is outlined.
    // A disabled primary is grey, as before.
    internal sealed class RoundButton : Button
    {
        public bool Primary;
        private bool _hover;
        private bool _down;

        public RoundButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Font = Ui.Body(14);
            BackColor = Color.White;
            Size = new Size(120, 36);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Color parent = Parent != null ? Parent.BackColor : Color.White;
            g.Clear(parent);
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            Color text;
            if (Primary)
            {
                Color fill;
                if (!Enabled) fill = Ui.Disabled;
                else if (_down) fill = Ui.BrandPress;
                else if (_hover) fill = Ui.BrandHover;
                else fill = Ui.Brand;
                Ui.FillRound(g, r, 4, fill);
                text = Color.White;
            }
            else
            {
                Color fill;
                if (!Enabled) fill = Ui.Tile;
                else if (_down) fill = Ui.Press;
                else if (_hover) fill = Ui.Hover;
                else fill = Ui.CardFill;
                Ui.FillRound(g, r, 4, fill);
                Ui.StrokeRound(g, r, 4, Enabled ? Ui.Border : Ui.Line);
                text = Enabled ? Ui.Ink : Ui.Muted;
            }
            Ui.Draw(g, Text, Font, text, new Rectangle(0, 0, Width, Height), Ui.CenterLine);
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _down = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
    }

    // The on/off switch of the Settings window: 40 x 20, Space or Enter or a click toggles it.
    internal sealed class ToggleSwitch : Painted
    {
        private bool _checked;
        public event EventHandler CheckedChanged;

        public ToggleSwitch()
        {
            Size = new Size(40, 20);
            TabStop = true;
            Cursor = Cursors.Hand;
        }

        public bool Checked
        {
            get { return _checked; }
            set
            {
                if (_checked == value) return;
                _checked = value;
                Invalidate();
                EventHandler h = CheckedChanged;
                if (h != null) h(this, EventArgs.Empty);
            }
        }

        private void Toggle()
        {
            Checked = !Checked;
        }

        protected override bool IsInputKey(Keys key)
        {
            if (key == Keys.Space || key == Keys.Enter) return true;
            return base.IsInputKey(key);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (Enabled && (e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter))
            {
                Toggle();
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (Enabled && e.Button == MouseButtons.Left)
            {
                Focus();
                Toggle();
            }
            base.OnMouseDown(e);
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(ParentColor());
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle track = new Rectangle(0, 0, Width - 1, Height - 1);
            if (_checked)
            {
                Ui.FillRound(g, track, 10, Enabled ? Ui.Brand : Ui.Disabled);
                Ui.FillRound(g, new Rectangle(Width - 4 - 12, 4, 12, 12), 6, Color.White);
            }
            else
            {
                Ui.FillRound(g, track, 10, Color.White);
                Ui.StrokeRound(g, track, 10, Enabled ? Ui.Soft : Ui.Border);
                Ui.FillRound(g, new Rectangle(4, 4, 12, 12), 6, Enabled ? Ui.Soft : Ui.Border);
            }
            if (Focused) ControlPaint.DrawFocusRectangle(g, new Rectangle(0, 0, Width, Height));
        }
    }
}
