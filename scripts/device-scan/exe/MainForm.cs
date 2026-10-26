using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PmwDeviceScan
{
    // What a finished scan hands back to the window.
    internal sealed class ScanOutcome
    {
        public string Path;
        public List<string> Lines;
    }

    internal sealed class MainForm : Form
    {
        private static readonly string[] StepNames = { "Hardware", "Serial and part numbers", "Licences", "Installed software" };
        private static readonly string[] StepGlyphs = { "processor", "#", "key", "apps" };
        private static readonly string[] StepRunSubs = { "Processor, memory, drives, network", "Reading\u2026",
            "Asking Windows\u2026 can take a minute", "Listing programs\u2026" };
        private static readonly string[] StepDoneSubs = { "Processor, memory, drives, network", "Serial and part numbers read",
            "Licences read", "Programs listed" };
        private static readonly string[] DoneNames = { "Serial number", "Product number", "Windows licence", "Office", "Installed software" };
        private static readonly string[] DoneGlyphs = { "#", "tag", "key", "office", "apps" };

        private enum Mode { None, KnownCountdown, DoneCountdown }

        private readonly Timer _timer = new Timer();
        private readonly Timer _anim = new Timer();
        private readonly ToolTip _tip = new ToolTip();
        private readonly string _demoPage;
        private readonly string _computer = ReportFiles.Computer();
        private readonly string _model = ModelName();
        private readonly List<PlaceRow> _placeRows = new List<PlaceRow>();

        private Mode _mode;
        private int _countdown;
        private string _knownPlace = "";
        private string _donePath = "";
        private BackgroundWorker _worker;
        private DateTime _scanStart;
        private int _scanSeconds;
        private int _phase;
        private bool _settingsFromDone;

        private Panel _new, _known, _scanPage, _done, _error;
        private Panel[] _pages;
        private HeaderBar _header;
        private IconButton _folderButton;
        private FooterBar _footer;
        private ContextMenuStrip _menu;
        private ToolStripMenuItem _menuSettings;

        private ComboBox _location, _department;
        private Panel _placeList;
        private Label _placeNone, _savesAs;
        private RoundButton _scan;

        private Card _knownCard;
        private GlyphView _knownPin;
        private Label _knownPlaceLabel, _knownFileLabel, _startingLabel;
        private CountRing _ring;
        private RoundButton _scanNow, _changePlace;

        private ActivityRow[] _stepRows;
        private ActivityRow[] _doneRows;
        private GlyphView _doneFolderIcon;
        private Label _doneFolder, _doneNote;
        private LinkLabel _openFolder;
        private RoundButton _close;

        private Card _errorBox;
        private Label _errorText, _whatToDo, _techText;
        private GlyphView[] _stepDots;
        private Label[] _stepTexts;
        private LinkLabel _techLink;
        private RoundButton _errorClose;

        public MainForm() : this(null) { }

        // demoPage is set only by the hidden --ui-demo switch: it shows one screen with sample
        // values and does not scan, save or start any timer that would close the window.
        public MainForm(string demoPage)
        {
            _demoPage = demoPage;
            Text = "PMW Device Scan";
            ClientSize = new Size(400, 588);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            ForeColor = Ui.Ink;
            Font = Ui.Body(14);

            _timer.Interval = 1000;
            _timer.Tick += OnTick;
            _anim.Interval = 60;
            _anim.Tick += OnAnimTick;

            BuildHeader();
            BuildFooter();
            BuildMenu();
            BuildNew();
            BuildKnown();
            BuildScanning();
            BuildDone();
            BuildError();
            _pages = new[] { _new, _known, _scanPage, _done, _error };
            _header.SetPlain(_computer, _model);

            if (demoPage != null)
            {
                Shown += delegate { ShowDemo(demoPage); };
                return;
            }

            // Started from Shown, not here: a scan that starts by itself must not
            // begin before the window exists to report on it.
            // Holding Shift while the program opens goes straight to Settings: with
            // "Start scanning by itself" on, the start screens are never shown.
            bool shift = (ModifierKeys & Keys.Shift) == Keys.Shift;
            Shown += delegate { if (shift) ShowSettings(); else Begin(); };
        }

        // The first screen: a machine scanned before, or a new one. Run again
        // after Settings, since a new folder can make a known machine new.
        private void Begin()
        {
            string previous = ReportFiles.FindPreviousFile();
            if (previous != null && ReportFiles.Current.AutoStart)
            {
                StartScan(ReportFiles.PlaceOfFile(previous));
            }
            else if (previous != null)
            {
                ShowKnown(previous);
            }
            else
            {
                FillNewChoices();
                SetPlace("", "");
                ShowPage(_new);
            }
        }

        // "Reports" or the folder set on the Settings screen, as people will read it.
        private static string FolderShown()
        {
            string f = ReportFiles.Current.ReportsFolder;
            return string.IsNullOrEmpty(f) ? "Reports" : f.TrimEnd('\\');
        }

        // ------------------------------------------------------------------
        //  Building the window
        // ------------------------------------------------------------------

        private void BuildHeader()
        {
            _header = new HeaderBar();
            _header.Location = new Point(0, 0);
            _header.Size = new Size(400, 60);
            _folderButton = new IconButton();
            _folderButton.Glyph = "folder";
            _folderButton.Tint = Ui.Ink;
            _folderButton.BackColor = Color.White;
            _folderButton.Location = new Point(348, 12);
            _folderButton.Size = new Size(36, 36);
            _folderButton.Visible = false;
            _tip.SetToolTip(_folderButton, "Open the reports folder");
            _folderButton.Click += delegate { OpenFolder(ReportFiles.PrimaryFolder()); };
            _header.Controls.Add(_folderButton);
            Controls.Add(_header);
        }

        private void BuildFooter()
        {
            _footer = new FooterBar();
            _footer.Location = new Point(0, 540);
            _footer.Size = new Size(400, 48);
            _footer.Gear.Click += delegate { OpenGearMenu(); };
            _footer.KeepOpenClicked += delegate { KeepOpen(); };
            Controls.Add(_footer);
        }

        // The gear's menu. Settings is greyed out while a scan runs.
        private void BuildMenu()
        {
            _menuSettings = new ToolStripMenuItem("Settings", null, delegate { ShowSettings(); });
            var folder = new ToolStripMenuItem("Open reports folder", null, delegate { OpenFolder(ReportFiles.PrimaryFolder()); });
            var about = new ToolStripMenuItem("About", null, delegate { ShowAbout(); });
            var close = new ToolStripMenuItem("Close", null, delegate { Close(); });
            _menu = new ContextMenuStrip();
            _menu.Items.Add(_menuSettings);
            _menu.Items.Add(folder);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(about);
            _menu.Items.Add(close);
            _menu.Opening += delegate { _menuSettings.Enabled = !Scanning; };
        }

        // Every screen sits between the header (0-59) and the footer (540-587).
        private Panel NewPage()
        {
            var p = new Panel();
            p.Location = new Point(0, 60);
            p.Size = new Size(400, 480);
            p.BackColor = Color.White;
            p.Visible = false;
            Controls.Add(p);
            return p;
        }

        private void BuildNew()
        {
            _new = NewPage();
            Ui.AddLabel(_new, "Where is this PC?", 16, 20, 368, 30, Ui.Semi(20), Ui.Ink);
            Ui.AddLabel(_new, "Asked once. Next time it remembers.", 16, 52, 368, 20, Ui.Body(13), Ui.Soft);
            Ui.AddLabel(_new, "Location", 16, 84, 178, 16, Ui.Body(12), Ui.Soft);
            Ui.AddLabel(_new, "Department (optional)", 206, 84, 178, 16, Ui.Body(12), Ui.Soft);

            _location = new ComboBox();
            _location.DropDownStyle = ComboBoxStyle.DropDown;
            _location.Location = new Point(16, 104);
            _location.Size = new Size(178, 28);
            _location.Font = Ui.Body(14);
            _new.Controls.Add(_location);

            _department = new ComboBox();
            _department.DropDownStyle = ComboBoxStyle.DropDown;
            _department.Location = new Point(206, 104);
            _department.Size = new Size(178, 28);
            _department.Font = Ui.Body(14);
            _new.Controls.Add(_department);

            Ui.AddLabel(_new, "Places on this drive", 16, 146, 368, 18, Ui.Semi(12), Ui.Soft);
            _placeList = new Panel();
            _placeList.Location = new Point(16, 168);
            _placeList.Size = new Size(368, 216);
            _placeList.AutoScroll = true;
            _placeList.BackColor = Color.White;
            _new.Controls.Add(_placeList);
            _placeNone = Ui.AddLabel(_placeList, "No places yet. Type one above.", 8, 12, 340, 20, Ui.Body(12), Ui.Soft);

            _savesAs = Ui.AddLabel(_new, "", 16, 394, 368, 18, Ui.Body(12), Ui.Soft);

            _scan = new RoundButton();
            _scan.Text = "Scan this PC";
            _scan.Primary = true;
            _scan.Location = new Point(16, 428);
            _scan.Size = new Size(368, 36);
            _scan.Font = Ui.Semi(14);
            _scan.Click += delegate { StartFromNew(); };
            _new.Controls.Add(_scan);

            _location.TextChanged += delegate { UpdateNew(); };
            _department.TextChanged += delegate { UpdateNew(); };
        }

        private void BuildKnown()
        {
            _known = NewPage();
            Ui.AddLabel(_known, "Scanned before", 16, 20, 368, 30, Ui.Semi(20), Ui.Ink);

            _knownCard = new Card();
            _knownCard.Location = new Point(16, 64);
            _knownCard.Size = new Size(368, 76);
            _knownCard.Fill = Ui.CardFill;
            _knownCard.Edge = Ui.CardLine;
            _knownCard.Radius = 8;
            _known.Controls.Add(_knownCard);

            _knownPin = new GlyphView();
            _knownPin.Glyph = "pin";
            _knownPin.Tint = Ui.Brand;
            _knownPin.Px = 22;
            _knownPin.Location = new Point(14, 26);
            _knownPin.Size = new Size(22, 24);
            _knownPin.BackColor = Ui.CardFill;
            _knownCard.Controls.Add(_knownPin);
            _knownPlaceLabel = Ui.AddLabel(_knownCard, "", 48, 16, 300, 24, Ui.Semi(16), Ui.Ink);
            _knownFileLabel = Ui.AddLabel(_knownCard, "", 48, 42, 300, 18, Ui.Body(12), Ui.Soft);

            _ring = new CountRing();
            _ring.Location = new Point(168, 160);
            _ring.BackColor = Color.White;
            _ring.Seconds = 3;
            _ring.Total = 3;
            _known.Controls.Add(_ring);

            _startingLabel = Ui.AddLabel(_known, "Starting the scan\u2026", 0, 234, 400, 20, Ui.Body(14), Ui.Ink);
            _startingLabel.TextAlign = ContentAlignment.MiddleCenter;

            _changePlace = new RoundButton();
            _changePlace.Text = "Change place";
            _changePlace.Location = new Point(16, 428);
            _changePlace.Size = new Size(180, 36);
            _changePlace.Font = Ui.Body(14);
            _known.Controls.Add(_changePlace);

            _scanNow = new RoundButton();
            _scanNow.Text = "Scan now";
            _scanNow.Primary = true;
            _scanNow.Location = new Point(204, 428);
            _scanNow.Size = new Size(180, 36);
            _scanNow.Font = Ui.Semi(14);
            _known.Controls.Add(_scanNow);

            _scanNow.Click += delegate
            {
                _timer.Stop();
                StartScan(_knownPlace);
            };
            _changePlace.Click += delegate
            {
                _timer.Stop();
                _mode = Mode.None;
                FillNewChoices();
                SplitAndSet(_knownPlace);
                ShowPage(_new);
            };
        }

        private void BuildScanning()
        {
            _scanPage = NewPage();
            _stepRows = new ActivityRow[4];
            for (int i = 0; i < 4; i++)
            {
                _stepRows[i] = new ActivityRow();
                _stepRows[i].Glyph = StepGlyphs[i];
                _stepRows[i].Title = StepNames[i];
                _stepRows[i].Sub = StepRunSubs[i];
                _stepRows[i].Location = new Point(0, 8 + i * 60);
                _stepRows[i].Size = new Size(400, 60);
                _scanPage.Controls.Add(_stepRows[i]);
            }
        }

        private void BuildDone()
        {
            _done = NewPage();
            _doneRows = new ActivityRow[5];
            for (int i = 0; i < 5; i++)
            {
                _doneRows[i] = new ActivityRow();
                _doneRows[i].ValueMode = true;
                _doneRows[i].SubColor = Ui.Ink;
                _doneRows[i].Glyph = DoneGlyphs[i];
                _doneRows[i].Title = DoneNames[i];
                _doneRows[i].Location = new Point(0, 8 + i * 50);
                _doneRows[i].Size = new Size(400, 50);
                _done.Controls.Add(_doneRows[i]);
            }

            _doneFolderIcon = new GlyphView();
            _doneFolderIcon.Glyph = "folder";
            _doneFolderIcon.Tint = Ui.Soft;
            _doneFolderIcon.Px = 16;
            _doneFolderIcon.Location = new Point(16, 362);
            _doneFolderIcon.Size = new Size(18, 18);
            _done.Controls.Add(_doneFolderIcon);

            _doneFolder = Ui.AddLabel(_done, "", 40, 359, 246, 24, Ui.Body(12), Ui.Soft);
            _doneFolder.AutoEllipsis = true;
            _doneFolder.TextAlign = ContentAlignment.MiddleLeft;

            _openFolder = new LinkLabel();
            _openFolder.Text = "Open folder";
            _openFolder.AutoSize = true;
            _openFolder.Location = new Point(296, 362);
            _openFolder.Font = Ui.Body(12);
            _openFolder.LinkColor = Ui.Brand;
            _openFolder.ActiveLinkColor = Ui.BrandHover;
            _openFolder.VisitedLinkColor = Ui.Brand;
            _openFolder.BackColor = Color.White;
            _openFolder.LinkClicked += delegate
            {
                OpenFolder(string.IsNullOrEmpty(_donePath) ? ReportFiles.PrimaryFolder() : Path.GetDirectoryName(_donePath));
            };
            _done.Controls.Add(_openFolder);

            _doneNote = Ui.AddLabel(_done, "", 16, 384, 368, 34, Ui.Body(12), Ui.Amber);
            _doneNote.Visible = false;

            _close = new RoundButton();
            _close.Text = "Close";
            _close.Primary = true;
            _close.Location = new Point(16, 428);
            _close.Size = new Size(368, 36);
            _close.Font = Ui.Semi(14);
            _close.Click += delegate { Close(); };
            _done.Controls.Add(_close);
        }

        private void BuildError()
        {
            _error = NewPage();
            _errorBox = new Card();
            _errorBox.Location = new Point(16, 18);
            _errorBox.Size = new Size(368, 60);
            _errorBox.Fill = Ui.Pink;
            _errorBox.Edge = Ui.PinkLine;
            _errorBox.Radius = 6;
            _error.Controls.Add(_errorBox);
            _errorText = Ui.AddLabel(_errorBox, "", 14, 12, 340, 36, Ui.Body(14), Ui.Ink);

            _whatToDo = Ui.AddLabel(_error, "What to do", 16, 0, 368, 22, Ui.Semi(14), Ui.Ink);
            string[] steps =
            {
                "Close this window.",
                "On the drive, double-click PMW-Device-Scan.bat. It does the same scan another way.",
                "Still stuck? Free some space, or check the drive's lock switch."
            };
            _stepDots = new GlyphView[3];
            _stepTexts = new Label[3];
            for (int i = 0; i < 3; i++)
            {
                _stepDots[i] = new GlyphView();
                _stepDots[i].Glyph = (i + 1).ToString();
                _stepDots[i].Tint = Ui.Ink;
                _stepDots[i].Fill = Ui.Tile;
                _stepDots[i].Radius = 11;
                _stepDots[i].Px = 12;
                _stepDots[i].Location = new Point(16, 0);
                _stepDots[i].Size = new Size(22, 22);
                _error.Controls.Add(_stepDots[i]);
                _stepTexts[i] = Ui.AddLabel(_error, steps[i], 48, 0, 352, 20, Ui.Body(14), Ui.Ink);
            }

            _techLink = new LinkLabel();
            _techLink.Text = "Technical detail";
            _techLink.AutoSize = true;
            _techLink.Location = new Point(16, 0);
            _techLink.Font = Ui.Body(12);
            _techLink.LinkColor = Ui.Soft;
            _techLink.ActiveLinkColor = Ui.Soft;
            _techLink.VisitedLinkColor = Ui.Soft;
            _techLink.BackColor = Color.White;
            _techLink.LinkClicked += delegate { _techText.Visible = !_techText.Visible; };
            _error.Controls.Add(_techLink);

            _techText = Ui.AddLabel(_error, "", 16, 0, 368, 20, Ui.Mono(), Ui.Soft);
            _techText.Visible = false;

            _errorClose = new RoundButton();
            _errorClose.Text = "Close";
            _errorClose.Location = new Point(16, 428);
            _errorClose.Size = new Size(368, 36);
            _errorClose.Font = Ui.Body(14);
            _errorClose.Click += delegate { Close(); };
            _error.Controls.Add(_errorClose);
        }

        // ------------------------------------------------------------------
        //  Pages
        // ------------------------------------------------------------------

        private void ShowPage(Panel page)
        {
            foreach (Panel p in _pages) p.Visible = (p == page);
            AcceptButton = (page == _new) ? _scan : null;
            _anim.Enabled = (page == _scanPage);
            _folderButton.Visible = (page == _new || page == _known);
            _footer.Spinning = (page == _scanPage);
            _footer.Gear.Visible = (page != _error);
            if (page == _new || page == _known) _header.SetPlain(_computer, _model);
            if (page == _scanPage) _header.SetPlain(_computer, string.IsNullOrEmpty(_knownPlace) ? _model : _knownPlace);
            if (page == _new) _footer.Status = "Ready";
            if (page != _done) _footer.ShowKeep = false;
        }

        private void ShowSettings()
        {
            _settingsFromDone = _done.Visible;
            _timer.Stop();
            _mode = Mode.None;
            using (var dialog = new SettingsForm(false))
            {
                dialog.ShowDialog(this);
            }
            LeaveSettings();
        }

        // Opened from Completed, Settings goes back there; from a start screen it
        // starts over, since a new folder can make a known machine new.
        private void LeaveSettings()
        {
            if (_settingsFromDone)
            {
                _footer.Status = "";
                _footer.ShowKeep = false;
                ShowPage(_done);
            }
            else
            {
                Begin();
            }
        }

        private void ShowAbout()
        {
            MessageBox.Show(this, "PMW Device Scan " + AppInfo.Version + "\r\n\r\nWrites the same report as PMW-Device-Scan.bat.",
                "About PMW Device Scan", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OpenGearMenu()
        {
            Size size = _menu.PreferredSize;
            _menu.Show(_footer.Gear, new Point(_footer.Gear.Width - size.Width, -size.Height - 4));
        }

        private void KeepOpen()
        {
            _timer.Stop();
            _mode = Mode.None;
            _footer.ShowKeep = false;
            _footer.Status = "Stays open until closed";
        }

        // Explorer on the folder; created first if it is missing.
        private void OpenFolder(string dir)
        {
            try
            {
                Directory.CreateDirectory(dir);
                Process.Start("explorer.exe", "\"" + dir + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "The folder could not be opened: " + ex.Message, "PMW Device Scan",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool Scanning
        {
            get { return _worker != null && _worker.IsBusy; }
        }

        // ------------------------------------------------------------------
        //  New machine: choices, places, the live "Saves as" line
        // ------------------------------------------------------------------

        private void FillNewChoices()
        {
            List<string> places = ReportFiles.KnownPlaces();

            var locations = new List<string> { "F1", "F3", "PML" };
            var departments = new List<string>();
            foreach (string place in places)
            {
                string[] parts = place.Split(new[] { ' ' }, 2);
                if (!locations.Any(x => string.Equals(x, parts[0], StringComparison.OrdinalIgnoreCase)))
                    locations.Add(parts[0]);
                if (parts.Length > 1 && parts[1].Length > 0 && !departments.Any(x => string.Equals(x, parts[1], StringComparison.OrdinalIgnoreCase)))
                    departments.Add(parts[1]);
            }
            departments.Sort(StringComparer.OrdinalIgnoreCase);

            _location.Items.Clear();
            _location.Items.AddRange(locations.ToArray());
            _department.Items.Clear();
            _department.Items.AddRange(departments.ToArray());

            BuildPlaceRows(places);
        }

        // How many reports on the drive carry each place, from their file names.
        private static Dictionary<string, int> PlaceCounts()
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in ReportFiles.ReportNames())
            {
                Match m = Regex.Match(name, @"^\[(.+?)\]");
                if (!m.Success) continue;
                string place = m.Groups[1].Value;
                int n;
                counts.TryGetValue(place, out n);
                counts[place] = n + 1;
            }
            return counts;
        }

        private void BuildPlaceRows(List<string> places)
        {
            Dictionary<string, int> counts = PlaceCounts();
            foreach (PlaceRow old in _placeRows)
            {
                _placeList.Controls.Remove(old);
                old.Dispose();
            }
            _placeRows.Clear();

            foreach (string place in places)
            {
                string name = place;
                int n;
                counts.TryGetValue(name, out n);
                var row = new PlaceRow();
                row.Place = name;
                row.Count = n + (n == 1 ? " PC" : " PCs");
                row.Click += delegate { SplitAndSet(name); };
                _placeRows.Add(row);
                _placeList.Controls.Add(row);
            }
            LayoutPlaces();
        }

        private void LayoutPlaces()
        {
            bool scroll = _placeRows.Count * 44 > _placeList.ClientSize.Height;
            int w = _placeList.ClientSize.Width - (scroll ? SystemInformation.VerticalScrollBarWidth : 0);
            for (int i = 0; i < _placeRows.Count; i++)
            {
                _placeRows[i].SetBounds(0, i * 44, w, 44);
            }
            _placeNone.Visible = _placeRows.Count == 0;
        }

        // "F1 ENGINEERING" becomes location F1 and department ENGINEERING.
        private void SplitAndSet(string place)
        {
            string[] parts = (place ?? "").Trim().Split(new[] { ' ' }, 2);
            string loc = parts.Length > 0 ? parts[0] : "";
            string dept = parts.Length > 1 ? parts[1] : "";
            SetPlace(loc, dept);
        }

        private void SetPlace(string loc, string dept)
        {
            _location.Text = loc;
            _department.Text = dept;
            UpdateNew();
        }

        private void UpdateNew()
        {
            string loc = _location.Text.Trim();
            _scan.Enabled = loc.Length > 0;
            string place = ReportFiles.CleanPlace(loc + " " + _department.Text);
            _savesAs.Text = "Saves as " + FolderShown() + "\\" + ReportFiles.FileName(place);
            foreach (PlaceRow row in _placeRows)
            {
                row.Picked = place.Length > 0 && string.Equals(row.Place, place, StringComparison.OrdinalIgnoreCase);
            }
        }

        private void StartFromNew()
        {
            if (_location.Text.Trim().Length == 0) return;
            StartScan(ReportFiles.CleanPlace(_location.Text + " " + _department.Text));
        }

        // ------------------------------------------------------------------
        //  Known machine
        // ------------------------------------------------------------------

        private void ShowKnown(string file)
        {
            _knownPlace = ReportFiles.PlaceOfFile(file);
            FillKnown(_knownPlace, File.GetLastWriteTime(file));
            _mode = Mode.KnownCountdown;
            _countdown = 3;
            _ring.Seconds = 3;
            _footer.Status = "Starts by itself in 3 s";
            ShowPage(_known);
            _timer.Start();
        }

        private void FillKnown(string place, DateTime saved)
        {
            _knownPlaceLabel.Text = place;
            _knownFileLabel.Text = "Last scan " + saved.ToString("d MMM yyyy") + " \u00B7 this one replaces it";
        }

        // ------------------------------------------------------------------
        //  Scanning
        // ------------------------------------------------------------------

        private void SetStep(int step)
        {
            for (int i = 0; i < 4; i++)
            {
                ActivityRow row = _stepRows[i];
                bool done = i < step;
                bool running = i == step;
                row.TitleBold = running;
                row.Progress = running;
                row.Phase = _phase;
                if (done)
                {
                    row.Sub = StepDoneSubs[i];
                    row.SubColor = Ui.Soft;
                    row.State = "Done";
                    row.StateColor = Ui.Green;
                }
                else if (running)
                {
                    row.Sub = StepRunSubs[i];
                    row.SubColor = Ui.Soft;
                    row.State = "Reading";
                    row.StateColor = Ui.Brand;
                }
                else
                {
                    row.Sub = (i == 3) ? "Waiting" : StepRunSubs[i];
                    row.SubColor = (i == 3) ? Ui.Waiting : Ui.Soft;
                    row.State = "";
                }
                row.Invalidate();
            }
            _footer.Status = step < 4
                ? "Scanning \u00B7 " + Math.Min(step + 1, 4) + " of 4 \u00B7 keep the drive plugged in"
                : "";
        }

        private void StartScan(string place)
        {
            _timer.Stop();
            _mode = Mode.None;
            _knownPlace = place;
            _scanStart = DateTime.Now;
            SetStep(0);
            ShowPage(_scanPage);

            var worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                List<string> lines = ReportBuilder.Build(step => worker.ReportProgress(step * 25, step));
                string path = ReportFiles.Save(lines, place);
                e.Result = new ScanOutcome { Path = path, Lines = lines };
            };
            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                SetStep((int)e.UserState);
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                OnScanFinished(e);
            };
            _worker = worker;
            worker.RunWorkerAsync();
        }

        private void OnScanFinished(RunWorkerCompletedEventArgs e)
        {
            _scanSeconds = (int)(DateTime.Now - _scanStart).TotalSeconds;
            if (e.Error != null)
            {
                if (e.Error is SaveFailedException)
                    ShowError("The report could not be saved on this drive or on the desktop. Check that the drive is not read-only or full.",
                        Detail(e.Error));
                else
                    ShowError("Something went wrong during the scan: " + e.Error.Message, Detail(e.Error));
                return;
            }

            var outcome = (ScanOutcome)e.Result;
            SetStep(4);
            FillDone(outcome);
            ShowPage(_done);
            if (ReportFiles.Current.AutoClose)
            {
                _mode = Mode.DoneCountdown;
                _countdown = 10;
                _footer.Status = "Closes by itself in 10 s";
                _footer.ShowKeep = true;
                _timer.Start();
            }
            else
            {
                _footer.Status = "";
                _footer.ShowKeep = false;
            }
        }

        private void FillDone(ScanOutcome o)
        {
            List<string> lines = o.Lines;
            _doneRows[0].Sub = FirstValue(lines, "Serial Number");
            _doneRows[1].Sub = FirstValue(lines, "Product Number");

            List<string> win = BlockOf(lines, "Windows License");
            string wl = "None found";
            if (win.Count > 0 && win[0] != "None found" && !win[0].StartsWith("Firmware key"))
            {
                string[] parts = win[0].Split('|');
                wl = WindowsChannelWords(parts[0].Trim()) + " \u00B7 " + (parts.Length > 1 ? parts[1].Trim() : "");
            }
            _doneRows[2].Sub = wl;

            string office = OfficeWords(BlockOf(lines, "Office License"));
            _doneRows[3].Sub = office;
            _doneRows[3].SubColor = office.IndexOf("personal copy", StringComparison.OrdinalIgnoreCase) >= 0 ? Ui.Amber : Ui.Ink;

            _doneRows[4].Sub = BlockOf(lines, "Installed Software").Count + " programs";
            foreach (ActivityRow row in _doneRows) row.Invalidate();

            _donePath = o.Path;
            _doneFolder.Text = ReportFiles.OnDrive(o.Path);
            _tip.SetToolTip(_doneFolder, o.Path);
            bool fell = !ReportFiles.InPrimaryFolder(o.Path);
            _doneNote.Visible = fell;
            if (fell)
                _doneNote.Text = "Could not reach " + FolderShown() + ", so the report went to the folder above.";

            _header.SetResult("Completed", Ui.Green, Dot(_computer, _knownPlace, _scanSeconds + " s"), true);
        }

        // The same words the device list uses: OEM, Retail, Volume (MAK), Volume (KMS).
        private static string WindowsChannelWords(string channel)
        {
            string c = channel.ToUpperInvariant().Replace(':', ' ').Replace('_', ' ');
            if (c.StartsWith("OEM")) return "OEM";
            if (c.StartsWith("RETAIL")) return "Retail";
            if (c.StartsWith("VOLUME") && c.Contains("MAK")) return "Volume (MAK)";
            if (c.StartsWith("VOLUME")) return "Volume (KMS)";
            return channel;
        }

        // The products installed, by name, with a personal copy called out.
        private static string OfficeWords(List<string> block)
        {
            bool company = false, personal = false, perpetual = false;
            foreach (string line in block)
            {
                string id = line.Split('|')[0].Trim();
                if (id.StartsWith("Signed-in licence") || id.StartsWith("Office ")) continue;
                if (System.Text.RegularExpressions.Regex.IsMatch(id, "HomePrem|HomeBusiness|HomeStudent|Personal", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) personal = true;
                else if (id.StartsWith("O365") || id.StartsWith("M365")) company = true;
                else if (id.Length > 0 && id != "None found") perpetual = true;
            }
            string words = company ? "Microsoft 365" : perpetual ? "Office, bought once" : personal ? "Personal Microsoft 365" : "None";
            if (personal && (company || perpetual)) words += " + a personal copy";
            return words;
        }

        private static string FirstValue(List<string> lines, string label)
        {
            List<string> block = BlockOf(lines, label);
            return block.Count > 0 ? block[0] : "None";
        }

        // The values under "Label:" up to the blank line that ends the block.
        private static List<string> BlockOf(List<string> lines, string label)
        {
            var r = new List<string>();
            int i = lines.IndexOf(label + ":");
            if (i < 0) return r;
            for (int j = i + 1; j < lines.Count && lines[j].Length > 0; j++) r.Add(lines[j]);
            return r;
        }

        private void ShowError(string message, string detail)
        {
            _timer.Stop();
            _mode = Mode.None;
            _errorText.Text = message;
            _techText.Text = detail;
            _techText.Visible = false;
            LayoutError();
            _header.SetResult("Could not finish", Ui.Red, Dot(_computer, _knownPlace), false);
            _footer.Status = "Stays open until closed";
            _footer.ShowKeep = false;
            ShowPage(_error);
        }

        // Stacks the error screen from its text: the message box, the three steps, then the technical detail.
        private void LayoutError()
        {
            int messageH = Ui.Lines(_errorText.Text, _errorText.Font, 340);
            _errorText.Size = new Size(340, messageH);
            _errorBox.Size = new Size(368, messageH + 24);

            int y = _errorBox.Bottom + 16;
            _whatToDo.Location = new Point(16, y);
            y += 30;
            for (int i = 0; i < 3; i++)
            {
                int h = Math.Max(22, Ui.Lines(_stepTexts[i].Text, _stepTexts[i].Font, 352));
                _stepDots[i].Location = new Point(16, y);
                _stepTexts[i].SetBounds(48, y, 352, h);
                y += h + 10;
            }
            _techLink.Location = new Point(16, y + 4);
            y += 30;
            _techText.SetBounds(16, y, 368, Math.Max(18, Ui.Lines(_techText.Text, _techText.Font, 368)));
        }

        private static string Detail(Exception ex)
        {
            return ex.GetType().Name + ": " + ex.Message;
        }

        // ------------------------------------------------------------------
        //  Timers and helpers
        // ------------------------------------------------------------------

        private void OnTick(object sender, EventArgs e)
        {
            if (_mode == Mode.KnownCountdown)
            {
                _countdown--;
                if (_countdown <= 0)
                {
                    _timer.Stop();
                    _mode = Mode.None;
                    StartScan(_knownPlace);
                }
                else
                {
                    _ring.Seconds = _countdown;
                    _footer.Status = "Starts by itself in " + _countdown + " s";
                }
            }
            else if (_mode == Mode.DoneCountdown)
            {
                _countdown--;
                if (_countdown <= 0)
                {
                    _timer.Stop();
                    Close();
                }
                else
                {
                    _footer.Status = "Closes by itself in " + _countdown + " s";
                }
            }
        }

        // Moves the marquee bars and the footer spinner while a scan runs.
        private void OnAnimTick(object sender, EventArgs e)
        {
            _phase += 2;
            foreach (ActivityRow row in _stepRows)
            {
                row.Phase = _phase;
                row.Invalidate();
            }
            _footer.Angle = (_footer.Angle + 12) % 360;
            _footer.Invalidate();
        }

        private static string Dot(params string[] parts)
        {
            var kept = new List<string>();
            foreach (string p in parts)
                if (!string.IsNullOrEmpty(p)) kept.Add(p);
            return string.Join(" \u00B7 ", kept.ToArray());
        }

        // Model of this PC, read from the machine itself.
        private static string ModelName()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("root\\cimv2", "SELECT Model FROM Win32_ComputerSystem"))
                using (ManagementObjectCollection rows = searcher.Get())
                {
                    foreach (ManagementObject o in rows)
                    {
                        object v = o["Model"];
                        return v == null ? "" : Convert.ToString(v).Trim();
                    }
                }
            }
            catch (Exception) { }
            return "";
        }

        // ------------------------------------------------------------------
        //  Demo: one screen with sample values, for checking the layout only
        // ------------------------------------------------------------------

        private void ShowDemo(string page)
        {
            _knownPlace = "F1 ENGINEERING";
            switch (page)
            {
                case "known":
                    FillKnown(_knownPlace, new DateTime(2026, 10, 12));
                    _ring.Seconds = 3;
                    _footer.Status = "Starts by itself in 3 s";
                    ShowPage(_known);
                    break;
                case "scanning":
                    _scanStart = DateTime.Now;
                    SetStep(2);
                    ShowPage(_scanPage);
                    break;
                case "done":
                    _scanSeconds = 58;
                    SetStep(4);
                    FillDone(SampleOutcome());
                    ShowPage(_done);
                    _footer.Status = "Closes by itself in 7 s";
                    _footer.ShowKeep = true;
                    break;
                case "error":
                    _knownPlace = "F1 IT";
                    ShowError("The report could not be saved. The drive is full or write-protected, and the desktop would not take it either.",
                        "UnauthorizedAccessException: Access to the path 'E:\\Reports\\[F1 IT] " + ReportFiles.Computer() + "_.txt' is denied.");
                    break;
                default:
                    FillNewChoices();
                    SetPlace("F1", "ENGINEERING");
                    ShowPage(_new);
                    break;
            }
        }

        private ScanOutcome SampleOutcome()
        {
            var lines = new List<string>
            {
                "Serial Number:", "5CD333KBWH", "",
                "Product Number:", "7Z8D7PA#UUF", "",
                "Windows License:", "Volume:KMS|Licensed", "",
                "Office License:", "O365ProPlusRetail|Licensed", "HomePremium2019|Licensed", "",
                "Installed Software:"
            };
            for (int i = 1; i <= 56; i++) lines.Add("Program " + i + "|1.0");
            return new ScanOutcome
            {
                Path = Path.Combine(ReportFiles.PrimaryFolder(), "[F1 ENGINEERING] " + ReportFiles.Computer() + "_.txt"),
                Lines = lines
            };
        }
    }
}
