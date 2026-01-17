using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KurosHandleCloser
{
    public class Settings
    {
        public string ProcessName { get; set; } = "";
        public string HandleName { get; set; } = "";
        public string ProfileName { get; set; } = "";
    }

    public class SettingsCollection
    {
        public Dictionary<string, Settings> Profiles { get; set; } = new Dictionary<string, Settings>();
    }

    public partial class MainForm : Form
    {
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CTLCOLOREDIT)
            {
                if (editBrush != IntPtr.Zero)
                {
                    m.Result = editBrush;
                    return;
                }
            }
            base.WndProc(ref m);
        }
        private bool isMonitoring = false;
        private CancellationTokenSource cancellationTokenSource;
        private Label creditLabel;
        private Panel processPanel;
        private Panel handlePanel;
        private Label processLabel;
        private Label handleLabel;
        private ComboBox processComboBox;
        private ComboBox handleComboBox;
        private Button startStopButton;
        private Button saveButton;
        private Button loadButton;
        private Button refreshButton;
        private Button handleRefreshButton;
        private Panel statusPanel;
        private Label statusLabel;
        private string selectedProcessName = "";
        private string selectedHandleName = "";
        private List<string> allProcesses = new List<string>();
        private List<string> allHandles = new List<string>();
        private bool isUpdatingProcessCombo = false;
        private bool isUpdatingHandleCombo = false;
        private const string SettingsFile = "HandleCloser_Settings.json";
        private const uint PROCESS_DUP_HANDLE = 0x0040;
        private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        private const uint STATUS_INFO_LENGTH_MISMATCH = 0xC0000004;
        private const uint DUPLICATE_CLOSE_SOURCE = 0x00000001;
        private const int SystemHandleInformation = 16;
        private const int ObjectNameInformation = 1;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_CAPTION_COLOR = 35;

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct SYSTEM_HANDLE_TABLE_ENTRY_INFO
        {
            public ushort ProcessId;
            public ushort CreatorBackTraceIndex;
            public byte ObjectTypeNumber;
            public byte HandleAttributes;
            public ushort HandleValue;
            public uint Reserved;
            public IntPtr Object;
            public uint GrantedAccess;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct UNICODE_STRING
        {
            public ushort Length;
            public ushort MaximumLength;
            public IntPtr Buffer;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct OBJECT_NAME_INFORMATION
        {
            public UNICODE_STRING Name;
        }

        [DllImport("ntdll.dll")]
        private static extern uint NtQuerySystemInformation(
            int SystemInformationClass,
            IntPtr SystemInformation,
            int SystemInformationLength,
            out int ReturnLength);

        [DllImport("ntdll.dll")]
        private static extern uint NtDuplicateObject(
            IntPtr SourceProcessHandle,
            IntPtr SourceHandle,
            IntPtr TargetProcessHandle,
            out IntPtr TargetHandle,
            uint DesiredAccess,
            uint Attributes,
            uint Options);

        [DllImport("ntdll.dll")]
        private static extern uint NtQueryObject(
            IntPtr Handle,
            int ObjectInformationClass,
            IntPtr ObjectInformation,
            int ObjectInformationLength,
            out int ReturnLength);

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateSolidBrush(uint crColor);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private const uint EM_SETSEL = 0x00B1;
        private const uint WM_KILLFOCUS = 0x0008;
        private const uint WM_CTLCOLOREDIT = 0x0133;
        private const int COLOR_WINDOW = 5;
        private IntPtr editBrush = IntPtr.Zero;

        [DllImport("advapi32.dll")]
        private static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);

        [DllImport("advapi32.dll")]
        private static extern bool LookupPrivilegeValue(string lpSystemName, string lpName, out LUID lpLuid);

        [DllImport("advapi32.dll")]
        private static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges, ref TOKEN_PRIVILEGES NewState, int BufferLength, IntPtr PreviousState, IntPtr ReturnLength);

        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public uint LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_PRIVILEGES
        {
            public int PrivilegeCount;
            public LUID Luid;
            public uint Attributes;
        }

        private const uint TOKEN_ADJUST_PRIVILEGES = 0x0020;
        private const uint TOKEN_QUERY = 0x0008;
        private const uint SE_PRIVILEGE_ENABLED = 0x00000002;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Kuro's Handle Closer";
            this.Size = new Size(520, 280);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("KurosHandleCloser.HandleCloser.ico"))
                {
                    if (stream != null)
                    {
                        this.Icon = new System.Drawing.Icon(stream);
                    }
                    else if (File.Exists("HandleCloser.ico"))
                    {
                        this.Icon = new System.Drawing.Icon("HandleCloser.ico");
                    }
                    else
                    {
                        this.Icon = CreateAppIcon();
                    }
                }
            }
            catch
            {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("KurosHandleCloser.HandleCloser.ico"))
                {
                    if (stream != null)
                    {
                        this.Icon = new System.Drawing.Icon(stream);
                    }
                    else if (File.Exists("HandleCloser.ico"))
                    {
                        this.Icon = new System.Drawing.Icon("HandleCloser.ico");
                    }
                    else
                    {
                        this.Icon = CreateAppIcon();
                    }
                }
            }
            catch
            {
                try
                {
                    if (File.Exists("HandleCloser.ico"))
                    {
                        this.Icon = new System.Drawing.Icon("HandleCloser.ico");
                    }
                    else
                    {
                        this.Icon = CreateAppIcon();
                    }
                }
                catch
                {
                    this.Icon = CreateAppIcon();
                }
            }
            }
            
            uint darkColor = (uint)((40 << 16) | (35 << 8) | 35);
            editBrush = CreateSolidBrush(darkColor);
            
            this.BackColor = Color.FromArgb(20, 20, 25);
            this.Paint += MainForm_Paint;
            this.Load += MainForm_Load;
            this.Click += MainForm_Click;

            creditLabel = new Label
            {
                Text = "Made by KuroSuika",
                Location = new Point(15, 25),
                Size = new Size(200, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Regular),
                ForeColor = Color.FromArgb(160, 160, 160),
                BackColor = Color.Transparent,
                AutoSize = true
            };
            this.Controls.Add(creditLabel);

            processPanel = new Panel
            {
                Location = new Point(15, 50),
                Size = new Size(480, 40),
                BackColor = Color.FromArgb(45, 45, 50),
                BorderStyle = BorderStyle.None,
                Padding = new Padding(12, 8, 12, 8)
            };
            processPanel.Paint += Panel_Paint;
            this.Controls.Add(processPanel);

            processLabel = new Label
            {
                Text = "Process:",
                Location = new Point(12, 10),
                Size = new Size(60, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                ForeColor = Color.FromArgb(220, 220, 220),
                BackColor = Color.Transparent,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            processPanel.Controls.Add(processLabel);

            processComboBox = new ComboBox
            {
                Location = new Point(75, 8),
                Size = new Size(370, 24),
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(35, 35, 40),
                ForeColor = Color.FromArgb(220, 220, 220),
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                MaxDropDownItems = 10
            };
            processComboBox.SelectedIndexChanged += ProcessComboBox_SelectedIndexChanged;
            processComboBox.DropDown += ProcessComboBox_DropDown;
            processComboBox.DropDownClosed += ProcessComboBox_DropDownClosed;
            processComboBox.Leave += ProcessComboBox_Leave;
            processComboBox.DrawMode = DrawMode.OwnerDrawFixed;
            processComboBox.DrawItem += ComboBox_DrawItem;
            processPanel.Controls.Add(processComboBox);

            refreshButton = new Button
            {
                Text = "↻",
                Location = new Point(450, 8),
                Size = new Size(25, 24),
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Padding = new Padding(0)
            };
            refreshButton.FlatAppearance.BorderSize = 0;
            refreshButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            refreshButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            refreshButton.Paint += RefreshButton_Paint;
            refreshButton.MouseEnter += (s, e) => refreshButton.Invalidate();
            refreshButton.MouseLeave += (s, e) => refreshButton.Invalidate();
            refreshButton.MouseDown += (s, e) => refreshButton.Invalidate();
            refreshButton.MouseUp += (s, e) => refreshButton.Invalidate();
            refreshButton.Click += RefreshButton_Click;
            processPanel.Controls.Add(refreshButton);

            handlePanel = new Panel
            {
                Location = new Point(15, 95),
                Size = new Size(480, 40),
                BackColor = Color.FromArgb(45, 45, 50),
                BorderStyle = BorderStyle.None,
                Padding = new Padding(12, 8, 12, 8)
            };
            handlePanel.Paint += Panel_Paint;
            this.Controls.Add(handlePanel);

            handleLabel = new Label
            {
                Text = "Handle:",
                Location = new Point(12, 10),
                Size = new Size(60, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                ForeColor = Color.FromArgb(220, 220, 220),
                BackColor = Color.Transparent,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            handlePanel.Controls.Add(handleLabel);

            handleComboBox = new ComboBox
            {
                Location = new Point(75, 8),
                Size = new Size(370, 24),
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(35, 35, 40),
                ForeColor = Color.FromArgb(220, 220, 220),
                FlatStyle = FlatStyle.Flat,
                Enabled = false,
                TabStop = false,
                MaxDropDownItems = 10
            };
            handleComboBox.DropDown += HandleComboBox_DropDown;
            handleComboBox.DropDownClosed += HandleComboBox_DropDownClosed;
            handleComboBox.SelectedIndexChanged += HandleComboBox_SelectedIndexChanged;
            handleComboBox.Leave += HandleComboBox_Leave;
            handleComboBox.DrawMode = DrawMode.OwnerDrawFixed;
            handleComboBox.DrawItem += ComboBox_DrawItem;
            handlePanel.Controls.Add(handleComboBox);

            handleRefreshButton = new Button
            {
                Text = "↻",
                Location = new Point(450, 8),
                Size = new Size(25, 24),
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Padding = new Padding(0),
                Enabled = false
            };
            handleRefreshButton.FlatAppearance.BorderSize = 0;
            handleRefreshButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            handleRefreshButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            handleRefreshButton.Paint += RefreshButton_Paint;
            handleRefreshButton.MouseEnter += (s, e) => handleRefreshButton.Invalidate();
            handleRefreshButton.MouseLeave += (s, e) => handleRefreshButton.Invalidate();
            handleRefreshButton.MouseDown += (s, e) => handleRefreshButton.Invalidate();
            handleRefreshButton.MouseUp += (s, e) => handleRefreshButton.Invalidate();
            handleRefreshButton.Click += HandleRefreshButton_Click;
            handlePanel.Controls.Add(handleRefreshButton);

            statusPanel = new Panel
            {
                Location = new Point(15, 140),
                Size = new Size(480, 35),
                BackColor = Color.FromArgb(45, 45, 50),
                BorderStyle = BorderStyle.None,
                Padding = new Padding(12, 8, 12, 8)
            };
            statusPanel.Paint += Panel_Paint;
            this.Controls.Add(statusPanel);

            statusLabel = new Label
            {
                Text = "Status: Ready",
                Location = new Point(12, 8),
                Size = new Size(456, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                ForeColor = Color.FromArgb(220, 220, 220),
                BackColor = Color.Transparent,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusPanel.Controls.Add(statusLabel);

            startStopButton = new Button
            {
                Text = "Start",
                Location = new Point(45, 180),
                Size = new Size(140, 40),
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            startStopButton.FlatAppearance.BorderSize = 0;
            startStopButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            startStopButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            startStopButton.Padding = new Padding(0);
            startStopButton.Margin = new Padding(0);
            startStopButton.Paint += StartStopButton_Paint;
            startStopButton.MouseEnter += (s, e) => startStopButton.Invalidate();
            startStopButton.MouseLeave += (s, e) => startStopButton.Invalidate();
            startStopButton.MouseDown += (s, e) => startStopButton.Invalidate();
            startStopButton.MouseUp += (s, e) => startStopButton.Invalidate();
            startStopButton.Click += StartStopButton_Click;
            this.Controls.Add(startStopButton);

            saveButton = new Button
            {
                Text = "Save Settings",
                Location = new Point(195, 180),
                Size = new Size(130, 40),
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            saveButton.FlatAppearance.BorderSize = 0;
            saveButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            saveButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            saveButton.Paint += SaveLoadButton_Paint;
            saveButton.MouseEnter += (s, e) => saveButton.Invalidate();
            saveButton.MouseLeave += (s, e) => saveButton.Invalidate();
            saveButton.MouseDown += (s, e) => saveButton.Invalidate();
            saveButton.MouseUp += (s, e) => saveButton.Invalidate();
            saveButton.Click += SaveButton_Click;
            this.Controls.Add(saveButton);

            loadButton = new Button
            {
                Text = "Load Settings",
                Location = new Point(335, 180),
                Size = new Size(120, 40),
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            loadButton.FlatAppearance.BorderSize = 0;
            loadButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            loadButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            loadButton.Paint += SaveLoadButton_Paint;
            loadButton.MouseEnter += (s, e) => loadButton.Invalidate();
            loadButton.MouseLeave += (s, e) => loadButton.Invalidate();
            loadButton.MouseDown += (s, e) => loadButton.Invalidate();
            loadButton.MouseUp += (s, e) => loadButton.Invalidate();
            loadButton.Click += LoadButton_Click;
            this.Controls.Add(loadButton);

            this.FormClosing += MainForm_FormClosing;
        }

        private async void MainForm_Load(object sender, EventArgs e)
        {
            int darkMode = 1;
            DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, sizeof(int));
            int blackColor = 0x000000;
            DwmSetWindowAttribute(this.Handle, DWMWA_CAPTION_COLOR, ref blackColor, sizeof(int));
            
            await RefreshProcessListAsync();
        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                e.DrawBackground();
                return;
            }

            ComboBox comboBox = sender as ComboBox;
            if (comboBox == null) return;

            e.DrawBackground();
            
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool isFocused = (e.State & DrawItemState.Focus) == DrawItemState.Focus;
            
            Color backColor = isSelected && isFocused
                ? Color.FromArgb(60, 60, 70)
                : Color.FromArgb(35, 35, 40);
            
            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            string text = comboBox.Items[e.Index].ToString();
            Color textColor = Color.FromArgb(220, 220, 220);
            using (SolidBrush brush = new SolidBrush(textColor))
            {
                Rectangle textRect = e.Bounds;
                textRect.X += 2;
                textRect.Width -= 4;
                e.Graphics.DrawString(text, e.Font, brush, textRect);
            }

            if (isSelected && isFocused)
            {
                e.DrawFocusRectangle();
            }
        }

        private void MainForm_Paint(object sender, PaintEventArgs e)
        {
            using (LinearGradientBrush brush = new LinearGradientBrush(
                new Point(0, 0),
                new Point(0, this.Height),
                Color.FromArgb(25, 25, 30),
                Color.FromArgb(15, 15, 20)))
            {
                e.Graphics.FillRectangle(brush, this.ClientRectangle);
            }
            
            using (SolidBrush dotBrush = new SolidBrush(Color.FromArgb(10, 255, 255, 255)))
            {
                Random rnd = new Random(12345);
                for (int i = 0; i < 50; i++)
                {
                    int x = rnd.Next(0, this.Width);
                    int y = rnd.Next(0, this.Height);
                    e.Graphics.FillEllipse(dotBrush, x, y, 2, 2);
                }
            }
        }

        private void Panel_Paint(object sender, PaintEventArgs e)
        {
            Panel panel = sender as Panel;
            if (panel == null) return;

            GraphicsPath path = new GraphicsPath();
            int radius = 6;
            path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
            path.AddArc(panel.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
            path.AddArc(panel.Width - radius * 2, panel.Height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(0, panel.Height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();

            panel.Region = new Region(path);

            using (LinearGradientBrush gradientBrush = new LinearGradientBrush(
                new Point(0, 0),
                new Point(0, panel.Height),
                Color.FromArgb(50, 50, 55),
                Color.FromArgb(40, 40, 45)))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.FillPath(gradientBrush, path);
            }

            using (Pen borderPen = new Pen(Color.FromArgb(80, 255, 255, 255), 1f))
            {
                e.Graphics.DrawPath(borderPen, path);
            }
        }

        private void StartStopButton_Paint(object sender, PaintEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            GraphicsPath path = new GraphicsPath();
            int radius = 6;
            path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
            path.AddArc(btn.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
            path.AddArc(btn.Width - radius * 2, btn.Height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(0, btn.Height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();

            btn.Region = new Region(path);

            Color topColor, bottomColor, borderColor;
            bool isStopped = btn.Text == "Start";
            
            if (!btn.Enabled)
            {
                topColor = Color.FromArgb(60, 60, 70);
                bottomColor = Color.FromArgb(50, 50, 60);
                borderColor = Color.FromArgb(40, 40, 50);
            }
            else if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)) && Control.MouseButtons == MouseButtons.Left)
            {
                if (isStopped)
                {
                    topColor = Color.FromArgb(25, 120, 45);
                    bottomColor = Color.FromArgb(20, 100, 35);
                    borderColor = Color.FromArgb(15, 90, 25);
                }
                else
                {
                    topColor = Color.FromArgb(160, 25, 40);
                    bottomColor = Color.FromArgb(140, 15, 30);
                    borderColor = Color.FromArgb(120, 10, 25);
                }
            }
            else if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)))
            {
                if (isStopped)
                {
                    topColor = Color.FromArgb(70, 70, 80);
                    bottomColor = Color.FromArgb(60, 60, 70);
                    borderColor = Color.FromArgb(50, 50, 60);
                }
                else
                {
                    topColor = Color.FromArgb(215, 55, 70);
                    bottomColor = Color.FromArgb(165, 30, 45);
                    borderColor = Color.FromArgb(180, 25, 40);
                }
            }
            else
            {
                if (isStopped)
                {
                    topColor = Color.FromArgb(80, 80, 90);
                    bottomColor = Color.FromArgb(60, 60, 70);
                    borderColor = Color.FromArgb(50, 50, 60);
                }
                else
                {
                    topColor = Color.FromArgb(235, 70, 85);
                    bottomColor = Color.FromArgb(185, 35, 50);
                    borderColor = Color.FromArgb(200, 35, 51);
                }
            }

            using (LinearGradientBrush gradientBrush = new LinearGradientBrush(
                new Point(0, 0),
                new Point(0, btn.Height),
                topColor,
                bottomColor))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.FillPath(gradientBrush, path);
            }

            using (Pen borderPen = new Pen(borderColor, 1.5f))
            {
                e.Graphics.DrawPath(borderPen, path);
            }

            TextRenderer.DrawText(e.Graphics, btn.Text, btn.Font, btn.ClientRectangle, btn.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
        }

        private void RefreshButton_Paint(object sender, PaintEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            GraphicsPath path = new GraphicsPath();
            int radius = 6;
            path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
            path.AddArc(btn.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
            path.AddArc(btn.Width - radius * 2, btn.Height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(0, btn.Height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();

            btn.Region = new Region(path);

            Color topColor, bottomColor, borderColor;
            if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)) && Control.MouseButtons == MouseButtons.Left)
            {
                topColor = Color.FromArgb(50, 50, 60);
                bottomColor = Color.FromArgb(40, 40, 50);
                borderColor = Color.FromArgb(30, 30, 40);
            }
            else if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)))
            {
                topColor = Color.FromArgb(70, 70, 80);
                bottomColor = Color.FromArgb(60, 60, 70);
                borderColor = Color.FromArgb(50, 50, 60);
            }
            else
            {
                topColor = Color.FromArgb(80, 80, 90);
                bottomColor = Color.FromArgb(60, 60, 70);
                borderColor = Color.FromArgb(50, 50, 60);
            }

            using (LinearGradientBrush gradientBrush = new LinearGradientBrush(
                new Point(0, 0),
                new Point(0, btn.Height),
                topColor,
                bottomColor))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.FillPath(gradientBrush, path);
            }

            using (Pen borderPen = new Pen(borderColor, 1.5f))
            {
                e.Graphics.DrawPath(borderPen, path);
            }

            TextRenderer.DrawText(e.Graphics, btn.Text, btn.Font, btn.ClientRectangle, btn.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
        }

        private void SaveLoadButton_Paint(object sender, PaintEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            GraphicsPath path = new GraphicsPath();
            int radius = 6;
            path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
            path.AddArc(btn.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
            path.AddArc(btn.Width - radius * 2, btn.Height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(0, btn.Height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();

            btn.Region = new Region(path);

            Color topColor, bottomColor, borderColor;
            if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)) && Control.MouseButtons == MouseButtons.Left)
            {
                topColor = Color.FromArgb(50, 50, 60);
                bottomColor = Color.FromArgb(40, 40, 50);
                borderColor = Color.FromArgb(30, 30, 40);
            }
            else if (btn.ClientRectangle.Contains(btn.PointToClient(Control.MousePosition)))
            {
                topColor = Color.FromArgb(70, 70, 80);
                bottomColor = Color.FromArgb(60, 60, 70);
                borderColor = Color.FromArgb(50, 50, 60);
            }
            else
            {
                topColor = Color.FromArgb(80, 80, 90);
                bottomColor = Color.FromArgb(60, 60, 70);
                borderColor = Color.FromArgb(50, 50, 60);
            }

            using (LinearGradientBrush gradientBrush = new LinearGradientBrush(
                new Point(0, 0),
                new Point(0, btn.Height),
                topColor,
                bottomColor))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.FillPath(gradientBrush, path);
            }

            using (Pen borderPen = new Pen(borderColor, 1.5f))
            {
                e.Graphics.DrawPath(borderPen, path);
            }

            TextRenderer.DrawText(e.Graphics, btn.Text, btn.Font, btn.ClientRectangle, btn.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
        }

        private System.Drawing.Icon CreateAppIcon()
        {
            try
            {
                using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.Clear(Color.Transparent);
                        
                        using (LinearGradientBrush circleBrush = new LinearGradientBrush(
                            new Point(0, 0),
                            new Point(32, 32),
                            Color.FromArgb(50, 50, 55),
                            Color.FromArgb(35, 35, 40)))
                        {
                            g.FillEllipse(circleBrush, 0, 0, 32, 32);
                        }
                        
                        using (Pen borderPen = new Pen(Color.FromArgb(70, 70, 75), 1.5f))
                        {
                            g.DrawEllipse(borderPen, 0, 0, 32, 32);
                        }
                        
                        using (Font font = new Font("MS Gothic", 18, FontStyle.Bold))
                        {
                            using (SolidBrush textBrush = new SolidBrush(Color.White))
                            {
                                StringFormat sf = new StringFormat
                                {
                                    Alignment = StringAlignment.Center,
                                    LineAlignment = StringAlignment.Center
                                };
                                g.DrawString("黒", font, textBrush, new RectangleF(0, 0, 32, 32), sf);
                            }
                        }
                    }
                    
                    IntPtr hIcon = bmp.GetHicon();
                    System.Drawing.Icon icon = System.Drawing.Icon.FromHandle(hIcon);
                    System.Drawing.Icon newIcon = new System.Drawing.Icon(icon, 32, 32);
                    icon.Dispose();
                    return newIcon;
                }
            }
            catch
            {
                return new System.Drawing.Icon(System.Drawing.SystemIcons.Application, 32, 32);
            }
        }

        private void MainForm_Click(object sender, EventArgs e)
        {
            this.ActiveControl = null;
            if (processComboBox.Focused || handleComboBox.Focused)
            {
                processComboBox.SelectedIndex = -1;
                handleComboBox.SelectedIndex = -1;
                ClearComboBoxSelection(processComboBox);
                ClearComboBoxSelection(handleComboBox);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopMonitoring();
            if (editBrush != IntPtr.Zero)
            {
                DeleteObject(editBrush);
                editBrush = IntPtr.Zero;
            }
        }

        private void ProcessComboBox_DropDown(object sender, EventArgs e)
        {
            if (processComboBox.Items.Count == 0 && allProcesses.Count > 0)
            {
                isUpdatingProcessCombo = true;
                processComboBox.Items.AddRange(allProcesses.ToArray());
                processComboBox.MaxDropDownItems = Math.Min(10, allProcesses.Count);
                isUpdatingProcessCombo = false;
            }
        }

        private void ProcessComboBox_DropDownClosed(object sender, EventArgs e)
        {
            ClearComboBoxSelection(processComboBox);
        }

        private void HandleComboBox_DropDown(object sender, EventArgs e)
        {
            if (handleComboBox.Items.Count == 0 && allHandles.Count > 0)
            {
                isUpdatingHandleCombo = true;
                handleComboBox.Items.AddRange(allHandles.ToArray());
                handleComboBox.MaxDropDownItems = Math.Min(10, allHandles.Count);
                isUpdatingHandleCombo = false;
            }
        }

        private void HandleComboBox_DropDownClosed(object sender, EventArgs e)
        {
            ClearComboBoxSelection(handleComboBox);
        }

        private void HandleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingHandleCombo) return;
            
            if (handleComboBox.SelectedIndex >= 0 && handleComboBox.SelectedIndex < handleComboBox.Items.Count)
            {
                selectedHandleName = handleComboBox.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(selectedProcessName) && !string.IsNullOrEmpty(selectedHandleName))
                {
                    startStopButton.Enabled = true;
                }
            }
            else
            {
                selectedHandleName = "";
                startStopButton.Enabled = false;
            }
        }

        private void ProcessComboBox_Enter(object sender, EventArgs e)
        {
            processComboBox.BackColor = Color.FromArgb(35, 35, 40);
            processComboBox.Invalidate();
            
            if (!processComboBox.DroppedDown)
            {
                processComboBox.DroppedDown = true;
            }
            
            Task.Delay(10).ContinueWith(t =>
            {
                if (this.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        ClearComboBoxSelection(processComboBox);
                    });
                }
            });
        }

        private void ProcessComboBox_Leave(object sender, EventArgs e)
        {
            processComboBox.BackColor = Color.FromArgb(35, 35, 40);
            ClearComboBoxSelection(processComboBox);
            processComboBox.Invalidate();
        }

        private void ClearComboBoxSelection(ComboBox comboBox)
        {
            try
            {
                IntPtr editHandle = FindWindowEx(comboBox.Handle, IntPtr.Zero, "Edit", null);
                if (editHandle != IntPtr.Zero)
                {
                    int textLength = comboBox.Text?.Length ?? 0;
                    SendMessage(editHandle, EM_SETSEL, textLength, textLength);
                    SendMessage(editHandle, EM_SETSEL, -1, -1);
                }
            }
            catch { }
        }

        private void ProcessComboBox_MouseClick(object sender, MouseEventArgs e)
        {
            Rectangle dropDownButtonRect = new Rectangle(
                processComboBox.Width - 20, 0, 20, processComboBox.Height);
            
            Rectangle textRect = new Rectangle(0, 0, processComboBox.Width - 20, processComboBox.Height);
            
            if (textRect.Contains(e.Location))
            {
                processComboBox.Focus();
                isUpdatingProcessCombo = true;
                processComboBox.Text = "";
                processComboBox.SelectedIndex = -1;
                isUpdatingProcessCombo = false;
                if (!processComboBox.DroppedDown)
                {
                    processComboBox.DroppedDown = true;
                }
                Task.Delay(20).ContinueWith(t =>
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            ClearComboBoxSelection(processComboBox);
                        });
                    }
                });
            }
            else if (dropDownButtonRect.Contains(e.Location) && !processComboBox.DroppedDown)
            {
                processComboBox.Focus();
                isUpdatingProcessCombo = true;
                processComboBox.Text = "";
                processComboBox.SelectedIndex = -1;
                isUpdatingProcessCombo = false;
                processComboBox.DroppedDown = true;
                Task.Delay(20).ContinueWith(t =>
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            ClearComboBoxSelection(processComboBox);
                        });
                    }
                });
            }
        }

        private void HandleComboBox_Leave(object sender, EventArgs e)
        {
            handleComboBox.BackColor = Color.FromArgb(35, 35, 40);
            ClearComboBoxSelection(handleComboBox);
            handleComboBox.Invalidate();
        }

        private void ProcessComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingProcessCombo) return;
            
            try
            {
                if (processComboBox.SelectedIndex >= 0 && processComboBox.SelectedIndex < processComboBox.Items.Count)
                {
                    selectedProcessName = processComboBox.SelectedItem.ToString();
                    startStopButton.Enabled = false;
                    RefreshHandleList();
                }
                else
                {
                    selectedProcessName = "";
                    isUpdatingHandleCombo = true;
                    handleComboBox.Items.Clear();
                    allHandles.Clear();
                    handleComboBox.Enabled = false;
                    handleRefreshButton.Enabled = false;
                    startStopButton.Enabled = false;
                    isUpdatingHandleCombo = false;
                }
            }
            catch
            {
            }
        }

        private void RefreshProcessList()
        {
            Task.Run(async () =>
            {
                await RefreshProcessListAsync();
            });
        }

        private async Task RefreshProcessListAsync()
        {
            var processes = new List<string>();

            await Task.Run(() =>
            {
                foreach (var proc in Process.GetProcesses())
                {
                    try
                    {
                        string displayName = $"{proc.ProcessName} (PID: {proc.Id})";
                        processes.Add(displayName);
                    }
                    catch { }
                }
            });

            allProcesses = processes.OrderBy(p => p).ToList();

            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    isUpdatingProcessCombo = true;
                    string savedText = processComboBox.SelectedItem?.ToString() ?? "";
                    processComboBox.Items.Clear();
                    processComboBox.Items.AddRange(allProcesses.ToArray());
                    if (!string.IsNullOrEmpty(savedText) && allProcesses.Contains(savedText))
                    {
                        processComboBox.SelectedItem = savedText;
                    }
                    isUpdatingProcessCombo = false;
                });
            }
            else
            {
                isUpdatingProcessCombo = true;
                string savedText = processComboBox.SelectedItem?.ToString() ?? "";
                processComboBox.Items.Clear();
                processComboBox.Items.AddRange(allProcesses.ToArray());
                if (!string.IsNullOrEmpty(savedText) && allProcesses.Contains(savedText))
                {
                    processComboBox.SelectedItem = savedText;
                }
                isUpdatingProcessCombo = false;
            }
        }

        private void RefreshHandleList()
        {
            allHandles.Clear();
            isUpdatingHandleCombo = true;
            int savedIndex = handleComboBox.SelectedIndex;
            string savedText = handleComboBox.SelectedItem?.ToString() ?? "";
            handleComboBox.Items.Clear();
            handleComboBox.Enabled = false;
            handleRefreshButton.Enabled = false;
            startStopButton.Enabled = false;
            isUpdatingHandleCombo = false;

            if (string.IsNullOrEmpty(selectedProcessName))
                return;

            string processName = selectedProcessName.Split('(')[0].Trim();
            var processes = Process.GetProcessesByName(processName);
            if (processes.Length == 0)
                return;

            uint processId = (uint)processes[0].Id;
            var handles = GetProcessHandles(processId);

            if (handles.Count > 0)
            {
                allHandles = handles.OrderBy(h => h).ToList();
                isUpdatingHandleCombo = true;
                handleComboBox.Items.AddRange(allHandles.ToArray());
                handleComboBox.Enabled = true;
                handleRefreshButton.Enabled = true;
                if (!string.IsNullOrEmpty(savedText) && allHandles.Contains(savedText))
                {
                    handleComboBox.SelectedItem = savedText;
                    if (processComboBox.SelectedIndex >= 0)
                    {
                        startStopButton.Enabled = true;
                    }
                }
                isUpdatingHandleCombo = false;
            }
        }

        private List<string> GetProcessHandles(uint processId)
        {
            var handles = new List<string>();
            IntPtr processHandle = OpenProcess(PROCESS_DUP_HANDLE | PROCESS_QUERY_LIMITED_INFORMATION, false, processId);
            if (processHandle == IntPtr.Zero)
            {
                return handles;
            }

            try
            {
                int bufferSize = 128 * 1024 * 1024;
                IntPtr buffer = Marshal.AllocHGlobal(bufferSize);

                try
                {
                    int returnLength;
                    uint status = NtQuerySystemInformation(SystemHandleInformation, buffer, bufferSize, out returnLength);

                    int maxRetries = 5;
                    for (int attempt = 0; attempt < maxRetries && status == STATUS_INFO_LENGTH_MISMATCH; attempt++)
                    {
                        bufferSize = returnLength + (4 * 1024 * 1024);
                        Marshal.FreeHGlobal(buffer);
                        buffer = Marshal.AllocHGlobal(bufferSize);
                        status = NtQuerySystemInformation(SystemHandleInformation, buffer, bufferSize, out returnLength);
                    }

                    if (status != 0)
                    {
                        return handles;
                    }
                    
                    int handleCount = Marshal.ReadInt32(buffer);
                    IntPtr offset = new IntPtr(buffer.ToInt64() + 8);
                    
                    try
                    {
                        var proc = Process.GetProcessById((int)processId);
                    }
                    catch
                    {
                        return handles;
                    }
                    
                    int entrySize = Marshal.SizeOf<SYSTEM_HANDLE_TABLE_ENTRY_INFO>();
                    IntPtr currentOffset = offset;
                    var seenHandles = new HashSet<string>();

                    for (int i = 0; i < handleCount; i++)
                    {
                        SYSTEM_HANDLE_TABLE_ENTRY_INFO entry = Marshal.PtrToStructure<SYSTEM_HANDLE_TABLE_ENTRY_INFO>(currentOffset);
                        uint pid = entry.ProcessId;

                        if (pid == processId)
                        {
                            string handleName = GetHandleName(processHandle, entry.HandleValue);
                            if (handleName != null && !string.IsNullOrWhiteSpace(handleName))
                            {
                                handleName = handleName.Trim();
                                if (!seenHandles.Contains(handleName))
                                {
                                    seenHandles.Add(handleName);
                                    handles.Add(handleName);
                                }
                            }
                        }

                        currentOffset = new IntPtr(currentOffset.ToInt64() + entrySize);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            finally
            {
                CloseHandle(processHandle);
            }

            return handles.OrderBy(h => h).ToList();
        }

        private void StartStopButton_Click(object sender, EventArgs e)
        {
            if (isMonitoring)
            {
                StopMonitoring();
            }
            else
            {
                string processText = processComboBox.SelectedItem?.ToString() ?? processComboBox.Text;
                string handleText = handleComboBox.SelectedItem?.ToString() ?? handleComboBox.Text;

                if (string.IsNullOrWhiteSpace(processText) || string.IsNullOrWhiteSpace(handleText))
                {
                    MessageBox.Show("Please select both a process and handle before starting.", "Cannot Start", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selectedProcessName = processText;
                selectedHandleName = handleText;
                StartMonitoring();
            }
        }

        private void StopMonitoring()
        {
            isMonitoring = false;
            cancellationTokenSource?.Cancel();
            startStopButton.Text = "Start";
            startStopButton.BackColor = Color.FromArgb(60, 60, 70);
            startStopButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            startStopButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);
            processComboBox.Enabled = true;
            handleComboBox.Enabled = handleComboBox.Items.Count > 0;
            handleRefreshButton.Enabled = handleComboBox.Enabled;
            UpdateStatus("Ready", Color.FromArgb(220, 220, 220));
            startStopButton.Invalidate();
        }

        public void StartMonitoring()
        {
            if (isMonitoring) return;

            isMonitoring = true;
            cancellationTokenSource = new CancellationTokenSource();
            startStopButton.Text = "Stop";
            startStopButton.BackColor = Color.FromArgb(220, 53, 69);
            startStopButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 35, 51);
            startStopButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(180, 25, 41);
            processComboBox.Enabled = false;
            handleComboBox.Enabled = false;
            startStopButton.Invalidate();

            Task.Run(() => MonitorLoop(cancellationTokenSource.Token));
        }

        private void MonitorLoop(CancellationToken cancellationToken)
        {
            EnableDebugPrivilege();

            var pidToInstanceNumber = new Dictionary<uint, int>();
            var pidLastCheckTime = new Dictionary<uint, long>();
            var pidHandleClosed = new HashSet<uint>();
            int nextInstanceNumber = 1;

            string targetProcessName = selectedProcessName.Split('(')[0].Trim();
            string targetHandleName = selectedHandleName.ToLower();

            UpdateStatus($"Monitoring {targetProcessName} - Closing handle: {selectedHandleName}", Color.FromArgb(100, 181, 246));

            while (!cancellationToken.IsCancellationRequested && isMonitoring)
            {
                var targetPids = FindProcessesByName(targetProcessName);
                var targetPidsSet = new HashSet<uint>(targetPids);
                long currentTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

                foreach (var pid in targetPids)
                {
                    if (cancellationToken.IsCancellationRequested) break;

                    if (!pidToInstanceNumber.ContainsKey(pid))
                    {
                        int instanceNumber = nextInstanceNumber++;
                        pidToInstanceNumber[pid] = instanceNumber;
                        pidLastCheckTime[pid] = 0;
                    }

                    int instanceNumberForPid = pidToInstanceNumber[pid];
                    long lastCheck = pidLastCheckTime.ContainsKey(pid) ? pidLastCheckTime[pid] : 0;
                    long timeSinceLastCheck = currentTime - lastCheck;

                    if (timeSinceLastCheck < 500)
                    {
                        continue;
                    }

                    pidLastCheckTime[pid] = currentTime;

                    IntPtr testHandle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
                    if (testHandle == IntPtr.Zero)
                    {
                        continue;
                    }
                    CloseHandle(testHandle);

                    if (CloseHandleForProcess(pid, targetHandleName, instanceNumberForPid))
                    {
                        if (!pidHandleClosed.Contains(pid))
                        {
                            pidHandleClosed.Add(pid);
                            UpdateStatus($"Handle closed for {targetProcessName} Instance {instanceNumberForPid} (PID: {pid})", Color.FromArgb(76, 175, 80));
                        }
                    }
                }

                foreach (var pid in pidToInstanceNumber.Keys.ToList())
                {
                    if (!targetPidsSet.Contains(pid))
                    {
                        pidToInstanceNumber.Remove(pid);
                        pidLastCheckTime.Remove(pid);
                        pidHandleClosed.Remove(pid);
                    }
                }

                Thread.Sleep(300);
            }

            UpdateStatus("Monitoring stopped", Color.FromArgb(220, 220, 220));
        }

        private void UpdateStatus(string message, Color? color = null)
        {
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    statusLabel.Text = $"Status: {message}";
                    if (color.HasValue)
                    {
                        statusLabel.ForeColor = color.Value;
                    }
                });
            }
            else
            {
                statusLabel.Text = $"Status: {message}";
                if (color.HasValue)
                {
                    statusLabel.ForeColor = color.Value;
                }
            }
        }

        private bool EnableDebugPrivilege()
        {
            try
            {
                IntPtr hToken;
                if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out hToken))
                    return false;

                try
                {
                    LUID luid;
                    if (!LookupPrivilegeValue(null, "SeDebugPrivilege", out luid))
                        return false;

                    TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES
                    {
                        PrivilegeCount = 1,
                        Luid = luid,
                        Attributes = SE_PRIVILEGE_ENABLED
                    };

                    return AdjustTokenPrivileges(hToken, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                }
                finally
                {
                    CloseHandle(hToken);
                }
            }
            catch
            {
                return false;
            }
        }

        private uint[] FindProcessesByName(string processName)
        {
            var processes = new List<uint>();

            foreach (var proc in Process.GetProcesses())
            {
                try
                {
                    if (proc.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                    {
                        processes.Add((uint)proc.Id);
                    }
                }
                catch { }
            }

            return processes.ToArray();
        }

        private bool CloseHandleForProcess(uint processId, string targetHandleName, int instanceNumber)
        {
            IntPtr processHandle = OpenProcess(PROCESS_DUP_HANDLE | PROCESS_QUERY_LIMITED_INFORMATION, false, processId);
            if (processHandle == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                int bufferSize = 128 * 1024 * 1024;
                IntPtr buffer = Marshal.AllocHGlobal(bufferSize);

                try
                {
                    int returnLength;
                    uint status = NtQuerySystemInformation(SystemHandleInformation, buffer, bufferSize, out returnLength);

                    int maxRetries = 5;
                    for (int attempt = 0; attempt < maxRetries && status == STATUS_INFO_LENGTH_MISMATCH; attempt++)
                    {
                        bufferSize = returnLength + (4 * 1024 * 1024);
                        Marshal.FreeHGlobal(buffer);
                        buffer = Marshal.AllocHGlobal(bufferSize);
                        status = NtQuerySystemInformation(SystemHandleInformation, buffer, bufferSize, out returnLength);
                    }

                    if (status != 0)
                    {
                        return false;
                    }
                    
                    int handleCount = Marshal.ReadInt32(buffer);
                    IntPtr offset = new IntPtr(buffer.ToInt64() + 8);
                    
                    try
                    {
                        var proc = Process.GetProcessById((int)processId);
                    }
                    catch
                    {
                        return false;
                    }
                    
                    int entrySize = Marshal.SizeOf<SYSTEM_HANDLE_TABLE_ENTRY_INFO>();
                    IntPtr currentOffset = offset;

                    for (int i = 0; i < handleCount; i++)
                    {
                        SYSTEM_HANDLE_TABLE_ENTRY_INFO entry = Marshal.PtrToStructure<SYSTEM_HANDLE_TABLE_ENTRY_INFO>(currentOffset);
                        uint pid = entry.ProcessId;

                        if (pid == processId)
                        {
                            string handleName = GetHandleName(processHandle, entry.HandleValue);
                            if (handleName != null)
                            {
                                handleName = handleName.Trim();
                                string handleNameLower = handleName.ToLower();
                                
                                bool isTargetHandle = handleNameLower.Contains(targetHandleName) ||
                                    handleNameLower.Equals(targetHandleName) ||
                                    (targetHandleName.Contains("singleton") && handleNameLower.Contains("singleton") && 
                                     (handleNameLower.Contains("roblox") || targetHandleName.Contains("roblox")));
                                
                                if (isTargetHandle)
                                {
                                    IntPtr duplicatedHandle;
                                    uint dupStatus = NtDuplicateObject(
                                        processHandle,
                                        new IntPtr(entry.HandleValue),
                                        GetCurrentProcess(),
                                        out duplicatedHandle,
                                        0,
                                        0,
                                        DUPLICATE_CLOSE_SOURCE);

                                    if (dupStatus == 0)
                                    {
                                        CloseHandle(duplicatedHandle);
                                        return true;
                                    }
                                }
                            }
                        }

                        currentOffset = new IntPtr(currentOffset.ToInt64() + entrySize);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            finally
            {
                CloseHandle(processHandle);
            }

            return false;
        }

        private string GetHandleName(IntPtr processHandle, ushort handleValue)
        {
            IntPtr duplicatedHandle;
            uint status = NtDuplicateObject(processHandle, new IntPtr(handleValue), GetCurrentProcess(), out duplicatedHandle, 0, 0, 0);
            if (status != 0)
            {
                return null;
            }

            try
            {
                int bufferSize = 4096;
                IntPtr buffer = Marshal.AllocHGlobal(bufferSize);

                try
                {
                    int returnLength;
                    status = NtQueryObject(duplicatedHandle, ObjectNameInformation, buffer, bufferSize, out returnLength);

                    if (status == 0)
                    {
                        ushort length = (ushort)Marshal.ReadInt16(buffer);
                        ushort maxLength = (ushort)Marshal.ReadInt16(new IntPtr(buffer.ToInt64() + 2));
                        
                        IntPtr stringPtr = Marshal.ReadIntPtr(new IntPtr(buffer.ToInt64() + 4));
                        long bufferStart = buffer.ToInt64();
                        
                        if (stringPtr == IntPtr.Zero || 
                            (stringPtr.ToInt64() < bufferStart || stringPtr.ToInt64() >= bufferStart + bufferSize))
                        {
                            IntPtr alignedPtr = Marshal.ReadIntPtr(new IntPtr(buffer.ToInt64() + 8));
                            if (alignedPtr != IntPtr.Zero && 
                                alignedPtr.ToInt64() >= bufferStart && 
                                alignedPtr.ToInt64() < bufferStart + bufferSize)
                            {
                                stringPtr = alignedPtr;
                            }
                        }
                        
                        if (length > 0 && length < bufferSize && length % 2 == 0)
                        {
                            IntPtr finalStringPtr = IntPtr.Zero;
                            
                            if (stringPtr != IntPtr.Zero)
                            {
                                long stringAddr = stringPtr.ToInt64();
                                if (stringAddr >= bufferStart && stringAddr + length <= bufferStart + bufferSize)
                                {
                                    finalStringPtr = stringPtr;
                                }
                            }
                            
                            if (finalStringPtr == IntPtr.Zero)
                            {
                                IntPtr offset12Ptr = new IntPtr(bufferStart + 12);
                                if (offset12Ptr.ToInt64() + length <= bufferStart + bufferSize)
                                {
                                    finalStringPtr = offset12Ptr;
                                }
                            }
                            
                            if (finalStringPtr == IntPtr.Zero)
                            {
                                IntPtr offset16Ptr = new IntPtr(bufferStart + 16);
                                if (offset16Ptr.ToInt64() + length <= bufferStart + bufferSize)
                                {
                                    finalStringPtr = offset16Ptr;
                                }
                            }
                            
                            if (finalStringPtr != IntPtr.Zero)
                            {
                                int charCount = length / 2;
                                if (charCount > 0 && charCount < 2048)
                                {
                                    string result = Marshal.PtrToStringUni(finalStringPtr, charCount);
                                    if (result != null && result.Length > 0)
                                    {
                                        return result;
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            finally
            {
                CloseHandle(duplicatedHandle);
            }

            return null;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            string processText = processComboBox.SelectedItem?.ToString() ?? "";
            string handleText = handleComboBox.SelectedItem?.ToString() ?? "";

            if (string.IsNullOrWhiteSpace(processText) || string.IsNullOrWhiteSpace(handleText))
            {
                MessageBox.Show("Please select both a process and handle before saving.", "Cannot Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string profileName = ShowInputDialog("Enter a name for this profile:", "Save Profile", $"Profile {DateTime.Now:yyyy-MM-dd HH:mm}");

            if (string.IsNullOrWhiteSpace(profileName))
            {
                return;
            }

            string processNameOnly = processText.Split('(')[0].Trim();

            Settings settings = new Settings
            {
                ProcessName = processNameOnly,
                HandleName = handleText,
                ProfileName = profileName
            };

            try
            {
                SettingsCollection collection = new SettingsCollection();
                
                if (File.Exists(SettingsFile))
                {
                    string existingJson = File.ReadAllText(SettingsFile);
                    try
                    {
                        collection = JsonSerializer.Deserialize<SettingsCollection>(existingJson) ?? new SettingsCollection();
                    }
                    catch
                    {
                        collection = new SettingsCollection();
                    }
                }

                collection.Profiles[profileName] = settings;

                string json = JsonSerializer.Serialize(collection, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFile, json);
                MessageBox.Show($"Profile '{profileName}' saved successfully!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void LoadButton_Click(object sender, EventArgs e)
        {
            await LoadSettingsAsync();
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            if (isMonitoring)
            {
                MessageBox.Show("Please stop monitoring before refreshing the process list.", "Cannot Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            RefreshProcessList();
            handleComboBox.Items.Clear();
            handleComboBox.Enabled = false;
            handleRefreshButton.Enabled = false;
            startStopButton.Enabled = false;
        }

        private void HandleRefreshButton_Click(object sender, EventArgs e)
        {
            if (isMonitoring)
            {
                MessageBox.Show("Please stop monitoring before refreshing the handle list.", "Cannot Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(selectedProcessName))
            {
                MessageBox.Show("Please select a process first.", "Cannot Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            RefreshHandleList();
        }

        private async Task LoadSettingsAsync()
        {
            if (!File.Exists(SettingsFile))
            {
                MessageBox.Show("No saved settings found.", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string json = await Task.Run(() => File.ReadAllText(SettingsFile));
                SettingsCollection collection = null;

                try
                {
                    collection = JsonSerializer.Deserialize<SettingsCollection>(json);
                }
                catch
                {
                    Settings oldSettings = JsonSerializer.Deserialize<Settings>(json);
                    if (oldSettings != null)
                    {
                        collection = new SettingsCollection();
                        collection.Profiles["Default"] = oldSettings;
                    }
                }

                if (collection == null || collection.Profiles == null || collection.Profiles.Count == 0)
                {
                    MessageBox.Show("No saved profiles found.", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var dialog = new Form())
                    {
                        dialog.Text = "Select Profile";
                        dialog.Size = new Size(400, 250);
                        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                        dialog.StartPosition = FormStartPosition.CenterParent;
                        dialog.MaximizeBox = false;
                        dialog.MinimizeBox = false;
                        dialog.BackColor = Color.FromArgb(20, 20, 25);

                        var label = new Label
                        {
                            Text = "Select a profile to load:",
                            Location = new Point(15, 15),
                            Size = new Size(350, 20),
                            ForeColor = Color.FromArgb(220, 220, 220)
                        };
                        dialog.Controls.Add(label);

                        var listBox = new ListBox
                        {
                            Location = new Point(15, 40),
                            Size = new Size(350, 120),
                            BackColor = Color.FromArgb(35, 35, 40),
                            ForeColor = Color.FromArgb(220, 220, 220),
                            BorderStyle = BorderStyle.FixedSingle
                        };
                        listBox.Items.AddRange(collection.Profiles.Keys.ToArray());
                        if (listBox.Items.Count > 0)
                            listBox.SelectedIndex = 0;
                        dialog.Controls.Add(listBox);

                        var loadBtn = new Button
                        {
                            Text = "Load",
                            Location = new Point(15, 170),
                            Size = new Size(75, 30),
                            DialogResult = DialogResult.OK,
                            BackColor = Color.FromArgb(60, 60, 70),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat
                        };
                        loadBtn.FlatAppearance.BorderSize = 0;
                        dialog.Controls.Add(loadBtn);

                        var renameBtn = new Button
                        {
                            Text = "Rename",
                            Location = new Point(100, 170),
                            Size = new Size(75, 30),
                            BackColor = Color.FromArgb(60, 60, 70),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat,
                            Enabled = listBox.Items.Count > 0
                        };
                        renameBtn.FlatAppearance.BorderSize = 0;
                        renameBtn.Click += (s, e) =>
                        {
                            if (listBox.SelectedItem == null) return;
                            string oldName = listBox.SelectedItem.ToString();
                            string newName = ShowInputDialog("Enter new profile name:", "Rename Profile", oldName);
                            if (!string.IsNullOrWhiteSpace(newName) && newName != oldName)
                            {
                                if (collection.Profiles.ContainsKey(newName))
                                {
                                    MessageBox.Show("A profile with that name already exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                                if (collection.Profiles.ContainsKey(oldName))
                                {
                                    collection.Profiles[newName] = collection.Profiles[oldName];
                                    collection.Profiles[newName].ProfileName = newName;
                                    collection.Profiles.Remove(oldName);
                                    
                                    listBox.Items.Clear();
                                    listBox.Items.AddRange(collection.Profiles.Keys.ToArray());
                                    for (int i = 0; i < listBox.Items.Count; i++)
                                    {
                                        if (listBox.Items[i].ToString() == newName)
                                        {
                                            listBox.SelectedIndex = i;
                                            break;
                                        }
                                    }
                                    
                                    try
                                    {
                                        string json = JsonSerializer.Serialize(collection, new JsonSerializerOptions { WriteIndented = true });
                                        File.WriteAllText(SettingsFile, json);
                                        MessageBox.Show("Profile renamed successfully!", "Rename Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"Failed to save: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                        };
                        dialog.Controls.Add(renameBtn);

                        var deleteBtn = new Button
                        {
                            Text = "Delete",
                            Location = new Point(185, 170),
                            Size = new Size(75, 30),
                            BackColor = Color.FromArgb(220, 53, 69),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat,
                            Enabled = listBox.Items.Count > 0
                        };
                        deleteBtn.FlatAppearance.BorderSize = 0;
                        deleteBtn.Click += (s, e) =>
                        {
                            if (listBox.SelectedItem == null) return;
                            string profileName = listBox.SelectedItem.ToString();
                            if (MessageBox.Show($"Are you sure you want to delete '{profileName}'?", "Delete Profile", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                collection.Profiles.Remove(profileName);
                                listBox.Items.Remove(profileName);
                                
                                if (listBox.Items.Count > 0)
                                    listBox.SelectedIndex = 0;
                                else
                                {
                                    renameBtn.Enabled = false;
                                    deleteBtn.Enabled = false;
                                }
                                
                                try
                                {
                                    string json = JsonSerializer.Serialize(collection, new JsonSerializerOptions { WriteIndented = true });
                                    File.WriteAllText(SettingsFile, json);
                                    MessageBox.Show("Profile deleted successfully!", "Delete Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Failed to save: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        };
                        dialog.Controls.Add(deleteBtn);

                        var cancelBtn = new Button
                        {
                            Text = "Cancel",
                            Location = new Point(290, 170),
                            Size = new Size(75, 30),
                            DialogResult = DialogResult.Cancel,
                            BackColor = Color.FromArgb(60, 60, 70),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat
                        };
                        cancelBtn.FlatAppearance.BorderSize = 0;
                        dialog.Controls.Add(cancelBtn);

                        dialog.AcceptButton = loadBtn;
                        dialog.CancelButton = cancelBtn;

                        if (dialog.ShowDialog(this) == DialogResult.OK && listBox.SelectedItem != null)
                        {
                            string selectedProfile = listBox.SelectedItem.ToString();
                            if (collection.Profiles.ContainsKey(selectedProfile))
                            {
                                await LoadProfile(collection.Profiles[selectedProfile]);
                            }
                        }
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task LoadProfile(Settings settings)
        {
            if (settings == null || string.IsNullOrEmpty(settings.ProcessName))
            {
                MessageBox.Show("Invalid profile data.", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (processComboBox.Items.Count == 0)
            {
                await RefreshProcessListAsync();
            }
            
            string savedProcessName = settings.ProcessName.Split('(')[0].Trim();
            bool found = false;
            for (int i = 0; i < processComboBox.Items.Count; i++)
            {
                string comboItem = processComboBox.Items[i].ToString();
                string comboProcessName = comboItem.Split('(')[0].Trim();
                
                if (comboProcessName.Equals(savedProcessName, StringComparison.OrdinalIgnoreCase))
                {
                    isUpdatingProcessCombo = true;
                    processComboBox.SelectedIndex = i;
                    isUpdatingProcessCombo = false;
                    found = true;
                    break;
                }
            }
            
            if (!found)
            {
                MessageBox.Show($"Process '{settings.ProcessName}' not found in current process list.", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            await Task.Delay(100);

            if (!string.IsNullOrEmpty(settings.HandleName))
            {
                RefreshHandleList();
                
                await Task.Delay(200);
                
                found = false;
                for (int i = 0; i < handleComboBox.Items.Count; i++)
                {
                    if (handleComboBox.Items[i].ToString() == settings.HandleName)
                    {
                        isUpdatingHandleCombo = true;
                        handleComboBox.SelectedIndex = i;
                        isUpdatingHandleCombo = false;
                        found = true;
                        break;
                    }
                }
                
                if (!found)
                {
                    MessageBox.Show($"Handle '{settings.HandleName}' not found for the selected process.", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Profile loaded successfully!", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Profile loaded successfully!", "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string ShowInputDialog(string text, string caption, string defaultValue = "")
        {
            Form prompt = new Form
            {
                Width = 400,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(20, 20, 25)
            };
            prompt.Paint += (s, e) =>
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(
                    new Point(0, 0),
                    new Point(0, prompt.Height),
                    Color.FromArgb(25, 25, 30),
                    Color.FromArgb(15, 15, 20)))
                {
                    e.Graphics.FillRectangle(brush, prompt.ClientRectangle);
                }
                
                using (SolidBrush dotBrush = new SolidBrush(Color.FromArgb(10, 255, 255, 255)))
                {
                    Random rnd = new Random(12345);
                    for (int i = 0; i < 20; i++)
                    {
                        int x = rnd.Next(0, prompt.Width);
                        int y = rnd.Next(0, prompt.Height);
                        e.Graphics.FillEllipse(dotBrush, x, y, 2, 2);
                    }
                }
            };

            Label textLabel = new Label
            {
                Left = 15,
                Top = 15,
                Text = text,
                ForeColor = Color.FromArgb(220, 220, 220),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            TextBox textBox = new TextBox
            {
                Left = 15,
                Top = 40,
                Width = 350,
                Text = defaultValue,
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.FromArgb(220, 220, 220),
                BorderStyle = BorderStyle.FixedSingle
            };

            Button confirmation = new Button
            {
                Text = "OK",
                Left = 210,
                Width = 75,
                Top = 75,
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            confirmation.FlatAppearance.BorderSize = 0;
            confirmation.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            confirmation.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);

            Button cancel = new Button
            {
                Text = "Cancel",
                Left = 290,
                Width = 75,
                Top = 75,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(60, 60, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cancel.FlatAppearance.BorderSize = 0;
            cancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 70, 80);
            cancel.FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 50, 60);

            confirmation.Click += (sender, e) => { prompt.Close(); };
            cancel.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(cancel);
            prompt.AcceptButton = confirmation;
            prompt.CancelButton = cancel;

            return prompt.ShowDialog(this) == DialogResult.OK ? textBox.Text : "";
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            if (!IsAdministrator())
            {
                MessageBox.Show(
                    "This program requires administrator privileges!\n\n" +
                    "Please right-click the EXE and select 'Run as administrator'",
                    "Admin Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            MainForm form = new MainForm();
            Application.Run(form);
        }

        private static bool IsAdministrator()
        {
            try
            {
                System.Security.Principal.WindowsIdentity identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }
    }
}
