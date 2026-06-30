using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MeetingGuardian
{
    public class ReminderForm : Form
    {
        // P/Invoke for FlashWindowEx
        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        private const uint FLASHW_ALL = 3;
        private const uint FLASHW_TIMERNOFG = 12;

        // Static tracking of open reminder forms for vertical stacking
        private static readonly System.Collections.Generic.List<ReminderForm> activeReminders = 
            new System.Collections.Generic.List<ReminderForm>();
        private static readonly object staticLock = new object();

        private OutlookMeeting meeting;
        private System.Windows.Forms.Timer countdownTimer;
        
        // UI Controls
        private Panel titleBar;
        private Label titleLabel;
        private Button closeButton;
        private Label subjectLabel;
        private Label countdownLabel;
        private Label timeMailboxLabel;
        private Label locationLabel;
        private Button dismissButton;
        private Button joinButton;

        // Dragging states
        private bool dragging = false;
        private Point dragCursorPoint;
        private Point dragFormPoint;

        public ReminderForm(OutlookMeeting meeting)
        {
            this.meeting = meeting;

            // Form properties
            this.Size = new Size(400, 220);
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.FromArgb(30, 30, 36); // Dark mode grey #1E1E24
            this.ForeColor = Color.White;
            this.TopMost = true;
            this.ShowInTaskbar = true;
            this.Text = "Meeting Guardian Reminder";
            this.StartPosition = FormStartPosition.Manual;

            // Load application icon (prioritizing embedded resource for standalone single-exe)
            try
            {
                System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (System.IO.Stream stream = assembly.GetManifestResourceStream("MeetingGuardian.appicon2.png"))
                {
                    if (stream != null)
                    {
                        using (Bitmap bmp = new Bitmap(stream))
                        {
                            this.Icon = Icon.FromHandle(bmp.GetHicon());
                        }
                    }
                    else
                    {
                        // Fallback: load from disk file
                        string pngPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appicon2.png");
                        if (!System.IO.File.Exists(pngPath))
                        {
                            pngPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\appicon2.png");
                        }

                        if (System.IO.File.Exists(pngPath))
                        {
                            using (Bitmap bmp = new Bitmap(pngPath))
                            {
                                this.Icon = Icon.FromHandle(bmp.GetHicon());
                            }
                        }
                    }
                }
            }
            catch { }

            InitializeComponents();

            // Set up second-by-second countdown timer
            countdownTimer = new System.Windows.Forms.Timer();
            countdownTimer.Interval = 1000;
            countdownTimer.Tick += new EventHandler(OnCountdownTimerTick);
            countdownTimer.Start();

            // Run initial tick immediately
            OnCountdownTimerTick(null, null);
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; } // Shows on top but doesn't steal active keyboard focus
        }

        private void InitializeComponents()
        {
            // Title Bar
            titleBar = new Panel();
            titleBar.Size = new Size(this.Width, 35);
            titleBar.BackColor = Color.FromArgb(24, 24, 28);
            titleBar.Location = new Point(0, 0);
            titleBar.MouseDown += new MouseEventHandler(OnTitleBarMouseDown);
            titleBar.MouseMove += new MouseEventHandler(OnTitleBarMouseMove);
            titleBar.MouseUp += new MouseEventHandler(OnTitleBarMouseUp);

            titleLabel = new Label();
            titleLabel.Text = "MEETING GUARDIAN ALERT";
            titleLabel.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
            titleLabel.ForeColor = Color.FromArgb(139, 92, 246); // Purple accent #8b5cf6
            titleLabel.Location = new Point(12, 10);
            titleLabel.AutoSize = true;
            titleLabel.MouseDown += new MouseEventHandler(OnTitleBarMouseDown);
            titleLabel.MouseMove += new MouseEventHandler(OnTitleBarMouseMove);
            titleLabel.MouseUp += new MouseEventHandler(OnTitleBarMouseUp);

            closeButton = new Button();
            closeButton.Text = "X";
            closeButton.Size = new Size(35, 35);
            closeButton.Dock = DockStyle.Right;
            closeButton.FlatStyle = FlatStyle.Flat;
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            closeButton.ForeColor = Color.FromArgb(156, 163, 175);
            closeButton.BackColor = Color.Transparent;
            closeButton.MouseEnter += delegate { closeButton.BackColor = Color.FromArgb(239, 68, 68); closeButton.ForeColor = Color.White; };
            closeButton.MouseLeave += delegate { closeButton.BackColor = Color.Transparent; closeButton.ForeColor = Color.FromArgb(156, 163, 175); };
            closeButton.Click += delegate { this.Close(); };

            titleBar.Controls.Add(titleLabel);
            titleBar.Controls.Add(closeButton);
            this.Controls.Add(titleBar);

            // Subject (Bold & Large)
            subjectLabel = new Label();
            subjectLabel.Text = meeting.Subject;
            subjectLabel.Font = new Font("Segoe UI", 12.5f, FontStyle.Bold);
            subjectLabel.ForeColor = Color.White;
            subjectLabel.Location = new Point(15, 48);
            subjectLabel.Size = new Size(370, 48); // Allow two lines
            subjectLabel.UseMnemonic = false;
            this.Controls.Add(subjectLabel);

            // Countdown Status
            countdownLabel = new Label();
            countdownLabel.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            countdownLabel.Location = new Point(15, 100);
            countdownLabel.Size = new Size(370, 20);
            this.Controls.Add(countdownLabel);

            // Time & Target Mailbox
            timeMailboxLabel = new Label();
            string timeString = string.Format("{0:h:mm tt} - {1:h:mm tt}", meeting.StartTime, meeting.EndTime);
            timeMailboxLabel.Text = string.Format("{0}   |   {1}", timeString, meeting.MailboxName);
            timeMailboxLabel.Font = new Font("Segoe UI", 9f, FontStyle.Regular);
            timeMailboxLabel.ForeColor = Color.FromArgb(209, 213, 219); // #d1d5db
            timeMailboxLabel.Location = new Point(15, 125);
            timeMailboxLabel.Size = new Size(370, 20);
            this.Controls.Add(timeMailboxLabel);

            // Location
            locationLabel = new Label();
            locationLabel.Text = string.Format("Location: {0}", string.IsNullOrEmpty(meeting.Location) ? "(None)" : meeting.Location);
            locationLabel.Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);
            locationLabel.ForeColor = Color.FromArgb(156, 163, 175); // #9ca3af
            locationLabel.Location = new Point(15, 147);
            locationLabel.Size = new Size(370, 18);
            locationLabel.UseMnemonic = false;
            this.Controls.Add(locationLabel);

            // Action: Dismiss Button
            dismissButton = new Button();
            dismissButton.Text = "Dismiss";
            dismissButton.Size = new Size(90, 28);
            dismissButton.Location = new Point(295, 178);
            dismissButton.FlatStyle = FlatStyle.Flat;
            dismissButton.FlatAppearance.BorderSize = 0;
            dismissButton.BackColor = Color.FromArgb(75, 85, 99); // #4b5563
            dismissButton.ForeColor = Color.White;
            dismissButton.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dismissButton.MouseEnter += delegate { dismissButton.BackColor = Color.FromArgb(107, 114, 128); };
            dismissButton.MouseLeave += delegate { dismissButton.BackColor = Color.FromArgb(75, 85, 99); };
            dismissButton.Click += delegate { this.Close(); };
            this.Controls.Add(dismissButton);

            // Action: Join Meeting Button (If Join URL detected)
            if (!string.IsNullOrEmpty(meeting.JoinUrl))
            {
                joinButton = new Button();
                joinButton.Text = "Join Meeting";
                joinButton.Size = new Size(110, 28);
                joinButton.Location = new Point(175, 178);
                joinButton.FlatStyle = FlatStyle.Flat;
                joinButton.FlatAppearance.BorderSize = 0;
                joinButton.BackColor = Color.FromArgb(16, 185, 129); // #10b981 (Emerald)
                joinButton.ForeColor = Color.White;
                joinButton.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                joinButton.MouseEnter += delegate { joinButton.BackColor = Color.FromArgb(52, 211, 153); };
                joinButton.MouseLeave += delegate { joinButton.BackColor = Color.FromArgb(16, 185, 129); };
                joinButton.Click += new EventHandler(OnJoinButtonClick);
                this.Controls.Add(joinButton);
            }

            // Register Paint handler to draw custom border
            this.Paint += new PaintEventHandler(OnFormPaint);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Perform vertical stacking calculation and registration
            lock (staticLock)
            {
                int offset = activeReminders.Count * (this.Height + 10);
                activeReminders.Add(this);

                Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
                this.Left = workingArea.Right - this.Width - 20;
                this.Top = workingArea.Bottom - this.Height - 20 - offset;
            }

            // Start Taskbar flashing
            StartFlashing();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            if (countdownTimer != null)
            {
                countdownTimer.Stop();
                countdownTimer.Dispose();
            }

            // Deregister from vertical stacking list
            lock (staticLock)
            {
                activeReminders.Remove(this);
            }
        }

        private void OnCountdownTimerTick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;

            if (now < meeting.StartTime)
            {
                TimeSpan remaining = meeting.StartTime - now;
                countdownLabel.Text = string.Format("Starts in {0}m {1:d2}s", (int)remaining.TotalMinutes, remaining.Seconds);
                countdownLabel.ForeColor = Color.FromArgb(245, 158, 11); // #f59e0b (Amber)
            }
            else if (now >= meeting.StartTime && now <= meeting.EndTime)
            {
                TimeSpan elapsed = now - meeting.StartTime;
                countdownLabel.Text = string.Format("Ongoing (Started {0}m ago)", (int)elapsed.TotalMinutes);
                countdownLabel.ForeColor = Color.FromArgb(16, 185, 129); // #10b981 (Emerald)
            }
            else
            {
                countdownLabel.Text = "Meeting has ended";
                countdownLabel.ForeColor = Color.FromArgb(239, 68, 68); // #ef4444 (Red)
            }
        }

        private void OnJoinButtonClick(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(meeting.JoinUrl))
            {
                try
                {
                    Process.Start(meeting.JoinUrl);
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not launch meeting link: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void OnFormPaint(object sender, PaintEventArgs e)
        {
            // Draw a subtle 1px border around the form
            using (Pen borderPen = new Pen(Color.FromArgb(139, 92, 246), 1f)) // Purple border #8b5cf6
            {
                e.Graphics.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
            }
        }

        #region Form Dragging Logic
        private void OnTitleBarMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragging = true;
                dragCursorPoint = Cursor.Position;
                dragFormPoint = this.Location;
            }
        }

        private void OnTitleBarMouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point dif = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                this.Location = Point.Add(dragFormPoint, new Size(dif));
            }
        }

        private void OnTitleBarMouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }
        #endregion

        #region P/Invoke Flashing Call
        private void StartFlashing()
        {
            try
            {
                FLASHWINFO fInfo = new FLASHWINFO();
                fInfo.cbSize = Convert.ToUInt32(Marshal.SizeOf(fInfo));
                fInfo.hwnd = this.Handle;
                fInfo.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG;
                fInfo.uCount = uint.MaxValue; // flash continuously
                fInfo.dwTimeout = 0; // use default system blink rate
                
                FlashWindowEx(ref fInfo);
            }
            catch { }
        }
        #endregion
    }
}
