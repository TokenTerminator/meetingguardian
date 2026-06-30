using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MeetingGuardian
{
    public class TrayApplicationContext : ApplicationContext
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DestroyIcon(IntPtr handle);

        private NotifyIcon trayIcon;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem leadTimeMenu;
        private ToolStripMenuItem startupMenu;
        private ToolStripLabel statusLabel;
        private ToolStripLabel mailboxStatusLabel;

        // Keep unmanaged HICON resource bitmap/icon alive in memory to prevent tray icon disappearing
        private Icon mainAppIcon;
        private Bitmap mainAppBitmap;

        private System.Windows.Forms.Timer scanIntervalTimer;
        private System.Windows.Forms.Timer reminderCheckTimer;

        private int leadMinutes = 0;
        private int refreshMinutes = 15;

        private List<OutlookMeeting> cachedMeetings = new List<OutlookMeeting>();
        private HashSet<string> alertedMeetings = new HashSet<string>();
        private List<ReminderForm> openForms = new List<ReminderForm>();

        private readonly object meetingsLock = new object();
        private readonly object formsLock = new object();
        private bool isScanning = false;
        private string lastScanStatus = "Not scanned yet";

        public TrayApplicationContext()
        {
            Logger.Log("TrayApplicationContext constructor start.");
            
            // Load settings from Registry
            LoadSettings();
            Logger.Log(string.Format("Settings loaded: LeadMinutes={0}, RefreshMinutes={1}", leadMinutes, refreshMinutes));

            // Create tray icon
            mainAppIcon = CreateTrayIcon();
            Logger.Log("Tray icon created.");

            // Create Context Menu
            BuildContextMenu();
            Logger.Log("Context menu built.");

            // Setup System Tray Icon
            trayIcon = new NotifyIcon();
            trayIcon.Icon = mainAppIcon;
            trayIcon.Text = "Meeting Guardian";
            trayIcon.ContextMenuStrip = contextMenu;
            trayIcon.Visible = true;
            Logger.Log("NotifyIcon initialized.");

            // Timer 1: Checks for upcoming alerts in the cached list every 30 seconds
            reminderCheckTimer = new System.Windows.Forms.Timer();
            reminderCheckTimer.Interval = 30000; // 30 seconds
            reminderCheckTimer.Tick += new EventHandler(OnReminderCheckTimerTick);
            reminderCheckTimer.Start();

            // Timer 2: Refreshes cache from Outlook every X minutes
            scanIntervalTimer = new System.Windows.Forms.Timer();
            scanIntervalTimer.Interval = refreshMinutes * 60 * 1000;
            scanIntervalTimer.Tick += new EventHandler(OnScanIntervalTimerTick);
            scanIntervalTimer.Start();
            Logger.Log("Timers started.");

            // Initial immediate scan (on a separate STA thread)
            StartOutlookScan();
        }

        private void LoadSettings()
        {
            leadMinutes = GetRegistryDword("LeadMinutes", 0);
            refreshMinutes = GetRegistryDword("RefreshMinutes", 15);

            int startupConfigured = GetRegistryDword("StartupConfigured", 0);
            if (startupConfigured == 0)
            {
                SetRunningAtStartup(true);
                SetRegistryDword("StartupConfigured", 1);
            }
        }

        private void BuildContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            // Status Label
            statusLabel = new ToolStripLabel();
            statusLabel.Text = "Status: Initializing...";
            statusLabel.ForeColor = Color.Gray;
            statusLabel.Font = new Font(statusLabel.Font, FontStyle.Italic);
            contextMenu.Items.Add(statusLabel);

            // Mailbox Status Label
            mailboxStatusLabel = new ToolStripLabel();
            mailboxStatusLabel.Text = "";
            mailboxStatusLabel.ForeColor = Color.FromArgb(75, 85, 99); // Medium gray
            mailboxStatusLabel.Visible = false;
            contextMenu.Items.Add(mailboxStatusLabel);

            contextMenu.Items.Add(new ToolStripSeparator());

            // Force Refresh Option
            ToolStripMenuItem refreshItem = new ToolStripMenuItem("Refresh Calendar Data", null, new EventHandler(OnForceRefreshClick));
            contextMenu.Items.Add(refreshItem);

            // Lead Time Submenu
            leadTimeMenu = new ToolStripMenuItem("Lead Time Configuration");
            UpdateLeadTimeMenuOptions();
            contextMenu.Items.Add(leadTimeMenu);

            // Close All Reminders
            ToolStripMenuItem closeAllItem = new ToolStripMenuItem("Close All Reminders", null, new EventHandler(OnCloseAllRemindersClick));
            contextMenu.Items.Add(closeAllItem);

            // Run at Windows Startup
            startupMenu = new ToolStripMenuItem("Run at Windows Startup", null, new EventHandler(OnStartupClick));
            startupMenu.Checked = IsRunningAtStartup();
            contextMenu.Items.Add(startupMenu);

            contextMenu.Items.Add(new ToolStripSeparator());

            // Exit
            ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit", null, new EventHandler(OnExitClick));
            contextMenu.Items.Add(exitItem);
        }

        private void UpdateLeadTimeMenuOptions()
        {
            leadTimeMenu.DropDownItems.Clear();

            int[] options = new int[] { 0, 5, 10, 15 };
            bool matchedAny = false;

            for (int i = 0; i < options.Length; i++)
            {
                int min = options[i];
                string text = string.Format("{0} Minutes" + (min == 0 ? " (At Start)" : ""), min);
                ToolStripMenuItem item = new ToolStripMenuItem(text, null, new EventHandler(OnLeadTimeOptionClick));
                item.Tag = min;
                item.Checked = (leadMinutes == min);
                if (item.Checked) matchedAny = true;
                leadTimeMenu.DropDownItems.Add(item);
            }

            leadTimeMenu.DropDownItems.Add(new ToolStripSeparator());

            string customText = string.Format("Custom ({0}m)...", leadMinutes);
            ToolStripMenuItem customItem = new ToolStripMenuItem(customText, null, new EventHandler(OnCustomLeadTimeClick));
            customItem.Tag = -1;
            customItem.Checked = !matchedAny;
            leadTimeMenu.DropDownItems.Add(customItem);
        }

        private void OnLeadTimeOptionClick(object sender, EventArgs e)
        {
            ToolStripMenuItem clickedItem = sender as ToolStripMenuItem;
            if (clickedItem != null)
            {
                int min = (int)clickedItem.Tag;
                leadMinutes = min;
                SetRegistryDword("LeadMinutes", min);
                UpdateLeadTimeMenuOptions();
                CheckUpcomingAlerts(); // Recheck immediately with new lead time
            }
        }

        private void OnCustomLeadTimeClick(object sender, EventArgs e)
        {
            using (CustomLeadMinutesForm form = new CustomLeadMinutesForm(leadMinutes))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    leadMinutes = form.SelectedMinutes;
                    SetRegistryDword("LeadMinutes", leadMinutes);
                    UpdateLeadTimeMenuOptions();
                    CheckUpcomingAlerts(); // Recheck immediately with new lead time
                }
            }
        }

        private void OnForceRefreshClick(object sender, EventArgs e)
        {
            StartOutlookScan();
        }

        private void OnCloseAllRemindersClick(object sender, EventArgs e)
        {
            CloseAllReminders();
        }

        private void OnStartupClick(object sender, EventArgs e)
        {
            bool isChecked = !startupMenu.Checked;
            SetRunningAtStartup(isChecked);
            startupMenu.Checked = isChecked;
        }

        private void OnExitClick(object sender, EventArgs e)
        {
            ExitApplication();
        }

        private void OnReminderCheckTimerTick(object sender, EventArgs e)
        {
            CheckUpcomingAlerts();
        }

        private void OnScanIntervalTimerTick(object sender, EventArgs e)
        {
            StartOutlookScan();
        }

        private void StartOutlookScan()
        {
            Logger.Log("StartOutlookScan requested.");
            lock (meetingsLock)
            {
                if (isScanning)
                {
                    Logger.Log("Outlook scan already in progress. Skipping request.");
                    return;
                }
                isScanning = true;
            }

            UpdateStatusLabel("Scanning Outlook...");

            Logger.Log("Spawning STA background thread for Outlook scan...");
            Thread scanThread = new Thread(new ThreadStart(ScanOutlookProcess));
            scanThread.SetApartmentState(ApartmentState.STA);
            scanThread.Start();
        }

        private void ScanOutlookProcess()
        {
            Logger.Log("ScanOutlookProcess background thread started.");
            List<OutlookMeeting> results = new List<OutlookMeeting>();
            List<string> scannedMailboxes = new List<string>();
            Outlook.Application outlookApp = null;
            Outlook.NameSpace mapiNamespace = null;
            Outlook.Accounts accounts = null;
            bool success = false;
            string errorMsg = null;

            try
            {
                // Initialize Outlook Interop
                try
                {
                    Logger.Log("Connecting to Outlook Interop Application...");
                    outlookApp = new Outlook.Application();
                }
                catch (Exception ex)
                {
                    throw new Exception("Could not start Outlook. Make sure it is installed. Details: " + ex.Message);
                }

                mapiNamespace = outlookApp.GetNamespace("MAPI");
                
                // Logon silently to the default profile
                mapiNamespace.Logon(Type.Missing, Type.Missing, false, false);

                accounts = mapiNamespace.Accounts;
                int accountsCount = accounts.Count;
                Logger.Log(string.Format("Found {0} configured accounts in Outlook profile.", accountsCount));

                for (int i = 1; i <= accountsCount; i++)
                {
                    Outlook.Account account = null;
                    Outlook.Store store = null;
                    Outlook.Folder calendarFolder = null;
                    Outlook.Items folderItems = null;
                    Outlook.Items restrictedItems = null;

                    try
                    {
                        account = accounts[i];
                        string accountName = account.DisplayName ?? string.Format("Account #{0}", i);
                        Logger.Log(string.Format("Checking account {0}/{1}: '{2}'...", i, accountsCount, accountName));

                        try
                        {
                            store = account.DeliveryStore;
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("Failed to retrieve DeliveryStore for account '{0}': {1}", accountName, ex.Message));
                            continue;
                        }

                        if (store == null)
                        {
                            Logger.Log(string.Format("Account '{0}' has no delivery store. Skipping.", accountName));
                            continue;
                        }

                        string mailboxName = store.DisplayName ?? accountName;
                        Logger.Log(string.Format("Scanning mailbox store: '{0}'...", mailboxName));

                        try
                        {
                            calendarFolder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
                        }
                        catch
                        {
                            // This store does not support calendars
                            Logger.Log(string.Format("Mailbox store '{0}' does not support calendars.", mailboxName));
                            continue;
                        }

                        if (calendarFolder == null) continue;

                        // Add to scanned mailboxes list
                        if (!scannedMailboxes.Contains(mailboxName))
                        {
                            scannedMailboxes.Add(mailboxName);
                        }

                        folderItems = calendarFolder.Items;
                        folderItems.Sort("[Start]", false);
                        folderItems.IncludeRecurrences = true;

                        // Only retrieve meetings in the next few hours (from 15 min ago to 4 hours from now)
                        DateTime startRange = DateTime.Now.AddMinutes(-15);
                        DateTime endRange = DateTime.Now.AddHours(4);

                        // Jet filter format: [Start] >= 'MM/dd/yyyy HH:mm'
                        string filter = string.Format("[Start] >= '{0}' AND [Start] <= '{1}'",
                            startRange.ToString("MM/dd/yyyy HH:mm"),
                            endRange.ToString("MM/dd/yyyy HH:mm"));

                        restrictedItems = folderItems.Restrict(filter);
                        Logger.Log(string.Format("Querying calendar items for '{0}'...", mailboxName));

                        object itemObj = null;
                        try
                        {
                            itemObj = restrictedItems.GetFirst();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("Failed to call GetFirst on '{0}': {1}", mailboxName, ex.Message));
                        }

                        int itemsProcessed = 0;
                        while (itemObj != null)
                        {
                            Outlook.AppointmentItem appItem = null;
                            try
                            {
                                appItem = itemObj as Outlook.AppointmentItem;
                                if (appItem != null)
                                {
                                    OutlookMeeting mtg = new OutlookMeeting();
                                    mtg.EntryId = appItem.EntryID;
                                    mtg.Subject = appItem.Subject ?? "(No Subject)";
                                    mtg.StartTime = appItem.Start;
                                    mtg.EndTime = appItem.End;
                                    mtg.Location = appItem.Location ?? "";
                                    mtg.MailboxName = mailboxName;

                                    string body = "";
                                    try
                                    {
                                        body = appItem.Body ?? "";
                                    }
                                    catch { }

                                    mtg.JoinUrl = ExtractJoinUrl(mtg.Location, body);

                                    results.Add(mtg);
                                    itemsProcessed++;
                                }
                            }
                            catch
                            {
                                // Skip individual corrupt or locked meetings
                            }
                            finally
                            {
                                if (appItem != null) Marshal.ReleaseComObject(appItem);
                            }

                            // Get next item and release current
                            object nextObj = null;
                            try
                            {
                                nextObj = restrictedItems.GetNext();
                            }
                            catch (Exception ex)
                            {
                                Logger.Log(string.Format("Failed to call GetNext on '{0}' during loop: {1}", mailboxName, ex.Message));
                            }

                            Marshal.ReleaseComObject(itemObj);
                            itemObj = nextObj;
                        }
                        Logger.Log(string.Format("Finished store '{0}': Processed {1} appointments.", mailboxName, itemsProcessed));
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("Error scanning account '{0}': {1}", account != null ? account.DisplayName : "unknown", ex.Message));
                    }
                    finally
                    {
                        if (restrictedItems != null) Marshal.ReleaseComObject(restrictedItems);
                        if (folderItems != null) Marshal.ReleaseComObject(folderItems);
                        if (calendarFolder != null) Marshal.ReleaseComObject(calendarFolder);
                        if (store != null) Marshal.ReleaseComObject(store);
                        if (account != null) Marshal.ReleaseComObject(account);
                    }
                }

                success = true;
                Logger.Log("Outlook scan succeeded.");
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
                Logger.Log("Outlook scan failed with exception: " + ex.ToString());
            }
            finally
            {
                Logger.Log("Releasing Outlook COM objects...");
                if (accounts != null) Marshal.ReleaseComObject(accounts);
                if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);

                // Run GC to release COM reference count wrapper objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                Logger.Log("Outlook COM objects released.");
            }

            // Report results back to the UI thread
            MethodInvoker reportResults = delegate
            {
                lock (meetingsLock)
                {
                    isScanning = false;
                    if (success)
                    {
                        cachedMeetings = results;
                        lastScanStatus = string.Format("Last sync: {0:HH:mm:ss} ({1} meetings)", DateTime.Now, results.Count);
                        Logger.Log(string.Format("Reporting scan success back to UI. Found {0} meetings.", results.Count));

                        // Group upcoming meetings by mailbox
                        Dictionary<string, int> mboxCounts = new Dictionary<string, int>();
                        for (int k = 0; k < scannedMailboxes.Count; k++)
                        {
                            mboxCounts[scannedMailboxes[k]] = 0;
                        }
                        for (int k = 0; k < results.Count; k++)
                        {
                            string name = results[k].MailboxName;
                            if (mboxCounts.ContainsKey(name))
                            {
                                mboxCounts[name]++;
                            }
                            else
                            {
                                mboxCounts[name] = 1;
                            }
                        }

                        // Format multi-line status text
                        string mboxStatus = "";
                        foreach (KeyValuePair<string, int> kvp in mboxCounts)
                        {
                            mboxStatus += string.Format("{0} - {1} upcoming meetings\n", kvp.Key, kvp.Value);
                        }
                        mboxStatus = mboxStatus.TrimEnd('\n');

                        UpdateMailboxStatusLabel(mboxStatus);
                    }
                    else
                    {
                        lastScanStatus = "Sync failed: " + (errorMsg.Length > 30 ? errorMsg.Substring(0, 30) + "..." : errorMsg);
                        Logger.Log("Reporting scan failure back to UI. Error: " + errorMsg);
                        UpdateMailboxStatusLabel("");
                    }
                }

                UpdateStatusLabel(lastScanStatus);

                if (success)
                {
                    // Check meetings immediately in case they start soon
                    CheckUpcomingAlerts();
                }
            };

            // Call reporting logic back on UI Thread context
            if (trayIcon.ContextMenuStrip != null && trayIcon.ContextMenuStrip.InvokeRequired)
            {
                try
                {
                    trayIcon.ContextMenuStrip.Invoke(reportResults);
                }
                catch { }
            }
            else
            {
                reportResults();
            }
        }

        private void CheckUpcomingAlerts()
        {
            DateTime now = DateTime.Now;
            // Scan for upcoming meetings starting in the window [now - 5 min, now + leadMinutes].
            // To be resilient to timer intervals, sleep states, and 0-minute lead times,
            // we trigger an alert if the meeting starts within the last 5 minutes or up to leadMinutes
            // from now, provided it hasn't ended yet and hasn't been alerted for.
            DateTime targetLimit = now.AddMinutes(leadMinutes);
            DateTime startLimit = now.AddMinutes(-5);

            List<OutlookMeeting> triggers = new List<OutlookMeeting>();

            lock (meetingsLock)
            {
                for (int i = 0; i < cachedMeetings.Count; i++)
                {
                    OutlookMeeting mtg = cachedMeetings[i];
                    
                    // Alert if the meeting starts in the range [now - 5 minutes, now + leadMinutes]
                    // and hasn't ended yet.
                    if (mtg.StartTime >= startLimit && mtg.StartTime <= targetLimit && mtg.EndTime > now)
                    {
                        // Create a unique alert key so we don't trigger the same meeting twice at the same lead time
                        string alertKey = string.Format("{0}_{1}_{2}", mtg.EntryId, mtg.StartTime.Ticks, leadMinutes);
                        
                        if (!alertedMeetings.Contains(alertKey))
                        {
                            alertedMeetings.Add(alertKey);
                            triggers.Add(mtg);
                        }
                    }
                }
            }

            // Spawn reminders in separate threads (non-blocking)
            for (int i = 0; i < triggers.Count; i++)
            {
                ShowReminder(triggers[i]);
            }
        }

        private void ShowReminder(OutlookMeeting meeting)
        {
            Logger.Log(string.Format("ShowReminder called for subject='{0}' start='{1:HH:mm:ss}' mailbox='{2}'", meeting.Subject, meeting.StartTime, meeting.MailboxName));
            Thread formThread = new Thread(delegate()
            {
                Logger.Log(string.Format("Spawning form thread for meeting: {0}", meeting.Subject));
                ReminderForm form = new ReminderForm(meeting);
                
                lock (formsLock)
                {
                    openForms.Add(form);
                    Logger.Log(string.Format("Form added to active open list. Active form count={0}", openForms.Count));
                }

                form.FormClosed += delegate(object s, FormClosedEventArgs e)
                {
                    lock (formsLock)
                    {
                        openForms.Remove(form);
                        Logger.Log(string.Format("Form closed and removed. Active form count={0}", openForms.Count));
                    }
                };

                Logger.Log(string.Format("Running Application.Run for form: {0}", meeting.Subject));
                Application.Run(form);
                Logger.Log(string.Format("Form thread terminated for: {0}", meeting.Subject));
            });

            formThread.SetApartmentState(ApartmentState.STA);
            formThread.Start();
        }

        private void CloseAllReminders()
        {
            List<ReminderForm> formsToClose;
            lock (formsLock)
            {
                formsToClose = new List<ReminderForm>(openForms);
            }

            for (int i = 0; i < formsToClose.Count; i++)
            {
                ReminderForm form = formsToClose[i];
                try
                {
                    if (form.InvokeRequired)
                    {
                        form.BeginInvoke(new MethodInvoker(form.Close));
                    }
                    else
                    {
                        form.Close();
                    }
                }
                catch { }
            }
        }

        private void UpdateStatusLabel(string text)
        {
            MethodInvoker update = delegate
            {
                statusLabel.Text = text;
            };

            if (contextMenu.InvokeRequired)
            {
                try { contextMenu.Invoke(update); } catch { }
            }
            else
            {
                update();
            }
        }

        private void UpdateMailboxStatusLabel(string text)
        {
            MethodInvoker update = delegate
            {
                if (string.IsNullOrEmpty(text))
                {
                    mailboxStatusLabel.Visible = false;
                }
                else
                {
                    mailboxStatusLabel.Text = text;
                    mailboxStatusLabel.Visible = true;
                }
            };

            if (contextMenu.InvokeRequired)
            {
                try { contextMenu.Invoke(update); } catch { }
            }
            else
            {
                update();
            }
        }

        private void ExitApplication()
        {
            // Stop timers
            if (scanIntervalTimer != null) scanIntervalTimer.Stop();
            if (reminderCheckTimer != null) reminderCheckTimer.Stop();

            // Close all alerts
            CloseAllReminders();

            // Clear tray icon
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }

            // Destroy the unmanaged icon handle to prevent leaks
            if (mainAppIcon != null)
            {
                try
                {
                    DestroyIcon(mainAppIcon.Handle);
                    mainAppIcon.Dispose();
                }
                catch { }
            }

            if (mainAppBitmap != null)
            {
                try
                {
                    mainAppBitmap.Dispose();
                }
                catch { }
            }

            ExitThread();
        }

        #region Helper: Programmatic Icon Drawing
        private Icon CreateTrayIcon()
        {
            // Try loading from Embedded Resource first (allows single exe distribution)
            try
            {
                System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (System.IO.Stream stream = assembly.GetManifestResourceStream("MeetingGuardian.appicon2.png"))
                {
                    if (stream != null)
                    {
                        mainAppBitmap = new Bitmap(stream);
                        mainAppIcon = Icon.FromHandle(mainAppBitmap.GetHicon());
                        return mainAppIcon;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to load embedded appicon2.png resource: " + ex.Message);
            }

            // Fallback: Try loading from disk file
            try
            {
                string pngPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appicon2.png");
                if (!System.IO.File.Exists(pngPath))
                {
                    pngPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\appicon2.png");
                }

                if (System.IO.File.Exists(pngPath))
                {
                    mainAppBitmap = new Bitmap(pngPath);
                    mainAppIcon = Icon.FromHandle(mainAppBitmap.GetHicon());
                    return mainAppIcon;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to load appicon2.png file: " + ex.Message);
            }

            try
            {
                mainAppBitmap = new Bitmap(16, 16);
                using (Graphics g = Graphics.FromImage(mainAppBitmap))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                    // Draw shield polygon: high contrast yellow/gold background (#f59e0b)
                    using (Brush brush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                    {
                        Point[] shieldPoints = new Point[]
                        {
                            new Point(8, 1),
                            new Point(14, 3),
                            new Point(14, 9),
                            new Point(8, 15),
                            new Point(2, 9),
                            new Point(2, 3)
                        };
                        g.FillPolygon(brush, shieldPoints);
                    }

                    // Draw a black checkmark inside the shield for high contrast
                    using (Pen pen = new Pen(Color.Black, 2f))
                    {
                        g.DrawLine(pen, 5, 8, 7, 10);
                        g.DrawLine(pen, 7, 10, 11, 5);
                    }
                }
                mainAppIcon = Icon.FromHandle(mainAppBitmap.GetHicon());
                return mainAppIcon;
            }
            catch
            {
                mainAppIcon = SystemIcons.Application;
                return mainAppIcon;
            }
        }
        #endregion

        #region Helper: Registry Configuration Settings
        private static int GetRegistryDword(string name, int defaultValue)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Meeting Guardian"))
                {
                    if (key != null)
                    {
                        object val = key.GetValue(name);
                        if (val != null)
                        {
                            return Convert.ToInt32(val);
                        }
                    }
                }
            }
            catch { }
            return defaultValue;
        }

        private static void SetRegistryDword(string name, int value)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Meeting Guardian"))
                {
                    if (key != null)
                    {
                        key.SetValue(name, value, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
            }
            catch { }
        }

        private bool IsRunningAtStartup()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
                {
                    if (key != null)
                    {
                        object val = key.GetValue("Meeting Guardian");
                        return val != null;
                    }
                }
            }
            catch { }
            return false;
        }

        private void SetRunningAtStartup(bool enable)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null)
                    {
                        if (enable)
                        {
                            string path = "\"" + Application.ExecutablePath + "\"";
                            key.SetValue("Meeting Guardian", path);
                        }
                        else
                        {
                            key.DeleteValue("Meeting Guardian", false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to update startup registry key: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Helper: Extract Meeting URLs
        private static string ExtractJoinUrl(string location, string body)
        {
            if (!string.IsNullOrEmpty(location))
            {
                string locUrl = FindUrl(location);
                if (!string.IsNullOrEmpty(locUrl)) return locUrl;
            }
            if (!string.IsNullOrEmpty(body))
            {
                return FindUrl(body);
            }
            return null;
        }

        private static string FindUrl(string text)
        {
            try
            {
                // Teams regex
                System.Text.RegularExpressions.Match teamsMatch =
                    System.Text.RegularExpressions.Regex.Match(text, @"https://teams\.microsoft\.com/l/meetup-join/[^\s""<>']+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (teamsMatch.Success) return teamsMatch.Value;

                // Zoom regex
                System.Text.RegularExpressions.Match zoomMatch =
                    System.Text.RegularExpressions.Regex.Match(text, @"https://[^\s""<>']+\.zoom\.us/[j|w]/[^\s""<>']+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (zoomMatch.Success) return zoomMatch.Value;

                // General fallback for any meeting link
                System.Text.RegularExpressions.Match generalMatch =
                    System.Text.RegularExpressions.Regex.Match(text, @"https://[^\s""<>']+(?:join|meeting)[^\s""<>']+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (generalMatch.Success) return generalMatch.Value;
            }
            catch { }
            return null;
        }
        #endregion
    }

    #region Helper Class: OutlookMeeting Data Structure
    public class OutlookMeeting
    {
        public string EntryId { get; set; }
        public string Subject { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Location { get; set; }
        public string MailboxName { get; set; }
        public string JoinUrl { get; set; }
    }
    #endregion

    #region Helper Form: Custom Lead Minutes Input Dialog
    public class CustomLeadMinutesForm : Form
    {
        private NumericUpDown numInput;
        private Button btnOk;
        private Button btnCancel;
        public int SelectedMinutes { get; set; }

        public CustomLeadMinutesForm(int currentMinutes)
        {
            this.Text = "Custom Lead Time";
            this.Size = new Size(240, 140);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(30, 30, 36);
            this.ForeColor = Color.White;

            Label lbl = new Label();
            lbl.Text = "Enter lead time in minutes (0 - 120):";
            lbl.Location = new Point(12, 12);
            lbl.Size = new Size(200, 20);

            numInput = new NumericUpDown();
            numInput.Minimum = 0;
            numInput.Maximum = 120;
            numInput.Value = currentMinutes;
            numInput.Location = new Point(15, 38);
            numInput.Size = new Size(80, 20);
            numInput.BackColor = Color.FromArgb(45, 45, 50);
            numInput.ForeColor = Color.White;

            btnOk = new Button();
            btnOk.Text = "OK";
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Location = new Point(125, 36);
            btnOk.Size = new Size(80, 24);
            btnOk.FlatStyle = FlatStyle.Flat;
            btnOk.ForeColor = Color.White;
            btnOk.FlatAppearance.BorderColor = Color.FromArgb(59, 130, 246);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(125, 66);
            btnCancel.Size = new Size(80, 24);
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.ForeColor = Color.LightGray;
            btnCancel.FlatAppearance.BorderColor = Color.Gray;

            this.Controls.Add(lbl);
            this.Controls.Add(numInput);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                SelectedMinutes = (int)numInput.Value;
            }
            base.OnFormClosing(e);
        }
    }
    #endregion
}
