using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace MeetingGuardian
{
    public static class Logger
    {
        public static void Log(string message)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                File.AppendAllText(logPath, string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1}\r\n", DateTime.Now, message));
            }
            catch { }
        }
    }

    static class Program
    {
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            Logger.Log("Application starting...");
            bool createdNew;
            
            try
            {
                mutex = new Mutex(true, "Global\\MeetingGuardianMutex", out createdNew);
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to create Mutex: " + ex.Message);
                createdNew = true; // fallback
            }

            if (!createdNew)
            {
                Logger.Log("Another instance is already running. Exiting.");
                MessageBox.Show(
                    "Meeting Guardian is already running in the System Tray.",
                    "Meeting Guardian",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Logger.Log("Launching TrayApplicationContext...");
                Application.Run(new TrayApplicationContext());
                Logger.Log("Application context exited cleanly.");
            }
            catch (Exception ex)
            {
                Logger.Log("Fatal error: " + ex.ToString());
                MessageBox.Show(
                    "An unexpected error occurred: " + ex.Message,
                    "Meeting Guardian - Fatal Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (mutex != null)
                {
                    try
                    {
                        mutex.ReleaseMutex();
                        Logger.Log("Mutex released.");
                    }
                    catch (ObjectDisposedException) { }
                    catch (ApplicationException) { }
                    mutex.Dispose();
                }
                Logger.Log("Application shutting down.");
            }
        }
    }
}
