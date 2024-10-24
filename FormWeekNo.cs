using System;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Threading;

namespace WeekNumberTray
{
    public partial class FormWeekNo : Form
    {
        private NotifyIcon? trayIcon;
        private ContextMenuStrip? trayMenu;
        private System.Windows.Forms.Timer? updateTimer;
        private ToolStripMenuItem? autoStartMenuItem;

        private const string AutoStartRegistryKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private const string AppName = "WeekNumberTray";

        // Declare a Mutex to ensure a single instance of the app
        private static Mutex? appMutex;

        public FormWeekNo()
        {
            // Ensure only one instance of the application is running
            if (!EnsureSingleInstance())
            {
                // If another instance is already running, show a message box and exit
                MessageBox.Show("The application is already running.", "WeekNumberTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0); // Terminate the current instance
            }

            InitializeTrayIcon();
            InitializeTimer();
            UpdateWeekNumberIcon();  // Initial update
        }

        private static bool EnsureSingleInstance()
        {
            appMutex = new Mutex(true, "WeekNumberTrayAppMutex", out var isNewInstance);
            return isNewInstance;
        }

        private void InitializeTrayIcon()
        {
            // Create a new ContextMenuStrip
            trayMenu = new ContextMenuStrip();

            // Add "Auto start" checkbox menu item
            autoStartMenuItem = new ToolStripMenuItem("Auto start", null, OnAutoStartClick)
            {
                CheckOnClick = true
            };
            autoStartMenuItem.Checked = IsAutoStartEnabled();
            trayMenu.Items.Add(autoStartMenuItem);

            trayMenu.Items.Add("Exit", null, (sender, e) =>
            {
                // Close the application
                trayIcon!.Visible = false;
                Application.Exit();
            });

            // Create the NotifyIcon
            trayIcon = new NotifyIcon
            {
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            // Listen for dark mode changes
            SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
        }

        private void InitializeTimer()
        {
            updateTimer = new System.Windows.Forms.Timer();
            updateTimer.Interval = 60 * 60 * 1000; // 60 seconds (update every hour)
            updateTimer.Tick += (sender, e) =>
            {
                UpdateWeekNumberIcon();
            };
            updateTimer.Start();
        }

        // This function checks if the system is in dark mode
        private bool IsDarkMode()
        {
            const string registryKey = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";
            var lightTheme = (int?)Registry.GetValue(registryKey, valueName, 1) ?? 1;
            return lightTheme == 0;  // 0 means dark mode, 1 means light mode
        }

        // Handle dark mode changes
        private void OnUserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
            {
                UpdateWeekNumberIcon();
            }
        }

        private void UpdateWeekNumberIcon()
        {
            int weekNumber = GetWeekNumber(DateTime.Now);
            bool darkMode = IsDarkMode();

            // Use different font colors depending on the theme
            Color textColor = darkMode ? Color.Black : Color.Black;
            Color backColor = darkMode ? Color.White: Color.Transparent; // Adjust background if needed

            trayIcon!.Icon = CreateTextIcon($"{weekNumber}", textColor, backColor);
            //trayIcon!.Icon = CreateTextIcon($"{weekNumber}","W", textColor, backColor);
            trayIcon.Text = $"Week: {weekNumber}";
        }

        private int GetWeekNumber(DateTime date)
        {
            CultureInfo ci = CultureInfo.CurrentCulture;
            return ci.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        }

        // Function to create a text-based icon for the tray
        private Icon CreateTextIcon(string text, Color textColor, Color backColor)
        {
            // Create a 16x16 bitmap
            Bitmap bitmap = new Bitmap(SystemInformation.SmallIconSize.Width, SystemInformation.SmallIconSize.Height);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Clear the background with the specified color
                g.Clear(backColor);

                // Set text rendering options for better quality
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                // Create a font (you can adjust the size to fit the text)
                using (Font font = new Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Pixel))
                {
                    // Measure the text size to center it
                    SizeF textSize = g.MeasureString(text, font);
                    float x = (bitmap.Width - textSize.Width) / 2;
                    float y = (bitmap.Height - textSize.Height) / 2;

                    // Draw the text
                    using (Brush textBrush = new SolidBrush(textColor))
                    {
                        g.DrawString(text, font, textBrush, x, y);
                    }
                }
            }
            // Optionally, convert the bitmap to an icon
            return Icon.FromHandle(bitmap.GetHicon());
        }

        protected override void OnLoad(EventArgs e)
        {
            // Hide the main window
            Visible = false;
            ShowInTaskbar = false;

            base.OnLoad(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                trayIcon?.Dispose();
                updateTimer?.Dispose();
                trayMenu?.Dispose();

                // Unsubscribe from the event
                SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;

                // Release the Mutex
                appMutex?.ReleaseMutex();
            }
            base.Dispose(disposing);
        }

        private void OnAutoStartClick(object? sender, EventArgs e)
        {
            if (autoStartMenuItem!.Checked)
            {
                EnableAutoStart();
            }
            else
            {
                DisableAutoStart();
            }
        }

        private bool IsAutoStartEnabled()
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey))
            {
                if (key != null)
                {
                    string? value = key.GetValue(AppName) as string;
                    return value == Application.ExecutablePath;
                }
            }
            return false;
        }

        private void EnableAutoStart()
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey, true))
            {
                key?.SetValue(AppName, Application.ExecutablePath);
            }
        }

        private void DisableAutoStart()
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryKey, true))
            {
                key?.DeleteValue(AppName, false);
            }
        }
    }
}
