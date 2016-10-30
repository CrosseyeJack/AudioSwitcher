using AudioSwitcher.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;

namespace AudioSwitcher
{
    public class SysTrayApp : Form
    {
        [STAThread]
        public static void Main(string[] args)
        {
            foreach (var arg in args)
            {
                if (arg.ToLower().Equals("autorun"))
                {
                    autorun = true;
                }
            }

            Application.Run(new SysTrayApp());
        }

        private static bool autorun = false;
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private int deviceCount;
        private int currentDeviceId;

        private MenuItem QuitOnComplete;
        private bool QuitOnCompleteFlag = false;
        private bool ChangeOnRunFlag = false;
        private bool RunOnStartUPFlag = false;

        private static List<Tuple<int, string, bool>> devices = new List<Tuple<int, string, bool>>();

        public SysTrayApp()
        {
            // Create a simple tray menu
            trayMenu = new ContextMenu();
            trayIcon = new NotifyIcon();

            QuitOnCompleteFlag = Properties.Settings.Default.QuitOnComplete;
            ChangeOnRunFlag = Properties.Settings.Default.ChangeOnRun;
            RunOnStartUPFlag = Properties.Settings.Default.RunOnStartUp;

            // Run this to make sure that we stop the app running on start up next time.
            if (!RunOnStartUPFlag)
                SetRunOnStartUp(RunOnStartUPFlag);

            devices = GetDevices();
            deviceCount = devices.Count;

            if (ChangeOnRunFlag && !String.IsNullOrEmpty(Properties.Settings.Default.PreferredDevice))
            {
                var prefedDevice = devices.Find(o => o.Item2.Equals(Properties.Settings.Default.PreferredDevice));
                if (prefedDevice != null)
                {
                    // TODO Get current selected device so we don't change to a device already in use causing a every minor blip in audio.
                    SelectDevice(prefedDevice.Item1);
                    Console.WriteLine("Auto Selected {0}", prefedDevice.Item2);
                    if (autorun && QuitOnCompleteFlag)
                    {
                        trayIcon.Dispose();
                        System.Environment.Exit(0);
                    }
                }
            }

            // Create a tray icon
            trayIcon.Text = "AudioSwitcher";
            trayIcon.Icon = new Icon(Resources.speaker, 40, 40);

            // Add menu to tray icon and show it
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;

            // Populate device list when menu is opened
            trayIcon.ContextMenu.Popup += PopulateDeviceList;

            // Register MEH on trayicon leftclick
            trayIcon.MouseUp += new MouseEventHandler(TrayIcon_LeftClick);
        }

        // Brings up the context menu on left click
        private void TrayIcon_LeftClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                mi.Invoke(trayIcon, null);
            }
        }

        //Gets the ID of the next sound device in the list
        // Not using this anymore (nor deviceCount for that matter) but leaving it in the code for now
        private int nextId()
        {
            if (currentDeviceId == deviceCount){
                currentDeviceId = 1;
            } else {
                currentDeviceId += 1;
            }
            return currentDeviceId;
        }

        

        #region Tray events

        private void PopulateDeviceList(object sender, EventArgs e)
        {
            // This list gets populated on start up and on every menu load. Incase more devices are added.
            // TODO Listen to Winodws for device changes and then repopulate the list automatically
            // Empty menu to prevent stuff to pile up
            trayMenu.MenuItems.Clear();

            // All all active devices
            devices = GetDevices();
            deviceCount = devices.Count;
            foreach (var tuple in devices)
            {
                var id = tuple.Item1;
                var deviceName = tuple.Item2;
                var isInUse = tuple.Item3;

                var item = new MenuItem {Checked = isInUse, Text = deviceName};
                item.Click += (s, a) => SelectDevice(id);

                trayMenu.MenuItems.Add(item);
            }

            // Add prefs and exit options
            trayMenu.MenuItems.Add("-"); // Add a Spacer

            var AutoRunSetting = new MenuItem { Text = "Run On Startup", Checked = RunOnStartUPFlag };
            AutoRunSetting.Click += AutoRunSettingAction;
            trayMenu.MenuItems.Add(AutoRunSetting);

            // ChangeOnRun - ChangeOnRunFlag
            var ChangeOnRunMI = new MenuItem { Text = "Change to Pref Device on Run", Checked = ChangeOnRunFlag };
            ChangeOnRunMI.Click += ChangeOnRunAction;
            trayMenu.MenuItems.Add(ChangeOnRunMI);

            // Quit On Change
            QuitOnComplete = new MenuItem { Text = "Quit On Autorun Complete", Checked = QuitOnCompleteFlag };
            QuitOnComplete.Click += QuitOncompleteAction;
            trayMenu.MenuItems.Add(QuitOnComplete);

            trayMenu.MenuItems.Add("-"); // Add a Spacer

            var exitItem = new MenuItem { Text = "Exit" };
            exitItem.Click += OnExit;
            trayMenu.MenuItems.Add(exitItem);
        }

        private void AutoRunSettingAction(object sender, EventArgs e)
        {
            RunOnStartUPFlag = !RunOnStartUPFlag;
            Properties.Settings.Default.RunOnStartUp = RunOnStartUPFlag;
            Properties.Settings.Default.Save();
            // Actually Make the change in Windows to run this application on startup.
            SetRunOnStartUp(RunOnStartUPFlag);
        }

        private void SetRunOnStartUp(bool runOnStartUPFlag)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (runOnStartUPFlag)
                rk.SetValue("AudioSwitcher", "\""+Application.ExecutablePath.ToString() + "\" autorun");
            else
                rk.DeleteValue("AudioSwitcher", false);
        }

        private void ChangeOnRunAction(object sender, EventArgs e)
        {
            ChangeOnRunFlag = !ChangeOnRunFlag;
            Properties.Settings.Default.ChangeOnRun = ChangeOnRunFlag;
            Properties.Settings.Default.Save();
        }

        private void QuitOncompleteAction(object sender, EventArgs e)
        {
            QuitOnCompleteFlag = !QuitOnCompleteFlag;
            QuitOnComplete.Checked = QuitOnCompleteFlag;
            Properties.Settings.Default.QuitOnComplete = QuitOnCompleteFlag;
            Properties.Settings.Default.Save();
        }

        #endregion

        #region EndPointController.exe interaction

        private static List<Tuple<int, string, bool>> GetDevices()
        {
            List<Tuple<int, string, bool>> deviceList = new List<Tuple<int, string, bool>>();
            var p = new Process
                        {
                            StartInfo =
                                {
                                    UseShellExecute = false,
                                    RedirectStandardOutput = true,
                                    CreateNoWindow = true,
                                    FileName = "EndPointController.exe",
                                    Arguments = "-f \"%d|%ws|%d|%d\""
                                }
                        };
            p.Start();
            p.WaitForExit();
            var stdout = p.StandardOutput.ReadToEnd().Trim();
            
            foreach (var line in stdout.Split('\n'))
            {
                var elems = line.Trim().Split('|');
                var deviceInfo = new Tuple<int, string, bool>(int.Parse(elems[0]), elems[1], elems[3].Equals("1"));
                deviceList.Add(deviceInfo);
            }

            return deviceList;
        }

        private static void SelectDevice(int id)
        {
            var selectedDevice = devices.Find(o => o.Item1 == id);
            if (selectedDevice == null)
            {
                // Couldn't find the selected device.
                // TODO Alert the user to try again
                Console.WriteLine("Unable to find device");
            }
            else
            {
                // Found the selected device, Save it as a pref and pass it on to EndPointController.
                Properties.Settings.Default.PreferredDevice = selectedDevice.Item2;
                Properties.Settings.Default.Save();
                var p = new Process
                {
                    StartInfo =
                                {
                                    UseShellExecute = false,
                                    RedirectStandardOutput = true,
                                    CreateNoWindow = true,
                                    FileName = "EndPointController.exe",
                                    Arguments = id.ToString(CultureInfo.InvariantCulture)
                                }
                };
                p.Start();
                p.WaitForExit();
            }
        }

        #endregion

        #region Main app methods

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                trayIcon.Dispose();
            }

            base.Dispose(isDisposing);
        }

        #endregion
    }
}