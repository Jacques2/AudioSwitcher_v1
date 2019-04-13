using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.Observables;
using FortyOne.AudioSwitcher.HotKeyData;
using FortyOne.AudioSwitcher.Properties;

namespace FortyOne.AudioSwitcher
{
    public partial class AudioSwitcher : Form
    {

        private static AudioSwitcher _instance;
        private readonly Icon _originalTrayIcon;

        private readonly Dictionary<DeviceIcon, string> ICON_MAP = new Dictionary<DeviceIcon, string>
        {
            {DeviceIcon.Speakers, "3010"},
            {DeviceIcon.Headphones, "3011"},
            {DeviceIcon.LineIn, "3012"},
            {DeviceIcon.Digital, "3013"},
            {DeviceIcon.DesktopMicrophone, "3014"},
            {DeviceIcon.Headset, "3015"},
            {DeviceIcon.Phone, "3016"},
            {DeviceIcon.Monitor, "3017"},
            {DeviceIcon.StereoMix, "3018"},
            {DeviceIcon.Kinect, "3020"},
            {DeviceIcon.Unknown, "3020"}
        };

        private DeviceState _deviceStateFilter = DeviceState.Active;
        public bool DisableHotKeyFunction = false;

        public AudioSwitcher()
        {
            InitializeComponent();
            HandleCreated += AudioSwitcher_HandleCreated;

            //creates file if it does not exist
            StreamWriter setupToggles = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "itemsToToggle.dat");
            setupToggles.Close();

            _originalTrayIcon = new Icon(notifyIcon1.Icon, 32, 32);

            Program.Settings.AutoStartWithWindows = Program.Settings.AutoStartWithWindows;
            checkBoxStartup.Checked = Program.Settings.AutoStartWithWindows;

            HotKeyManager.HotKeyPressed += HotKeyManager_HotKeyPressed;
            AudioDeviceManager.Controller.AudioDeviceChanged.Subscribe(AudioDeviceManager_AudioDeviceChanged);

            ChangeMute();

            MinimizeFootprint();
        }

        public static AudioSwitcher Instance
        {
            get { return _instance ?? (_instance = new AudioSwitcher()); }
        }

        public bool TrayIconVisible
        {
            get { return notifyIcon1.Visible; }
            set
            {
                try
                {
                    notifyIcon1.Visible = value;
                }
                catch
                {
                } // rubbish error
            }
        }

        public string AssemblyVersion
        {
            get { return FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion; }
        }

        private void AudioSwitcher_HandleCreated(object sender, EventArgs e)
        {
            BeginInvoke(new Action(Form_Load));
        }
        private void Form_Load()
        {
            var icon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables("%windir%\\system32\\control.exe"));
            LoadItemsToToggle();
            if (icon != null)
            {
                using (var i = new Bitmap(25, 25))
                using (var g = Graphics.FromImage(i))
                {
                    g.DrawImage(new Bitmap(icon.ToBitmap(), 25, 25), new Rectangle(0, 0, 25, 25));
                    g.FillRectangle(Brushes.Black, new Rectangle(0, 13, 12, 12));
                    g.DrawImage(Resources.shortcut, new Rectangle(1, 14, 10, 10));
                }

                icon.Dispose();

                checkBoxDisabledDevices.Checked = Program.Settings.ShowDisabledDevices;

                //Set, or remove the disconnected filter
                if (Program.Settings.ShowDisabledDevices)
                    _deviceStateFilter |= DeviceState.Disabled;
                else
                    _deviceStateFilter ^= DeviceState.Disabled;
            }
            Hide();
            MinimizeFootprint();
            ChangeMute();
        }

        private void ChangeMute()
        {
            bool mute = Program.Settings.MuteDevices;
            IEnumerable<IDevice> list = AudioDeviceManager.Controller.GetPlaybackDevices(_deviceStateFilter).ToList();
            IEnumerable<IDevice> list2 = AudioDeviceManager.Controller.GetCaptureDevices(_deviceStateFilter).ToList();
            foreach (var ad in list)
            {
                if (itemsToToggle.Contains(ad.Id))
                {
                    var dev = (IDevice)ad;
                    dev.SetMuteAsync(mute);
                }
            }
            foreach (var ad in list2)
            {
                if (itemsToToggle.Contains(ad.Id))
                {
                    var dev = (IDevice)ad;
                    dev.SetMuteAsync(mute);
                }
            }
            if (mute)
            {
                notifyIcon1.Icon = Properties.Resources.red;
                checkBoxMute.Checked = true;
            }
            else
            {
                notifyIcon1.Icon = Properties.Resources.green;
                checkBoxMute.Checked = false;
            }
        }
        private void AudioDeviceManager_AudioDeviceChanged(DeviceChangedArgs e)
        {
            Action refreshAction = () => { };

            if (InvokeRequired)
                BeginInvoke(refreshAction);
            else
                refreshAction();
        }

        [DllImport("psapi.dll")]
        private static extern int EmptyWorkingSet(IntPtr hwProc);

        private static void MinimizeFootprint()
        {
            EmptyWorkingSet(Process.GetCurrentProcess().Handle);
        }

        private void memoryCleaner_Tick(object sender, EventArgs e)
        {
            MinimizeFootprint();
        }

        private void preferencesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Show();
        }

        private static Icon ExtractIconFromPath(string path)
        {
            try
            {
                var iconPath = path.Split(',');
                Icon icon;
                if (iconPath.Length == 2)
                    icon = IconExtractor.Extract(Environment.ExpandEnvironmentVariables(iconPath[0]),
                        Int32.Parse(iconPath[1]), true);
                else
                    icon = new Icon(iconPath[0]);

                return icon;
            }
            catch
            {
                //return a digital as a place holder
                return IconExtractor.Extract(Environment.ExpandEnvironmentVariables("%windir%\\system32\\mmres.dll"), -3013, true);
            }
        }

        List<Guid> itemsToToggle = new List<Guid>();

        private void SaveItemsToToggle()
        {
            StreamWriter sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "itemsToToggle.dat");
            for (int i = 0; i < itemsToToggle.Count; i++)
            {
                sw.WriteLine(itemsToToggle[i]);
            }
            sw.Close();
        }
        private void LoadItemsToToggle()
        {
            itemsToToggle.Clear();
            StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "itemsToToggle.dat");
            string lineRead = sr.ReadLine();
            while (lineRead != null)
            {
                itemsToToggle.Add(Guid.Parse(lineRead));
                lineRead = sr.ReadLine();
            }
            sr.Close();
        }

        private void RefreshNotifyIconItems()
        {
            notifyIconStrip.Items.Clear();
            var playbackCount = 0;
            var recordingCount = 0;
            checkBoxDisabledDevices.Checked = Program.Settings.ShowDisabledDevices;
            if (Program.Settings.ShowDisabledDevices)
                _deviceStateFilter |= DeviceState.Disabled;

            IEnumerable<IDevice> list = AudioDeviceManager.Controller.GetPlaybackDevices(_deviceStateFilter).ToList();

            foreach (var ad in list)
            {
                if (FavouriteDeviceManager.FavouriteDeviceCount > 0 && !FavouriteDeviceManager.IsFavouriteDevice(ad))
                    continue;

                var item = new ToolStripMenuItem
                {
                    Text = ad.FullName,
                    Tag = ad,
                    Checked = false
                };
                if (itemsToToggle.Contains(ad.Id))
                {
                    item.Checked = true;
                }
                notifyIconStrip.Items.Add(item);
                playbackCount++;
            }

            if (playbackCount > 0)
                notifyIconStrip.Items.Add(new ToolStripSeparator());

            list = AudioDeviceManager.Controller.GetCaptureDevices(_deviceStateFilter).ToList();

            foreach (var ad in list)
            {
                if (FavouriteDeviceManager.FavouriteDeviceCount > 0 && !FavouriteDeviceManager.IsFavouriteDevice(ad))
                    continue;

                var item = new ToolStripMenuItem
                {
                    Text = ad.FullName,
                    Tag = ad,
                    Checked = false
                };
                if (itemsToToggle.Contains(ad.Id))
                {
                    item.Checked = true;
                }
                notifyIconStrip.Items.Add(item);
                recordingCount++;
            }

            if (recordingCount > 0)
                notifyIconStrip.Items.Add(new ToolStripSeparator());

            notifyIconStrip.Items.Add(preferencesToolStripMenuItem);

            notifyIconStrip.Items.Add(exitToolStripMenuItem);

            var defaultDevice = AudioDeviceManager.Controller.DefaultPlaybackDevice;
            var notifyText = "Audio Switcher";

            //The maximum length of the noitfy text is 64 characters. This keeps it under
            if (defaultDevice != null)
            {
                var devName = defaultDevice.FullName ?? defaultDevice.Name ?? notifyText;

                if (devName.Length >= 64)
                    notifyText = devName.Substring(0, 60) + "...";
                else
                    notifyText = devName;
            }

            notifyIcon1.Text = notifyText;
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Program.Settings.CloseToTray)
                {
                    e.Cancel = true;
                    Hide();
                    MinimizeFootprint();
                }
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            ToggleMute();
        }
        private void ToggleMute()
        {
            if (checkBoxMute.CheckState == CheckState.Checked)
            {
                checkBoxMute.CheckState = CheckState.Unchecked;
            }
            else
            {
                checkBoxMute.CheckState = CheckState.Checked;
            }
            ChangeMute();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            //RefreshPlaybackDevices();
            //RefreshRecordingDevices();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void notifyIconStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem != null && e.ClickedItem.Tag is IDevice)
            {
                IDevice dev = (IDevice)e.ClickedItem.Tag;
                if (itemsToToggle.Contains(dev.Id))
                {
                    itemsToToggle.Remove(dev.Id);
                }
                else
                {
                    itemsToToggle.Add(dev.Id);
                }
                SaveItemsToToggle();
                ChangeMute();
            }
        }
        private void AudioSwitcher_ResizeEnd(object sender, EventArgs e)
        {
            Program.Settings.WindowWidth = Width;
            Program.Settings.WindowHeight = Height;
        }
        private void NotifyIcon1_MouseClick_1(object sender, MouseEventArgs e)
        {
            RefreshNotifyIconItems();
        }

        private void CheckBoxDisabledDevices_CheckedChanged(object sender, EventArgs e)
        {
            Program.Settings.ShowDisabledDevices = checkBoxDisabledDevices.Checked;
            //Set, or remove the disconnected filter
            if (Program.Settings.ShowDisabledDevices)
                _deviceStateFilter |= DeviceState.Disabled;
            else
                _deviceStateFilter ^= DeviceState.Disabled;
        }

        private void CheckBoxMute_CheckedChanged(object sender, EventArgs e)
        {
            Program.Settings.MuteDevices = checkBoxMute.Checked;

            //mute or unmute devices
            if (Program.Settings.MuteDevices)
                _deviceStateFilter |= DeviceState.Disabled;
            else
                _deviceStateFilter ^= DeviceState.Disabled;
            ChangeMute();
        }

        private void CheckBoxStartup_CheckedChanged(object sender, EventArgs e)
        {
            Program.Settings.AutoStartWithWindows = checkBoxStartup.Checked;
        }

        private void ButtonHotKey_Click(object sender, EventArgs e)
        {
            HotKeyManager.ClearAll();
            var hkf = new HotKeyForm();
            hkf.ShowDialog(this);
        }

        private void HotKeyManager_HotKeyPressed(object sender, EventArgs e)
        {
            //Double check here before handling
            if (DisableHotKeyFunction || Program.Settings.DisableHotKeys)
                return;

            if (sender is HotKey)
            {
                ToggleMute();
            }
        }
    }
}