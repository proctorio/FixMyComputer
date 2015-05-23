using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Fix_My_Computer
{
    /// <summary>
    /// Proctorio Inc, Open Source Initiative 2015 https://proctorio.com
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImportAttribute("powrprof.dll", EntryPoint = "PowerSetActiveScheme")]
        public static extern uint PowerSetActiveScheme(IntPtr UserPowerKey, ref Guid ActivePolicyGuid);

        public MainWindow()
        {
            InitializeComponent();
        }

        // remove the window icon, clean is key
        protected override void OnSourceInitialized(EventArgs e) { IconHelper.RemoveIcon(this); }

        // delegates for ui update outside of render thread
        public delegate void UpdateProgressCallback(double v);
        public delegate void ShowCheckCallback();
        public delegate void UpdateTitleCallback(string m);
        public delegate void HideCheckCallback();
        public delegate void SetCursorCallback();
        public delegate void ShowXCallback();
        public delegate void ShowDesktopCCallback();
        public delegate void ShowRebootCallback();
        public delegate void ShowLaptopCCallback();
        public delegate void HideDesktopCCallback();
        public delegate void HideLaptopCCallback();
        public delegate void HideRebootCallback();

        // are we ready for reboot?
        public bool isReady = false;
        public bool isFired = false;

        // thread to animate and unmute microphones
        private void UIThread()
        {
            // hang tight
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("hang tight") });

            // determine laptop vs desktop icon
            // just check for battery... i know, if they remove the battery then this is a false value but whatever
            if(SystemInformation.PowerStatus.BatteryChargeStatus == BatteryChargeStatus.NoSystemBattery)
            {
                DesktopC.Dispatcher.Invoke(new ShowDesktopCCallback(this.ShowDesktopC));
            }
            else
            {
                LaptopC.Dispatcher.Invoke(new ShowDesktopCCallback(this.ShowLaptopC));
            }

            // sleep a bit
            Thread.Sleep(500);

            // find me some microphones
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("finding startup items to fix...") });

            // remove all startup items
            // first pass, stage 1
            // x86
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                RegistryKey machinestartupKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                string[] items = machinestartupKey.GetValueNames();
                int items_length = items.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + items[i].ToLower()) });

                    // delete it, skip nod32
                    if (items[i] != "egui") machinestartupKey.DeleteValue(items[i], false);
                    Thread.Sleep(500);
                }
                machinestartupKey.Close();
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // x64
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                RegistryKey machinestartupKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                string[] items = machinestartupKey.GetValueNames();
                int items_length = items.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + items[i].ToLower()) });

                    // delete it, skip nod32
                    if (items[i] != "egui") machinestartupKey.DeleteValue(items[i], false);
                    Thread.Sleep(500);
                }
                machinestartupKey.Close();
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // second pass, stage 1
            // x86
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                RegistryKey userstartupKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                string[] items = userstartupKey.GetValueNames();
                int items_length = items.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + items[i].ToLower()) });

                    // delete it
                    try
                    {
                        // delete it, skip nod32
                        if (items[i] != "egui") userstartupKey.DeleteValue(items[i], false);
                        Thread.Sleep(500);
                    }
                    catch { }
                }
                userstartupKey.Close();
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // x64
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                RegistryKey userstartupKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                string[] items = userstartupKey.GetValueNames();
                int items_length = items.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + items[i].ToLower()) });

                    // delete it
                    try
                    {
                        // delete it, skip nod32
                        if (items[i] != "egui") userstartupKey.DeleteValue(items[i], false);
                        Thread.Sleep(500);
                    }
                    catch { }
                }
                userstartupKey.Close();
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // third pass, stage 2
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                FileInfo[] thefiles = di.GetFiles("*.lnk");
                int items_length = thefiles.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + System.IO.Path.GetFileNameWithoutExtension(thefiles[i].Name)).ToLower() });

                    // delete it
                    try
                    {
                        File.Delete(thefiles[i].FullName);
                        string itemName = thefiles[i].Name;
                        Thread.Sleep(500);
                    }
                    catch { }
                }
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // fourth pass, stage 2
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
                DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                FileInfo[] thefiles = di.GetFiles("*.exe");
                int items_length = thefiles.Length;
                for (int i = 0; i < items_length; i++)
                {
                    // taking care of..
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { (1.0 - (((double)items_length - (double)i) / (double)items_length)) * 100 });
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("found: " + System.IO.Path.GetFileNameWithoutExtension(thefiles[i].Name)).ToLower() });

                    // delete it
                    try
                    {
                        File.Delete(thefiles[i].FullName);
                        string itemName = thefiles[i].Name;
                        Thread.Sleep(500);
                    }
                    catch { }
                }
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // update power settings to high performance
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("set high performance mode") });
            P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });
            Thread.Sleep(1500);

            // set to high performance
            try
            {
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 33 });
                Guid max = new Guid("8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c");
                PowerSetActiveScheme(IntPtr.Zero, ref max);
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });
                Thread.Sleep(500);
            }
            catch { }

            // hide the icon
            DesktopC.Dispatcher.Invoke(new HideDesktopCCallback(this.hideDesktopC));
            LaptopC.Dispatcher.Invoke(new HideLaptopCCallback(this.HideLaptopC));

            // show the checkmark
            CheckMark.Dispatcher.Invoke(new ShowCheckCallback(this.ShowCheck));

            // done
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "done with computer" });

            // Zzzz
            Thread.Sleep(2000);

            // reset progressbar
            P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });

            // hide the check mark
            CheckMark.Dispatcher.Invoke(new HideCheckCallback(this.HideCheck));

            // click to reboot
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "click anywhere to reboot" });

            // show the reboot
            Reboot.Dispatcher.Invoke(new ShowRebootCallback(this.ShowReboot));

            // set the cursor
            CheckMark.Dispatcher.Invoke(new SetCursorCallback(this.SetCursor));
            isReady = true;                       
        }

        // update the progressbar
        private void UpdateProgress(double v) { P.Value = v; }

        // show the checkmark
        private void ShowCheck() { CheckMark.Visibility = Visibility.Visible; }

        // show the laptop
        private void ShowLaptopC() { LaptopC.Visibility = Visibility.Visible; }

        // show the desktop
        private void ShowDesktopC() { DesktopC.Visibility = Visibility.Visible; }

        // show the reboot
        private void ShowReboot() { Reboot.Visibility = Visibility.Visible; }

        // hide the laptop
        private void HideLaptopC() { LaptopC.Visibility = Visibility.Hidden; }

        // hide the desktop
        private void hideDesktopC() { DesktopC.Visibility = Visibility.Hidden; }

        // title change for progress of events
        private void UpdateTitle(string m) { this.Title = m; }

        // hide the checkmark
        private void HideCheck() { CheckMark.Visibility = Visibility.Hidden; }

        // hide the reboot
        private void HideReboot() { Reboot.Visibility = Visibility.Hidden; }

        // set cursor to hand
        private void SetCursor() { Mouse.OverrideCursor = System.Windows.Input.Cursors.Hand; }

        // show the x
        private void ShowX() { XX.Visibility = Visibility.Visible; }

        // load it up
        private void OnWindowLoaded(object sender, RoutedEventArgs e) { Thread UI = new Thread(new ThreadStart(UIThread)); UI.Start(); }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // did we fire?
            if (isFired)
            {
                // we are done here
                Environment.Exit(0);
            }

            // are we ready?
            if (isReady)
            {
                // yes we have now
                isFired = true;

                // hide reboot
                Reboot.Dispatcher.Invoke(new HideRebootCallback(this.HideReboot));

                bool privilege;
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                privilege = principal.IsInRole(WindowsBuiltInRole.Administrator);
                if (privilege == true)
                {
                    // rebooting
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "rebooting..." });

                    // reboot command
                    Process.Start("ShutDown", "/r");

                    // we are done here
                    Environment.Exit(0);
                }
                else
                {
                    // failure, can't continue
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "unable to automatically restart the computer" });

                    // show failure X
                    XX.Dispatcher.Invoke(new ShowXCallback(this.ShowX));
                }               
            }
        }
    }
}