using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Reflection;
using System.ComponentModel;

namespace ShortcutLister
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void buttonOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            System.IO.Directory.CreateDirectory(App.FolderPath);
            Process.Start(App.FolderPath);            
            return;
        }

        private void LoadSettings()
        {
            checkBoxLaunchOnStartup.IsChecked = Properties.Settings.Default.LaunchOnStartup;
            return;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            e.Cancel = true;        // Keep app minimized instead of closing. Can close via right click on icon.
            Hide();
            SaveSettings();
            return;
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.Save();

            SetStartup(Properties.Settings.Default.LaunchOnStartup);

            return;
        }

        private void SetStartup(bool bLaunchOnStartup)
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            key.SetValue(Properties.Resources.AppName, Assembly.GetExecutingAssembly().Location);
            return;
        }
    }
}
