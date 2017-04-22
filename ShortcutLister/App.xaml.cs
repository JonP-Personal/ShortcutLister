using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;

namespace ShortcutLister
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private System.Windows.Forms.NotifyIcon NotifyIcon = null;
        private MainWindow SettingsWindow = null;

        public static String FolderName = "ShortcutLister";
        public static String FolderPath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), FolderName); }
        }

        public App()
        {
            ShowNotificationIcon();

            // If different language resources are desired later, can add more. I made a temp one for now.
            //ShortcutLister.Properties.Resources.Culture = new System.Globalization.CultureInfo("sv-SE");

            InitializeSettingsWindow();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            RemoveNotificationIcon();            
        }

        /// <summary>
        /// Shows the notification icon in the system tray
        /// </summary>
        private void ShowNotificationIcon()
        {
            Stream iconStream = Application.GetResourceStream(new Uri("pack://application:,,,/Main2.ico")).Stream;

            NotifyIcon = new System.Windows.Forms.NotifyIcon();
            NotifyIcon.Icon = new System.Drawing.Icon(iconStream);
            NotifyIcon.Visible = true;
            NotifyIcon.ContextMenuStrip = CreateContextMenu();
            NotifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            iconStream.Dispose();

            return;
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            if (SettingsWindow != null)
            {
                SettingsWindow.Show();
                SettingsWindow.WindowState = WindowState.Normal;
            }
            return;
        }

        /// <summary>
        /// Removes the notification icon from the system tray
        /// </summary>
        private void RemoveNotificationIcon()
        {
            if (NotifyIcon != null)
                NotifyIcon.Visible = false;
            return;
        }

        /// <summary>
        /// Creates the context menu, which is a list of all Shortcuts in the "Documents\ShortcutLister" folder
        /// </summary>
        /// <returns>The context menu to use</returns>
        private System.Windows.Forms.ContextMenuStrip CreateContextMenu()
        {
            System.Windows.Forms.ContextMenuStrip contextMenu = null;
            System.Windows.Forms.ToolStripMenuItem menuItem = null;
            ShortcutHelper helper = new ShortcutHelper(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), FolderName));
            List<ShortcutItem> listShortcuts = null;
            System.Drawing.Icon icon = null;

            contextMenu = new System.Windows.Forms.ContextMenuStrip();
            listShortcuts = helper.GetShortcutItems();

            foreach (ShortcutItem item in listShortcuts)
            {
                icon = System.Drawing.Icon.ExtractAssociatedIcon(item.TargetFileName);                

                menuItem = new System.Windows.Forms.ToolStripMenuItem();
                menuItem.Text = item.DisplayName;                
                menuItem.Image = icon.ToBitmap();
                menuItem.Click += MenuItem_Click;
                menuItem.Tag = item;

                icon.Dispose();                

                contextMenu.Items.Add(menuItem);
            }
                        
            return contextMenu;
        }

        /// <summary>
        /// Called when menu item clicked. This launches the shortcut that was selected.
        /// </summary>
        /// <param name="sender">The clicked on ToolStripMenuItem object</param>
        /// <param name="e">Event args</param>
        private void MenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolStripMenuItem itemSelected = (System.Windows.Forms.ToolStripMenuItem)sender;
            ProcessStartInfo process = new ProcessStartInfo();
            ShortcutItem item = (ShortcutItem)itemSelected.Tag;

            process.Arguments = item.Arguments;
            process.WorkingDirectory = item.WorkingDirectory;
            process.FileName = item.TargetFileName;

            if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_RESTORE)
                process.WindowStyle = ProcessWindowStyle.Normal;
            else if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_MINIMIZE)
                process.WindowStyle = ProcessWindowStyle.Minimized;
            else if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_MAXIMIZE)
                process.WindowStyle = ProcessWindowStyle.Maximized;

            Process.Start(process);

            return;
        }

        private void InitializeSettingsWindow()
        {
            SettingsWindow = new ShortcutLister.MainWindow();
            SettingsWindow.StateChanged += SettingsWindow_StateChanged;

            return;
        }

        private void SettingsWindow_StateChanged(object sender, EventArgs e)
        {
            if (SettingsWindow.WindowState == WindowState.Minimized)
                SettingsWindow.Visibility = Visibility.Hidden;
            return;
        }
    }
}
