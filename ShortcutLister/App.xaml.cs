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
        private DateTime LatestFileCreateDate = DateTime.MinValue;
        private DateTime LatestFileModifiedDate = DateTime.MinValue;
        private int FileCount = 0;

        private readonly String COMMAND_CLOSE = "Close";
        private readonly String COMMAND_SETTINGS = "Settings";

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
            NotifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            NotifyIcon.Click += NotifyIcon_Click;

            //NotifyIcon.ContextMenuStrip = CreateContextMenu();
            CheckContextMenu();     // This sets the ContextMenuStrip, as well as other tracking variables

            iconStream.Dispose();

            return;
        }

        private void NotifyIcon_Click(object sender, EventArgs e)
        {
            // Treat left click as showing/hiding the options as well
            if (NotifyIcon != null && NotifyIcon.ContextMenuStrip != null)
            {
                System.Windows.Forms.MouseEventArgs args = (System.Windows.Forms.MouseEventArgs)e;

                if (args.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    if (NotifyIcon.ContextMenuStrip.Visible == false)
                        NotifyIcon.ContextMenuStrip.Show(System.Windows.Forms.Cursor.Position);
                    else
                        NotifyIcon.ContextMenuStrip.Visible = false;
                }
            }
            return;
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            ShowSettings();
            return;
        }

        /// <summary>
        /// Shows the settings window
        /// </summary>
        private void ShowSettings()
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
            ShortcutHelper helper = new ShortcutHelper(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), FolderName));
            List<ShortcutItem> listShortcuts = null;
            System.Windows.Forms.ToolStripMenuItem menuItem = null;

            contextMenu = new System.Windows.Forms.ContextMenuStrip();
            listShortcuts = helper.GetShortcutItems();

            AddItemsToMenu(contextMenu.Items, listShortcuts);

            // Add Separator bar
            contextMenu.Items.Add(new System.Windows.Forms.ToolStripSeparator());

            // Add Settings option
            menuItem = new System.Windows.Forms.ToolStripMenuItem();
            menuItem.Text = ShortcutLister.Properties.Resources.Settings;
            menuItem.Click += MenuItem_Click;
            menuItem.Tag = COMMAND_SETTINGS;
            contextMenu.Items.Add(menuItem);

            // Add Exit option
            menuItem = new System.Windows.Forms.ToolStripMenuItem();
            menuItem.Text = ShortcutLister.Properties.Resources.Close;
            menuItem.Click += MenuItem_Click;
            menuItem.Tag = COMMAND_CLOSE;
            contextMenu.Items.Add(menuItem);

            return contextMenu;
        }

        private void AddItemsToMenu(System.Windows.Forms.ToolStripItemCollection menuItems, List<ShortcutItem> listShortcuts)
        {
            if (menuItems == null)
            {
                Log.Error("AddItemsToMenu() failed, menu is null");
                return;
            }
            if (listShortcuts == null)
            {
                Log.Error("AddItemsToMEnu() failed, shortcut list is null");
                return;
            }

            System.Drawing.Icon icon = null;
            System.Windows.Forms.ToolStripMenuItem menuItem = null;

            foreach (ShortcutItem item in listShortcuts)
            {                
                icon = null;
                if (item.IconFile != null && item.IconFile.Length > 0)
                {
                    // Use icon file if it exists
                    try
                    {
                        icon = System.Drawing.Icon.ExtractAssociatedIcon(item.IconFile);
                    }
                    catch (Exception e)
                    {
                        Log.Error("AddItemsToMenu() extracting IconFile", e);
                    }
                }                
                if (icon == null && item.TargetFileName != null && item.TargetFileName.Length > 0)
                {
                    // If no icon still, grab from target exe
                    try
                    {
                        icon = System.Drawing.Icon.ExtractAssociatedIcon(item.TargetFileName);
                    }
                    catch (Exception e)
                    {
                        Log.Error("AddITemsToMenu() extracting icon from TargetFileName", e);
                    }
                }
                if (item.Children != null && item.Children.Count > 0)
                    icon = null;        // Folder icon

                menuItem = new System.Windows.Forms.ToolStripMenuItem();
                menuItem.Text = item.DisplayName;                
                if (icon != null)
                    menuItem.Image = icon.ToBitmap();
                menuItem.Click += MenuItem_Click;
                menuItem.Tag = item;

                if (item.Children != null && item.Children.Count > 0)
                {
                    AddItemsToMenu(menuItem.DropDownItems, item.Children);
                }

                if (icon != null)
                    icon.Dispose();

                menuItems.Add(menuItem);
            }

            return;
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
            ShortcutItem item = null;
            String sCommand = null;

            if (itemSelected.Tag is String)
            {
                sCommand = (String)itemSelected.Tag;
            }
            else if (itemSelected.Tag is ShortcutItem)
            {
                item = (ShortcutItem)itemSelected.Tag;
            }
            else
            {
                Log.Warning("MenuItem_Click() unsupported type");
            }

            if (item != null)
            {
                // This is a shortcut, just launch it
                if (item.TargetFileName != null && item.TargetFileName.Length > 0)
                {
                    process.Arguments = item.Arguments;
                    process.WorkingDirectory = item.WorkingDirectory;
                    process.FileName = item.TargetFileName;

                    if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_RESTORE)
                        process.WindowStyle = ProcessWindowStyle.Normal;
                    else if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_MINIMIZE)
                        process.WindowStyle = ProcessWindowStyle.Minimized;
                    else if (item.ShowCommand == ShortcutItem.SHOWCOMMAND_MAXIMIZE)
                        process.WindowStyle = ProcessWindowStyle.Maximized;

                    try
                    {
                        Process.Start(process);
                    }
                    catch (Exception exception)
                    {
                        Log.Error("MenuItem_Click() Process.Start()", exception);
                    }
                }
                else
                {
                    Log.Warning("MenuItem_Click() blank shortcut target");
                }
            }
            else if (sCommand != null)
            {
                // Command by string. Special cases.
                if (sCommand.Equals(COMMAND_CLOSE))
                    Shutdown();
                else if (sCommand.Equals(COMMAND_SETTINGS))
                    ShowSettings();
                else
                    Log.Warning("MenuItem_Click() unsupported string command: " + sCommand);
            }

            return;
        }

        /// <summary>
        /// Initializes settings window. Only needs to be called once. Does not show it (call ShowSettingsWindow() to show it).
        /// </summary>
        private void InitializeSettingsWindow()
        {
            SettingsWindow = new ShortcutLister.MainWindow();
            SettingsWindow.StateChanged += SettingsWindow_StateChanged;

            return;
        }

        private void SettingsWindow_StateChanged(object sender, EventArgs e)
        {
            if (SettingsWindow != null && SettingsWindow.WindowState == WindowState.Minimized)
                SettingsWindow.Visibility = Visibility.Hidden;
            return;
        }

        /// <summary>
        /// Checks if the contents of the shortcut folder have changed, and if so, recreates the context menu.
        /// </summary>
        public void CheckContextMenu(bool bForceRemake = false)
        {
            if (NotifyIcon == null)
            {
                Log.Error("CheckContextMenu() invalid notify icon");
                return;
            }

            ShortcutHelper helper = new ShortcutHelper(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), FolderName));
            DateTime dtLatestCreateDate = DateTime.MinValue;
            DateTime dtLatestModifiedDate = DateTime.MinValue;
            int iFileCount = 0;

            helper.GetFolderInfo(out dtLatestCreateDate, out dtLatestModifiedDate, out iFileCount);
            if (NotifyIcon.ContextMenuStrip == null || bForceRemake  == true || dtLatestCreateDate > LatestFileCreateDate || dtLatestModifiedDate > LatestFileModifiedDate || iFileCount != FileCount)
            {
                // Remember for next check, so we don't rebuild the menu unneeded
                LatestFileCreateDate = dtLatestCreateDate;
                LatestFileModifiedDate = dtLatestModifiedDate;
                FileCount = iFileCount;

                // Rebuild menu
                NotifyIcon.ContextMenuStrip = CreateContextMenu();
                NotifyIcon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;
            }

            return;
        }

        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Verify shortcuts haven't been added/removed since last time we showed the context menu.
            CheckContextMenu(false);
            return;
        }
    }
}
