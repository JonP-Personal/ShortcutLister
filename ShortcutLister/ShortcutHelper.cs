using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ShortcutLister
{
    class ShortcutItem
    {
        public static readonly int SHOWCOMMAND_RESTORE = 1;
        public static readonly int SHOWCOMMAND_MINIMIZE = 2;
        public static readonly int SHOWCOMMAND_MAXIMIZE = 3;

        public String DisplayName = null;
        public String ShortcutFileName = null;
        public String TargetFileName = null;
        public String IconFile = null;

        public String Arguments = null;
        public String WorkingDirectory = null;
        public int ShowCommand = 0;

        public List<ShortcutItem> Children = null;
    }

    class ShortcutHelper
    {
        private String Folder = null;

        public ShortcutHelper(String sFolder)
        {
            Folder = sFolder;
        }

        public List<ShortcutItem> GetShortcutItems()
        {
            return GetShortcutItems(Folder);
        }

        public List<ShortcutItem> GetShortcutItems(String sFolder)
        {
            List<ShortcutItem> list = new List<ShortcutItem>();
            string[] sFiles = null;
            ShortcutItem item = null;
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder folder = null;
            Shell32.FolderItem folderItem = null;
            Shell32.ShellLinkObject shellLink = null;
            string sPathOnly = null;
            string sFileOnly = null;           

            // Add all shortcuts
            sFiles = Directory.GetFiles(sFolder, "*.lnk");
            foreach (string sFile in sFiles)
            {
                item = new ShortcutItem();
                item.ShortcutFileName = sFile;

                // Get properties from the shortcut file
                sPathOnly = Path.GetDirectoryName(sFile);
                sFileOnly = Path.GetFileName(sFile);
                folder = shell.NameSpace(sPathOnly);
                folderItem = folder.ParseName(sFileOnly);
                if (folderItem != null)
                {
                    shellLink = folderItem.GetLink;

                    item.TargetFileName = shellLink.Path;
                    item.DisplayName = folderItem.Name;
                    item.Arguments = shellLink.Arguments;
                    item.WorkingDirectory = shellLink.WorkingDirectory;
                    item.ShowCommand = shellLink.ShowCommand;

                    shellLink.GetIconLocation(out item.IconFile);

                    list.Add(item);
                }
            }

            // Add folders
            sFiles = Directory.GetDirectories(sFolder);
            foreach (string sFile in sFiles)
            {
                if (sFile == "." || sFile == "..")
                    continue;

                if (Directory.Exists(sFile) == true)
                {
                    item = new ShortcutItem();
                    item.DisplayName = Path.GetFileName(sFile);
                    item.Children = GetShortcutItems(sFile);

                    if (item.Children != null && item.Children.Count > 0)
                        list.Add(item);
                }
            }

            return list;
        }

        public void GetFolderInfo(out DateTime LatestCreateDateUTC, out DateTime LatestModifiedDateUTC, out int FileCount)
        {
            GetFolderInfo(Folder, out LatestCreateDateUTC, out LatestModifiedDateUTC, out FileCount);
            return;
        }

        public void GetFolderInfo(String sFolder, out DateTime LatestCreateDateUTC, out DateTime LatestModifiedDateUTC, out int FileCount)
        {
            String[] sFiles = null;
            DateTime dtLatestCreateDate = DateTime.MinValue;
            DateTime dtLatestModifiedDate = DateTime.MinValue;
            int iFileCount = 0;
            DateTime dtCheck = DateTime.MinValue;

            sFiles = Directory.GetFiles(sFolder, "*.lnk");
            foreach (string sFile in sFiles)
            {
                if (sFile == null || sFile == "." || sFile == "..")
                    continue;

                dtCheck = File.GetCreationTimeUtc(sFile);
                if (dtCheck > dtLatestCreateDate)
                    dtLatestCreateDate = dtCheck;

                dtCheck = File.GetLastWriteTimeUtc(sFile);
                if (dtCheck > dtLatestModifiedDate)
                    dtLatestModifiedDate = dtCheck;

                iFileCount++;
            }

            LatestCreateDateUTC = dtLatestCreateDate;
            LatestModifiedDateUTC = dtLatestModifiedDate;
            FileCount = iFileCount;

            return;
        }
    }
}
