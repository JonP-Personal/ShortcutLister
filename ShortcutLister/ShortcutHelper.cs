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
            List<ShortcutItem> list = new List<ShortcutItem>();
            string[] sFiles = null;
            ShortcutItem item = null;
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder folder = null;
            Shell32.FolderItem folderItem = null;
            Shell32.ShellLinkObject shellLink = null;
            string sPathOnly = null;
            string sFileOnly = null;           

            sFiles = Directory.GetFiles(Folder, "*.lnk");
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

            return list;
        }

    }
}
