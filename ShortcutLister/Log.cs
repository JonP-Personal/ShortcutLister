using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShortcutLister
{
    public class Log
    {
        public static void Out(String sLine)
        {
            // Add logic to write to file later if desired
            System.Diagnostics.Debug.WriteLine(sLine);
            return;
        }

        public static void Error(String sLine, Exception e)
        {
            Error(sLine + "\r\n" + e.ToString());
            return;
        }

        public static void Error(String sLine)
        {
            Out(sLine);
            System.Diagnostics.Debug.Assert(false, sLine);
            return;
        }

        public static void Warning(String sLine)
        {
            Out(sLine);
            return;
        }
    }
}
