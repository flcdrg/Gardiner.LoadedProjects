using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Gardiner.LoadedProjects
{
    internal class NativeMethods
    {
        [DllImport( "shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int PathRelativePathTo( StringBuilder pszPath, string pszFrom, int dwAttrFrom, string pszTo, int dwAttrTo );
    }
}