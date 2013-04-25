using System;
using Microsoft.VisualStudio.Shell.Interop;

namespace Gardiner.LoadedProjects
{
    // http://social.msdn.microsoft.com/Forums/en-US/vsx/thread/60fdd7b4-2247-4c18-b1da-301390edabf3
    public sealed class HierarchyPathPair
    {
        public HierarchyPathPair( IVsUIHierarchy hier, string hierPath )
        {
            _hierarchy = hier;
            _hierarchyPath = hierPath;
        }

        private readonly string _hierarchyPath;
        private readonly IVsUIHierarchy _hierarchy;

        /// <summary>
        /// File system path of item, relative to solution root
        /// </summary>
        public string HierarchyPath
        {
            get { return _hierarchyPath; }
        }

        public IVsUIHierarchy Hierarchy
        {
            get { return _hierarchy; }
        }
    }
}