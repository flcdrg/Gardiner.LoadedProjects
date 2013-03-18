using System;
using Microsoft.VisualStudio.Shell.Interop;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    public sealed class HierarchyPathPair
    {
        public HierarchyPathPair( IVsUIHierarchy hier, string hierPath )
        {
            Hierarchy = hier;
            HierarchyPath = hierPath;
        }

        public readonly string HierarchyPath;
        public readonly IVsUIHierarchy Hierarchy;
    }
}