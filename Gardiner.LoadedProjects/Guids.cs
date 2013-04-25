// Guids.cs
// MUST match guids.h

using System;

namespace Gardiner.LoadedProjects
{
    static class GuidList
    {
        public const string guidGardiner_LoadedProjectsPkgString = "4bf1fa9e-38c5-46ca-9718-a133986099dc";
        public const string guidGardiner_LoadedProjectsCmdSetString = "bd1ef86f-540d-4ebe-8ce2-83ab9d142e5b";

        public static readonly Guid guidGardiner_LoadedProjectsCmdSet = new Guid(guidGardiner_LoadedProjectsCmdSetString);
        public const string OutputWindowId = "C376C4E8-8E26-4D6F-886C-551A088EF57D";
    };
}