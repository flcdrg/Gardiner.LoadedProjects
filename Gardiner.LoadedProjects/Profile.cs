using System;
using System.Collections.Generic;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    public class Profile
    {
        public Profile()
        {
            UnloadedProjects = new List<string>();
        }

        public List<string> UnloadedProjects { get; private set; }

        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}