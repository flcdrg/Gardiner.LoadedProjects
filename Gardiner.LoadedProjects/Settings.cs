using System;
using System.Collections.Generic;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    public class Settings
    {
        public IList<Profile> Profiles { get; private set; }

        public Settings()
        {
            Profiles = new List<Profile>();
        }
    }
}