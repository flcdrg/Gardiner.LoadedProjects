using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    [XmlRoot("Settings")]
    public class Settings
    {
        public List<Profile> Profiles { get; private set; }

        public Settings()
        {
            Profiles = new List<Profile>();
        }
    }
}