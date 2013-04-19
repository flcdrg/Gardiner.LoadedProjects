using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    public class Profile
    {
        public Profile()
        {
            UnloadedProjects = new List<string>();
        }

        /// <summary>
        /// Array of project paths, relative to solution
        /// </summary>
        [XmlArrayItem("Project")]
        public List<string> UnloadedProjects { get; private set; }

        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}