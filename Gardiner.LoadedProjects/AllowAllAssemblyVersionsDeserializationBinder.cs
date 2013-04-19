using System.Linq;
using System.Collections.Generic;
using System;
using System.Reflection;
using System.Runtime.Serialization;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    internal sealed class AllowAllAssemblyVersionsDeserializationBinder : SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            String currentAssembly = Assembly.GetExecutingAssembly().FullName;

            // In this case we are always using the current assembly
            assemblyName = currentAssembly;

            // Get the type using the typeName and assemblyName
            Type typeToDeserialize = Type.GetType(String.Format("{0}, {1}", typeName, assemblyName));

            return typeToDeserialize;
        }
    }
}