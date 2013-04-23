using System;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace Gardiner.LoadedProjects
{
    public class ProjectEnumerator
    {
        private readonly IVsSolution _solution;

        public ProjectEnumerator( IVsSolution solution )
        {
            if ( solution == null )
                throw new ArgumentNullException( "solution" );
            this._solution = solution;
        }

        public IEnumerable<IVsHierarchy> LoadedProjects
        {
            get
            {
                IEnumHierarchies hierEnum = GetProjectEnum( __VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION );
                return EnumerateProjects( hierEnum );
            }
        }

        public IEnumerable<IVsHierarchy> UnloadedProjects
        {
            get
            {
                IEnumHierarchies hierEnum = GetProjectEnum( __VSENUMPROJFLAGS.EPF_UNLOADEDINSOLUTION );
                return EnumerateProjects( hierEnum );
            }
        }

        private IEnumerable<IVsHierarchy> EnumerateProjects( IEnumHierarchies hierEnum )
        {
            IVsHierarchy[] hierFetched = new IVsHierarchy[2];
            uint fetchCount;
            int res = VSConstants.S_OK;
            while ( ( ( res = hierEnum.Next( (uint) hierFetched.Length, hierFetched, out fetchCount ) ) ==
                      VSConstants.S_OK ) &&
                    ( fetchCount == hierFetched.Length ) )
            {
                foreach ( IVsHierarchy hier in hierFetched )
                {
                    yield return hier;
                }
            }
            // If Next returns less than the number we asked for it will return S_FALSE and the count of items it returned, so mop
            // those up here. This only matters if you change the hierFetched array above to hold more than a single item.
            if ( fetchCount != 0 )
            {
                for ( int i = 0; i < fetchCount; ++i )
                {
                    yield return hierFetched[ i ];
                }
            }
        }

        private IEnumHierarchies GetProjectEnum( __VSENUMPROJFLAGS enumFlags )
        {
            Guid ignored = Guid.Empty;
            IEnumHierarchies hierEnum;

            ErrorHandler.ThrowOnFailure( this._solution.GetProjectEnum( (uint) enumFlags, ref ignored, out hierEnum ) );
            return hierEnum;
        }
    }

    public static class ExtensionMethods
    {
        public static void ForEach<T>( this IEnumerable<T> collection, Action<T> func )
        {
            if ( collection == null )
                throw new ArgumentNullException( "collection" );
            if ( func == null )
                throw new ArgumentNullException( "func" );

            foreach ( T t in collection )
            {
                func( t );
            }
        }
    }
}