using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration( UseManagedResourcesOnly = true )]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration( "#110", "#112", "1.0", IconResourceID = 400 )]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource( "Menus.ctmenu", 1 )]
    [Guid( GuidList.guidGardiner_LoadedProjectsPkgString )]
    public sealed class Gardiner_LoadedProjectsPackage : Package, IVsPersistSolutionOpts
    {

        public delegate void ProcessHierarchyNode( IVsHierarchy hierarchy, uint itemid, int recursionLevel );

        public ProcessHierarchyNode _processNode = new ProcessHierarchyNode( DisplayHierarchyNode );

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public Gardiner_LoadedProjectsPackage()
        {
            Debug.WriteLine( "Entering constructor for: {0}", ToString() ) ;
            AddOptionKey(SettingsKey);
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine("Entering Initialize() of: {0}", ToString() );
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = GetService( typeof( IMenuCommandService ) ) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                var menuCommandID = new CommandID( GuidList.guidGardiner_LoadedProjectsCmdSet, (int) PkgCmdIDList.cmdidLoadedProjects );
                var menuItem = new MenuCommand( MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }

            _settings = new Settings();
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback( object sender, EventArgs e )
        {
            // Show a Message Box to prove we were here
            var uiShell = (IVsUIShell) GetService( typeof (SVsUIShell) );
            Guid clsid = Guid.Empty;
            int result;
            /*            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                                   0,
                                   ref clsid,
                                   "Loaded Projects",
                                   string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
                                   string.Empty,
                                   0,
                                   OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                   OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                                   OLEMSGICON.OLEMSGICON_INFO,
                                   0,        // false
                                   out result));*/

            //TraverseAllItemsFlatCommand();

            Example();

            using ( var frm = new frmProfiles() )
            {
                frm.Unloaded = _unloaded.Select( x => x.HierarchyPath ).ToList();
                frm.Settings = _settings;

                if ( frm.ShowDialog() != DialogResult.Cancel )
                {
                    if ( frm.SelectedProfile != null )
                    {
                        var profile = frm.SelectedProfile;

                        IVsUIHierarchyWindow slnExpHierWin = VsShellUtilities.GetUIHierarchyWindow( this,
                                                                                                    VSConstants
                                                                                                        .StandardToolWindows
                                                                                                        .SolutionExplorer );

                        var solutionService = (IVsSolution) GetService( typeof( SVsSolution ) );

                        foreach ( var project in profile.UnloadedProjects )
                        {
                            Debug.WriteLine( project );

                            var item = _loaded.FirstOrDefault( x => x.HierarchyPath == project );

                            if ( item != null )
                            {
                                if ( slnExpHierWin == null )
                                {
                                    Debug.Fail( "Failed to get the solution explorer hierarchy window!" );
                                }
                                else
                                {

                                    ErrorHandler.ThrowOnFailure(
                                        solutionService.CloseSolutionElement(
                                            (uint) __VSSLNCLOSEOPTIONS.SLNCLOSEOPT_UnloadProject, item.Hierarchy, 0 ) );

                                }

                            }
                        }

                        var dte = (DTE) GetService( typeof( DTE ) );

                        foreach ( var pair in _unloaded )
                        {
                            if (profile.UnloadedProjects.All( x => x != pair.HierarchyPath ))
                            {

                                slnExpHierWin.ExpandItem( pair.Hierarchy, (uint) VSConstants.VSITEMID.Root, EXPANDFLAGS.EXPF_SelectItem );

                                dte.ExecuteCommand( "Project.ReloadProject", "" ); 
                            }
                        }
                    }
                }
            }
        }

        private IList<HierarchyPathPair> _loaded;
        private IList<HierarchyPathPair> _unloaded;


        private static string GetFullPathToItem( IVsHierarchy hier )
        {
            // Most hierarchies in the solution will QI to IVsProject, which allows us to use GetMkDocument to get the full path to the project file,
            // stub hierarchies (i.e. unloaded hierarchies) do not implement IVsProject, luckily their IVsHierarchy::GetCanonicalName returns the
            // full path, so we fall back on that.
            string name;
            var proj = hier as IVsProject;

            object prjObject;
            if ( hier.GetProperty( 0xfffffffe, -2027, out prjObject ) >= 0 )
            {
                var p2 = (EnvDTE.Project) prjObject;

                if ( p2 != null && p2.Kind == ProjectKinds.vsProjectKindSolutionFolder )
                    return string.Empty;

            }


            object pVar;
            hier.GetProperty( (uint) VSConstants.VSITEMID.Root, (int) __VSHPROPID.VSHPROPID_Name, out pVar );
            Debug.WriteLine( pVar );

            if ( proj != null )
            {
                ErrorHandler.ThrowOnFailure( proj.GetMkDocument( (uint) VSConstants.VSITEMID.Root, out name ) );
            }
            else
            {
                ErrorHandler.ThrowOnFailure( hier.GetCanonicalName( (uint) VSConstants.VSITEMID.Root, out name ) );
            }
            return name;
        }


        private void Example()
        {
            //var msg = new StringBuilder( String.Format( "The loaded projects in the current solution are:{0}{0}", Environment.NewLine ) );

            _loaded = new List<HierarchyPathPair>();
            _unloaded = new List<HierarchyPathPair>();

            var enumerator = new ProjectEnumerator( (IVsSolution) GetService( typeof( SVsSolution ) ) );
            var loadedProjects = enumerator.LoadedProjects.Select( h => new HierarchyPathPair( (IVsUIHierarchy) h, GetFullPathToItem( h ) ) );
            loadedProjects.ForEach( hpp => _loaded.Add( hpp ) );


            //msg.AppendFormat( "{0}The unloadedprojects in the current solution are:{0}{0}", Environment.NewLine );
            var unloadedProjects = new List<HierarchyPathPair>( enumerator.UnloadedProjects.Select( h => new HierarchyPathPair( (IVsUIHierarchy) h, GetFullPathToItem( h ) ) ) );

            unloadedProjects.ForEach( hpp => _unloaded.Add( hpp ) );

            //HierarchyPathPair toSelect = unloadedProjects.FirstOrDefault();

            /*
                        if ( toSelect != null )
                        {
                            msg.AppendFormat( "{0} I will select {1} in the solution explorer wheen this dialog is dismissed.", Environment.NewLine, toSelect.HierarchyPath );
                        }
                        VsShellUtilities.ShowMessageBox( this,
                                                        msg.ToString(),
                                                        "Info",
                                                        OLEMSGICON.OLEMSGICON_INFO,
                                                                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                                                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST );
            */
            /*            if ( toSelect != null )
                        {
                            IVsUIHierarchyWindow slnExpHierWin = VsShellUtilities.GetUIHierarchyWindow( this, VSConstants.StandardToolWindows.SolutionExplorer );
                            if ( slnExpHierWin == null )
                            {
                                Debug.Fail( "Failed to get the solution explorer hierarchy window!" );
                            }
                            else
                            {
                                slnExpHierWin.ExpandItem( toSelect.Hierarchy, (uint) VSConstants.VSITEMID.Root, EXPANDFLAGS.EXPF_SelectItem );
                            }
                        }*/
        }

        /// <summary>
        /// Traverse the flat list of project hierarchies in the Solution's Projects Collection.
        /// This enumeration does not recursively traverse into nested hierarchies.
        /// </summary>
        private void TraverseAllItemsFlatCommand()
        {
            //Get the solution service so we can traverse each project hierarchy contained within.
            IVsSolution solution = (IVsSolution) GetService( typeof( SVsSolution ) );
            if ( null != solution )
            {
                IEnumHierarchies penum;
                Guid nullGuid = Guid.Empty;

                //You can ask the solution to enumerate projects based on the __VSENUMPROJFLAGS flags passed in. For
                //example if you want to only enumerate C# projects use EPF_MATCHTYPE and pass C# project guid. See
                //Common\IDL\vsshell.idl for more details.
                int hr = solution.GetProjectEnum( (uint) __VSENUMPROJFLAGS.EPF_ALLINSOLUTION, ref nullGuid, out penum );
                ErrorHandler.ThrowOnFailure( hr );
                if ( ( VSConstants.S_OK == hr ) && ( penum != null ) )
                {
                    OutputCommandString( "\n\nTraverse All Items in Flat Projects Collection:\n" );

                    uint fetched;
                    var rgelt = new IVsHierarchy[ 1 ];
                    while ( penum.Next( 1, rgelt, out fetched ) == 0 && fetched == 1 )
                    {
                        //Get the root hierarchy of each project so we can walk the tree.
                        this.EnumHierarchyItemsFlat( VSConstants.VSITEMID_ROOT, rgelt[ 0 ], 0, false, _processNode );
                    }
                }
            }
        }

        /// <summary>
        /// Enumerates over the hierarchy items for the given hierarchy.
        /// This enumeration does not recursively traverse into nested hierarchies.
        /// </summary>
        /// <param name="hierarchy">hierarchy to enmerate over.</param>
        /// <param name="itemid">item id of the hierarchy</param>
        /// <param name="recursionLevel">Depth of recursion. e.g. if recursion started with the Solution
        /// node, then : Level 0 -- Solution node, Level 1 -- children of Solution, etc.</param>
        /// <param name="visibleNodesOnly">true if only nodes visible in the Solution Explorer should
        /// be traversed. false if all project items should be traversed.</param>
        /// <param name="processNodeFunc">pointer to function that should be processed on each
        /// node as it is visited in the depth first enumeration.</param>
        private void EnumHierarchyItemsFlat( uint itemid, IVsHierarchy hierarchy, int recursionLevel, bool visibleNodesOnly, ProcessHierarchyNode processNodeFunc )
        {
            // Display name and type of the node in the Output Window
            processNodeFunc( hierarchy, itemid, recursionLevel );

            int hr;
            object pVar;

            recursionLevel++;

            //Get the first child node of the current hierarchy being walked
            hr = hierarchy.GetProperty( itemid,
                ( visibleNodesOnly ? (int) __VSHPROPID.VSHPROPID_FirstVisibleChild : (int) __VSHPROPID.VSHPROPID_FirstChild ),
                out pVar );
            ErrorHandler.ThrowOnFailure( hr );
            if ( VSConstants.S_OK == hr )
            {
                //We are using Depth first search so at each level we recurse to check if the node has any children
                // and then look for siblings.
                uint childId = GetItemId( pVar );
                while ( childId != VSConstants.VSITEMID_NIL )
                {
                    EnumHierarchyItemsFlat( childId, hierarchy, recursionLevel, visibleNodesOnly, processNodeFunc );
                    hr = hierarchy.GetProperty( childId,
                        ( visibleNodesOnly ? (int) __VSHPROPID.VSHPROPID_NextVisibleSibling : (int) __VSHPROPID.VSHPROPID_NextSibling ),
                        out pVar );
                    if ( VSConstants.S_OK == hr )
                    {
                        childId = GetItemId( pVar );
                    }
                    else
                    {
                        ErrorHandler.ThrowOnFailure( hr );
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the item id.
        /// </summary>
        /// <param name="pvar">VARIANT holding an itemid.</param>
        /// <returns>Item Id of the concerned node</returns>
        private uint GetItemId( object pvar )
        {
            if ( pvar == null ) return VSConstants.VSITEMID_NIL;
            if ( pvar is int ) return (uint) (int) pvar;
            if ( pvar is uint ) return (uint) pvar;
            if ( pvar is short ) return (uint) (short) pvar;
            if ( pvar is ushort ) return (ushort) pvar;
            if ( pvar is long ) return (uint) (long) pvar;
            return VSConstants.VSITEMID_NIL;
        }

        /// <summary>
        /// This function diplays the name of the Hierarchy node. This function is passed to the 
        /// Hierarchy enumeration routines to process the current node.
        /// </summary>
        /// <param name="hierarchy">Hierarchy of the current node</param>
        /// <param name="itemid">Itemid of the current node</param>
        /// <param name="recursionLevel">Depth of recursion in hierarchy enumeration. We add one tab
        /// for each level in the recursion.</param>
        private static void DisplayHierarchyNode( IVsHierarchy hierarchy, uint itemid, int recursionLevel )
        {
            object pVar;

            string text = "";

            for ( int i = 0; i < recursionLevel; i++ )
                text += "\t";

            var vsProject = hierarchy as IVsProject;

            if ( vsProject == null )
                return;


            //Get the name of the root node in question here and dump its value
            hierarchy.GetProperty( itemid, (int) __VSHPROPID.VSHPROPID_Name, out pVar );
            text += (string) pVar;
            OutputCommandString( text );
        }

        /// <summary>
        /// This functions prints on the debug ouput and on the generic pane of the output window
        /// a text.
        /// </summary>
        /// <param name="text">text to send to Output Window.</param>
        private static void OutputCommandString( string text )
        {
            // Build the string to write on the debugger and output window.
            var outputText = new StringBuilder( text );
            outputText.Append( "\n" );

            var outWindow = GetGlobalService( typeof( SVsOutputWindow ) ) as IVsOutputWindow;

            // Use e.g. Tools -> Create GUID to make a stable, but unique GUID for your pane.
            // Also, in a real project, this should probably be a static constant, and not a local variable
            var customGuid = new Guid( "C376C4E8-8E26-4D6F-886C-551A088EF57D" );
            const string customTitle = "Loaded Projects Output";
            outWindow.CreatePane( ref customGuid, customTitle, 1, 1 );

            IVsOutputWindowPane customPane;
            outWindow.GetPane( ref customGuid, out customPane );

            customPane.OutputString( outputText.ToString() );
            customPane.Activate(); // Brings this pane into view

            /*           // Now print the string on the output window.
                       // The first step is to get a reference to IVsOutputWindow.
                       IVsOutputWindow outputWindow = Package.GetGlobalService( typeof( SVsOutputWindow ) ) as IVsOutputWindow;

                       // If we fail to get it we can exit now.
                       if ( null == outputWindow )
                       {
                           Trace.WriteLine( "Failed to get a reference to IVsOutputWindow" );
                           return;
                       }

                       // Now get the window pane for the general output.
                       Guid guidGeneral = Microsoft.VisualStudio.VSConstants.GUID_OutWindowGeneralPane;
                       IVsOutputWindowPane windowPane;

                       // following instructions on MEF Output Window forum. if this doesn't work, use the commented out stuff below this
                       if ( Microsoft.VisualStudio.ErrorHandler.Failed( outputWindow.GetPane( ref guidGeneral, out windowPane ) ) && ( Microsoft.VisualStudio.ErrorHandler.Succeeded( outputWindow.CreatePane( ref guidGeneral, null, 1, 1 ) ) ) )
                       {
                           outputWindow.GetPane( ref guidGeneral, out windowPane );
                       }

                       if ( Microsoft.VisualStudio.ErrorHandler.Failed( outputWindow.GetPane( ref guidGeneral, out windowPane ) ) )
                       {
                           Trace.WriteLine( "Failed to get a reference to the Output Window General pane" );
                           return;
                       }

                       // following instructions on MEF Output Window forum. if this doesn't work, use the commented out stuff below this
                       if ( windowPane != null )
                       {
                           Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure( windowPane.Activate() );

                           if ( Microsoft.VisualStudio.ErrorHandler.Failed( windowPane.OutputString( outputText.ToString() ) ) )
                           {
                               Trace.WriteLine( "Failed to write on the output window" );
                           }

                       }
             */

        }






        public static EnvDTE.Project ToDteProject( IVsHierarchy hierarchy )
        {
            if ( hierarchy == null ) throw new ArgumentNullException( "hierarchy" );
            object prjObject = null;
            if ( hierarchy.GetProperty( 0xfffffffe, -2027, out prjObject ) >= 0 )
            {
                return (EnvDTE.Project) prjObject;
            }
            throw new ArgumentException( "Hierarchy is not a project." );
        }


        private const string SettingsKey = "Gardiner.LoadedProjects";
        private Settings _settings;

        protected override void OnLoadOptions( string key, Stream stream )
        {
            if ( key == SettingsKey )
            {
                try
                {
                    var formatter = new BinaryFormatter();
                    _settings = (Settings) formatter.Deserialize( stream );
                }
                catch ( Exception )
                {
                    _settings = new Settings();
                }
            }
            else
            {
                base.OnLoadOptions( key, stream );
            }
        }

        protected override void OnSaveOptions( string key, Stream stream )
        {
            if ( key == SettingsKey )
            {
                try
                {
                    if ( null != _settings )
                    {
                        var formatter = new BinaryFormatter();
                        formatter.Serialize( stream, _settings );
                    }
                }
                catch ( Exception ex )
                {
                    Debug.WriteLine( ex );
                }
            }
            else
            {
                base.OnSaveOptions( key, stream );
            }
        }

    }
}
