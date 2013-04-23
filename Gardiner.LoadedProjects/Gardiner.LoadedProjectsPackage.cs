using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Gardiner.LoadedProjects
{
    /// <summary>
    ///     This is the class that implements the package exposed by this assembly.
    ///     The minimum requirement for a class to be considered a valid package for Visual Studio
    ///     is to implement the IVsPackage interface and register itself with the shell.
    ///     This package uses the helper classes defined inside the Managed Package Framework (MPF)
    ///     to do it: it derives from the Package class that provides the implementation of the
    ///     IVsPackage interface and uses the registration attributes defined in the framework to
    ///     register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidGardiner_LoadedProjectsPkgString)]
    // ReSharper disable InconsistentNaming
    public sealed class Gardiner_LoadedProjectsPackage : Package, IVsSolutionEvents
        // ReSharper restore InconsistentNaming
    {
        private const int FILE_ATTRIBUTE_DIRECTORY = 0x10;
        private const int FILE_ATTRIBUTE_NORMAL = 0x80;
        private const string FileSuffix = ".LoadedProjects.User";

        private const string SettingsKey = "Gardiner.LoadedProjects";
        private const string OutputWindowId = "C376C4E8-8E26-4D6F-886C-551A088EF57D";
        private static IVsOutputWindowPane _customPane;
        private IList<HierarchyPathPair> _loaded;
        private Settings _settings;
        private IList<HierarchyPathPair> _unloaded;
        private DTE _dte;
        private uint _solutionCookie;
        private bool _settingsModified;

        /// <summary>
        ///     Default constructor of the package.
        ///     Inside this method you can place any initialization code that does not require
        ///     any Visual Studio service because at this point the package object is created but
        ///     not sited yet inside Visual Studio environment. The place to do all the other
        ///     initialization is the Initialize method.
        /// </summary>
        public Gardiner_LoadedProjectsPackage()
        {
            Debug.WriteLine("Entering constructor for: {0}", ToString());
            AddOptionKey(SettingsKey);
        }


        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation

        /// <summary>
        ///     Initialization of the package; this method is called right after the package is sited, so this is the place
        ///     where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine("Entering Initialize() of: {0}", ToString());
            base.Initialize();

            _dte = GetService(typeof(SDTE)) as DTE;

            // listen for solution events
            var solution = (IVsSolution)GetService(typeof(SVsSolution));
            ErrorHandler.ThrowOnFailure(solution.AdviseSolutionEvents(this, out _solutionCookie));
   
            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                var menuCommandID = new CommandID(GuidList.guidGardiner_LoadedProjectsCmdSet, (int) PkgCmdIDList.cmdidLoadedProjects);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            PrepareOutput();
        }

        private void SaveSettings( bool closing )
        {
            string solutionPath = _dte.Solution.FullName + FileSuffix;

            try
            {
                if (null != _settings)
                {
                    using (var fs = new FileStream(solutionPath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        var serializer = new XmlSerializer(typeof(Settings));
                        serializer.Serialize(fs, _settings);
                    }
                    OutputCommandString(string.Format("Saved settings to {0}", solutionPath));
                }
            }
            catch (Exception ex)
            {
                OutputCommandString(string.Format("Exception saving options. {0}", ex));
            }
            finally
            {
                _settingsModified = false;

                if ( closing )
                    _settings = null;
            }

        }

        private void SolutionOpened()
        {
            if (_settings == null)
            {
                string solutionPath = _dte.Solution.FullName + FileSuffix;

                if (File.Exists(solutionPath))
                {
                    try
                    {
                        using (var fs = new FileStream(solutionPath, FileMode.Open))
                        {
                            var serializer = new XmlSerializer(typeof(Settings));
                            _settings = (Settings)serializer.Deserialize(fs);
                        }

                        OutputCommandString("Loaded options");
                    }
                    catch (Exception ex)
                    {
                        OutputCommandString(string.Format("Exception loading options. {0}", ex));
                        _settings = new Settings();
                    }


                }
                else
                    _settings = new Settings();
            }

            _settings.PropertyChanged += SettingsOnPropertyChanged;

        }

        /// <summary>
        ///     This function is the callback used to execute a command when the a menu item is clicked.
        ///     See the Initialize method to see how the menu item is associated to this function using
        ///     the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            SolutionOpened();

            GetProjects();

            using (var frm = new frmProfiles())
            {

                string solutionFolder = Path.GetDirectoryName(_dte.Solution.FullName);

                if (string.IsNullOrEmpty(solutionFolder))
                {
                    OutputCommandString("Could not load path for solution");
                    return;
                }

                _settingsModified = false;
                frm.Unloaded = _unloaded.Select(x => GetRelativePath(solutionFolder, x.HierarchyPath)).ToList();
                frm.Settings = _settings;

                var dialogResult = frm.ShowDialog();

                if (dialogResult != DialogResult.Cancel)
                {
                    if ( _settingsModified )
                        SaveSettings( false );

                    if (frm.SelectedProfile != null)
                    {
                        var profile = frm.SelectedProfile;

                        IVsUIHierarchyWindow slnExpHierWin = VsShellUtilities.GetUIHierarchyWindow(this, VSConstants.StandardToolWindows.SolutionExplorer);

                        var solutionService = (IVsSolution) GetService(typeof(SVsSolution));

                        foreach (var project in profile.UnloadedProjects)
                        {
                            // convert to absolute path
                            string absoluteProjectPath = Path.GetFullPath(Path.Combine(solutionFolder, project));
                            var item = _loaded.FirstOrDefault(x => x.HierarchyPath.Equals(absoluteProjectPath, StringComparison.CurrentCultureIgnoreCase));

                            if (item != null)
                            {
                                if (slnExpHierWin == null)
                                {
                                    Debug.Fail("Failed to get the solution explorer hierarchy window!");
                                }
                                else
                                {
                                    ErrorHandler.ThrowOnFailure(
                                        solutionService.CloseSolutionElement(
                                            (uint) __VSSLNCLOSEOPTIONS.SLNCLOSEOPT_UnloadProject, item.Hierarchy, 0));

                                    OutputCommandString(string.Format(CultureInfo.CurrentCulture, "Unloaded {0}", item.HierarchyPath));
                                }
                            }
                        }

                        var dte = (DTE) GetService(typeof(DTE));

                        foreach (var pair in _unloaded)
                        {
                            if (profile.UnloadedProjects.All(x => x != pair.HierarchyPath))
                            {
                                ErrorHandler.ThrowOnFailure(slnExpHierWin.ExpandItem(pair.Hierarchy, (uint) VSConstants.VSITEMID.Root, EXPANDFLAGS.EXPF_SelectItem));

                                dte.ExecuteCommand("Project.ReloadProject");

                                OutputCommandString(string.Format(CultureInfo.CurrentCulture, "Reloaded {0}", pair.HierarchyPath));
                            }
                        }
                    }
                }
            }
        }

        private void SettingsOnPropertyChanged( object sender, PropertyChangedEventArgs propertyChangedEventArgs )
        {
             _settingsModified = true;
        }


        private static string GetFullPathToItem(IVsHierarchy hier)
        {
            // Most hierarchies in the solution will QI to IVsProject, which allows us to use GetMkDocument to get the full path to the project file,
            // stub hierarchies (i.e. unloaded hierarchies) do not implement IVsProject, luckily their IVsHierarchy::GetCanonicalName returns the
            // full path, so we fall back on that.
            string name;
            var proj = hier as IVsProject;

            object prjObject;
            if (hier.GetProperty(0xfffffffe, -2027, out prjObject) >= 0)
            {
                var p2 = (Project) prjObject;

                if (p2 != null && p2.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    return string.Empty;
            }


            object pVar;
            ErrorHandler.ThrowOnFailure(hier.GetProperty((uint) VSConstants.VSITEMID.Root, (int) __VSHPROPID.VSHPROPID_Name, out pVar));

            int hr;
            if (proj != null)
                hr = proj.GetMkDocument((uint) VSConstants.VSITEMID.Root, out name);
            else
                hr = hier.GetCanonicalName((uint) VSConstants.VSITEMID.Root, out name);

            if (hr == 0)
                return name;
            
            Debug.WriteLine("Failed to get full path for {0}", pVar);

            return string.Empty;
        }


        private void GetProjects()
        {
            _loaded = new List<HierarchyPathPair>();
            _unloaded = new List<HierarchyPathPair>();

            var enumerator = new ProjectEnumerator((IVsSolution) GetService(typeof(SVsSolution)));
            var loadedProjects =
                enumerator.LoadedProjects
                          .Select(x => new {hierarchy = (IVsUIHierarchy) x, path = GetFullPathToItem(x)})
                          .Where(x => x.path != null)
                          .Select(
                              h => new HierarchyPathPair(h.hierarchy, h.path));

            try
            {
                loadedProjects.ForEach(hpp =>
                {
                    if (_loaded != null)
                        _loaded.Add(hpp);
                });

                var unloadedProjects =
                    new List<HierarchyPathPair>(
                        enumerator.UnloadedProjects
                                  .Select(x => new {hierarchy = (IVsUIHierarchy) x, path = GetFullPathToItem(x)})
                                  .Where(x => x.path != null)
                                  .Select(
                                      h => new HierarchyPathPair(h.hierarchy, h.path)));


                unloadedProjects.ForEach(hpp => _unloaded.Add(hpp));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        }

        private static void PrepareOutput()
        {
            var outWindow = (IVsOutputWindow) GetGlobalService(typeof(SVsOutputWindow));

            var customGuid = new Guid(OutputWindowId);
            const string customTitle = "Loaded Projects Output";
            ErrorHandler.ThrowOnFailure(outWindow.CreatePane(ref customGuid, customTitle, 1, 1));

            ErrorHandler.ThrowOnFailure(outWindow.GetPane(ref customGuid, out _customPane));

            ErrorHandler.ThrowOnFailure(_customPane.Activate()); // Brings this pane into view
        }

        private static void OutputCommandString(string text)
        {
            if (_customPane == null)
                PrepareOutput();

            if ( _customPane != null )
                ErrorHandler.ThrowOnFailure(_customPane.OutputString(text + "\n"));
        }

        public static Project ToDteProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null)
                throw new ArgumentNullException("hierarchy");
            object prjObject;
            if (hierarchy.GetProperty(0xfffffffe, -2027, out prjObject) >= 0)
            {
                return (Project) prjObject;
            }
            throw new ArgumentException("Hierarchy is not a project.");
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            // Need to do this before solution closes, so we know the path of the solution
            _settings.PropertyChanged -= SettingsOnPropertyChanged;

            if (_settingsModified)
                SaveSettings(true);
            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public static string GetRelativePath(string fromPath, string toPath)
        {
            int fromAttr = GetPathAttribute(fromPath);
            int toAttr = GetPathAttribute(toPath);

            var path = new StringBuilder(260); // MAX_PATH
            if (NativeMethods.PathRelativePathTo(
                path,
                fromPath,
                fromAttr,
                toPath,
                toAttr) == 0)
            {
                throw new ArgumentException("Paths must have a common prefix");
            }
            return path.ToString();
        }

        private static int GetPathAttribute(string path)
        {
            var di = new DirectoryInfo(path);
            if (di.Exists)
            {
                return FILE_ATTRIBUTE_DIRECTORY;
            }

            var fi = new FileInfo(path);
            if (fi.Exists)
            {
                return FILE_ATTRIBUTE_NORMAL;
            }

            throw new FileNotFoundException();
        }
    }
}