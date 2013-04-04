using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace DavidGardiner.Gardiner_LoadedProjects
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
    public sealed class Gardiner_LoadedProjectsPackage : Package
        // ReSharper restore InconsistentNaming
    {
        private const string SettingsKey = "Gardiner.LoadedProjects";
        private const string OutputWindowId = "C376C4E8-8E26-4D6F-886C-551A088EF57D";
        private static IVsOutputWindowPane _customPane;
        private IList<HierarchyPathPair> _loaded;
        private Settings _settings;
        private IList<HierarchyPathPair> _unloaded;

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

            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                var menuCommandID = new CommandID(GuidList.guidGardiner_LoadedProjectsCmdSet,
                                                  (int) PkgCmdIDList.cmdidLoadedProjects);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            _settings = new Settings();

            PrepareOutput();
        }

        /// <summary>
        ///     This function is the callback used to execute a command when the a menu item is clicked.
        ///     See the Initialize method to see how the menu item is associated to this function using
        ///     the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            GetProjects();

            using (var frm = new frmProfiles())
            {
                frm.Unloaded = _unloaded.Select(x => x.HierarchyPath).ToList();
                frm.Settings = _settings;

                if (frm.ShowDialog() != DialogResult.Cancel)
                {
                    if (frm.SelectedProfile != null)
                    {
                        var profile = frm.SelectedProfile;

                        IVsUIHierarchyWindow slnExpHierWin = VsShellUtilities.GetUIHierarchyWindow(this, VSConstants.StandardToolWindows.SolutionExplorer);

                        var solutionService = (IVsSolution) GetService(typeof(SVsSolution));

                        foreach (var project in profile.UnloadedProjects)
                        {
                            var item = _loaded.FirstOrDefault(x => x.HierarchyPath == project);

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
            else
            {
                Debug.WriteLine("Failed to get full path for {0}", pVar);

                return string.Empty;
            }
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


        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        protected override void OnLoadOptions(string key, Stream stream)
        {
            if (key == SettingsKey)
            {
                try
                {
                    var formatter = new DataContractSerializer(typeof(Settings));
                    _settings = (Settings) formatter.ReadObject(stream);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Exception loading options. {0}", ex);
                    _settings = new Settings();
                }
            }
            else
            {
                base.OnLoadOptions(key, stream);
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        protected override void OnSaveOptions(string key, Stream stream)
        {
            if (key == SettingsKey)
            {
                try
                {
                    if (null != _settings)
                    {
                        var formatter = new DataContractSerializer(typeof(Settings));
                        formatter.WriteObject(stream, _settings);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Exception saving options. {0}", ex);
                }
            }
            else
            {
                base.OnSaveOptions(key, stream);
            }
        }
    }

    internal sealed class AllowAllAssemblyVersionsDeserializationBinder : System.Runtime.Serialization.SerializationBinder
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