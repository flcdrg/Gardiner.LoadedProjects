using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Gardiner.LoadedProjects
{
    public partial class frmProgress : Form
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly IList<HierarchyPathPair> _loaded;
        private readonly IList<HierarchyPathPair> _unloaded;
        private readonly IList<string> _unloadedProjects;
        private IVsOutputWindowPane _customPane;

        public frmProgress()
        {
            InitializeComponent();
        }

        public frmProgress(IServiceProvider serviceProvider, IList<HierarchyPathPair> loaded, IList<HierarchyPathPair> unloaded, IList<string> unloadedProjects) : this()
        {
            _serviceProvider = serviceProvider;
            _loaded = loaded;
            _unloaded = unloaded;
            _unloadedProjects = unloadedProjects;
        }

        private void PrepareOutput()
        {
            var outWindow = (IVsOutputWindow) _serviceProvider.GetService(typeof(SVsOutputWindow));

            var customGuid = new Guid(GuidList.OutputWindowId);

            const string customTitle = "Loaded Projects Output";
            ErrorHandler.ThrowOnFailure(outWindow.CreatePane(ref customGuid, customTitle, 1, 1));

            ErrorHandler.ThrowOnFailure(outWindow.GetPane(ref customGuid, out _customPane));

            ErrorHandler.ThrowOnFailure(_customPane.Activate()); // Brings this pane into view
        }

        private void OutputCommandString(string text)
        {
            ErrorHandler.ThrowOnFailure(_customPane.OutputString(text + "\n"));
        }

        private void frmProgress_Shown(object sender, EventArgs e)
        {
            Application.DoEvents();
            System.Threading.Thread.Sleep(0);

            PrepareOutput();

            var projectsToUnload = _unloadedProjects
                //.Select(project => Path.GetFullPath(Path.Combine(_solutionFolder, project)))
                .Select(project => _loaded.FirstOrDefault(x => x.HierarchyPath.Equals(project, StringComparison.CurrentCultureIgnoreCase)))
                .Where(item => item != null)
                .ToList();

            var projectsToReload = _unloaded
                .Where(pair => _unloadedProjects.All(x => x != pair.HierarchyPath))
                .ToList();

            var dte = (DTE)_serviceProvider.GetService(typeof(SDTE));

            IVsUIHierarchyWindow slnExpHierWin = VsShellUtilities.GetUIHierarchyWindow(_serviceProvider, VSConstants.StandardToolWindows.SolutionExplorer);

            if (slnExpHierWin == null)
            {
                OutputCommandString(string.Format(CultureInfo.CurrentCulture, "Failed to get the solution explorer hierarchy window"));
                return;
            }

            var solutionService = (IVsSolution)_serviceProvider.GetService(typeof(SVsSolution));

            int count = projectsToUnload.Count;

            if (count > 0)
            {
                int increment = 100 / count;

                for (int i = 0; i < count; i++)
                {
                    var item = projectsToUnload[i];

                    progressBar1.Increment(increment);
                    lblCurrent.Text = item.HierarchyPath;
                    lblAction.Text = string.Format( "Unloading {0} of {1}", i, count );

                    Application.DoEvents();

                    ErrorHandler.ThrowOnFailure(solutionService.CloseSolutionElement((uint)__VSSLNCLOSEOPTIONS.SLNCLOSEOPT_UnloadProject, item.Hierarchy, 0));

                    OutputCommandString(string.Format(CultureInfo.CurrentCulture, "Unloaded {0}", item.HierarchyPath));
                }
            }

            count = projectsToReload.Count;

            if (count > 0)
            {
                int increment = 100 / count;

                progressBar1.Value = 0;

                for (int i = 0; i < count; i++)
                {
                    var item = projectsToReload[i];

                    progressBar1.Increment(increment);
                    lblCurrent.Text = item.HierarchyPath;
                    lblAction.Text = string.Format( "Reloading {0} of {1}", i, count );

                    ErrorHandler.ThrowOnFailure(slnExpHierWin.ExpandItem(item.Hierarchy, (uint)VSConstants.VSITEMID.Root, EXPANDFLAGS.EXPF_SelectItem));

                    dte.ExecuteCommand("Project.ReloadProject");

                    OutputCommandString(string.Format(CultureInfo.CurrentCulture, "Reloaded {0}", item.HierarchyPath));

                    Application.DoEvents();
                }
            }

            Close();
        }

    }
}
