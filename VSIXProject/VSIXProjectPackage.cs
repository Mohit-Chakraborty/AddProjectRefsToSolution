using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(VSIXProjectPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class VSIXProjectPackage : AsyncPackage
    {
        /// <summary>
        /// VSIXProjectPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "de32f89a-27d0-43ff-992f-9070de6e3f72";

        private SolutionEventsListener solutionEventsListener;

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var solution = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            Assumes.Present(solution);

            this.solutionEventsListener = new SolutionEventsListener(solution, this);
            int hr = solution.AdviseSolutionEvents(this.solutionEventsListener, out uint cookie);

            if (ErrorHandler.Succeeded(hr))
            {
                this.solutionEventsListener.Cookie = cookie;
            }
            else
            {
                WriteMessage("*** FAILED to advise to solution events. ***");
            }

            this.OnSolutionOpen(solution);
        }

        #endregion

        internal void OnSolutionOpen(IVsSolution solution)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var loadedProjects = this.GetPathsOfLoadedProjectsInSolution(solution);

            // Collect the full set of transitive project references of all the loaded projects
            List<string> allReferencedProjectPaths = new List<string>();

            foreach (string loadedProjectPath in loadedProjects)
            {
                CollectProjectReferencePaths(loadedProjectPath, allReferencedProjectPaths);
            }

            foreach (var referencedProjectPath in allReferencedProjectPaths)
            {
                WriteMessage("Adding project to solution: " + referencedProjectPath);
                int hr = ((IVsSolution6)solution).AddExistingProject(referencedProjectPath, null, out IVsHierarchy addedProjectHierarchy);

                if (ErrorHandler.Failed(hr))
                {
                    WriteMessage("FAILED. hr = " + hr);
                }
            }
        }

        private IEnumerable<string> GetPathsOfLoadedProjectsInSolution(IVsSolution solution)
        {
            Requires.NotNull(solution, nameof(solution));

            ThreadHelper.ThrowIfNotOnUIThread();

            ErrorHandler.ThrowOnFailure(solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies));

            var projectPaths = new List<string>();

            var hierarchy = new IVsHierarchy[1];
            int hr = enumHierarchies.Next(1, hierarchy, out uint countFetched);

            while (ErrorHandler.Succeeded(hr) && countFetched == 1)
            {
                int hr1 = solution.GetSolutionInfo(out string solutionDir, out _, out _);
                int hr2 = solution.GetUniqueNameOfProject(hierarchy[0], out string projectUniqueName);

                string projectPath = Path.Combine((string)solutionDir, (string)projectUniqueName);

                if (ErrorHandler.Succeeded(hr1) && ErrorHandler.Succeeded(hr2))
                {
                    projectPaths.Add(projectPath);
                }

                hr = enumHierarchies.Next(1, hierarchy, out countFetched);
            }

            return projectPaths;
        }

        private void CollectProjectReferencePaths(string projectPath, List<string> allReferencedProjectPaths)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<string> referencedProjectPaths = new List<string>();

            try
            {
                referencedProjectPaths = GetProjectReferencePaths(projectPath);
            }
            catch (Exception e)
            {
                WriteMessage("*** FAILED to read project references. ***\t" + e.Message);
            }

            foreach (var referencedProjectPath in referencedProjectPaths)
            {
                WriteMessage(projectPath + "->" + referencedProjectPath);

                if (!allReferencedProjectPaths.Contains(referencedProjectPath))
                {
                    allReferencedProjectPaths.Add(referencedProjectPath);

                    CollectProjectReferencePaths(referencedProjectPath, allReferencedProjectPaths);
                }
            }
        }

        private List<string> GetProjectReferencePaths(string projectPath)
        {
            var projectReferencePaths = new List<string>();

            using (var fileStream = File.OpenRead(projectPath))
            {
                using (XmlReader xmlReader = XmlReader.Create(fileStream))
                {
                    if (xmlReader.ReadToDescendant("ProjectReference"))
                    {
                        do
                        {
                            if (xmlReader.MoveToAttribute("Include"))
                            {
                                var includeValue = xmlReader.ReadContentAsString();

                                if (!string.IsNullOrEmpty(includeValue))
                                {
                                    var projectFullPath = Path.Combine(Path.GetDirectoryName(projectPath), includeValue);

                                    // Save the canonical path of the project.
                                    projectReferencePaths.Add(new FileInfo(projectFullPath).FullName);
                                }
                            }
                        }
                        while (xmlReader.ReadToNextSibling("ProjectReference"));
                    }
                }
            }

            return projectReferencePaths;
        }

        private void WriteMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var generalOutputWindowPane = this.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, "Add referenced projects to solution");

            generalOutputWindowPane?.Activate();
            generalOutputWindowPane?.OutputString(message + System.Environment.NewLine);
        }
    }
}
