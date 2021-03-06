﻿using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
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

#pragma warning disable IDE0052 // Remove unread private members
        private PackageHelper packageHelper;
#pragma warning restore IDE0052 // Remove unread private members

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

            this.packageHelper = new PackageHelper(this);

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
                PackageHelper.WriteMessage("*** FAILED to advise to solution events. ***");
            }

            this.OnSolutionOpen(solution);
        }

        #endregion

        internal void OnSolutionOpen(IVsSolution solution)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var loadedProjects = SolutionHelper.GetLoadedProjectsInSolution(solution);

            foreach (var loadedProject in loadedProjects)
            {
                AddProjectReferencesToSolution(solution, loadedProject);
            }
        }

        internal void AddProjectReferencesToSolution(IVsSolution solution, KeyValuePair<string, IVsHierarchy> project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var referencedProjectPaths = new List<string>();

            try
            {
                referencedProjectPaths = ProjectHelper.GetProjectReferencePaths(project);
            }
            catch (Exception e)
            {
                PackageHelper.WriteMessage("*** FAILED to read project references. ***\t" + e.Message);
            }

            foreach (string referencedProjectPath in referencedProjectPaths)
            {
                PackageHelper.WriteMessage(string.Empty);

                try
                {
                    var newProjectHierarchy = SolutionHelper.AddProjectToSolution(solution, referencedProjectPath);

                    if (newProjectHierarchy != null)
                    {
                        AddProjectReferencesToSolution(solution, new KeyValuePair<string, IVsHierarchy>(referencedProjectPath, newProjectHierarchy));
                    }
                }
                catch (Exception e)
                {
                    PackageHelper.WriteMessage("*** FAILED to resolve project properties. ***\t" +  e.Message);
                }
            }

            SharedProjectHelper.AddSharedProjectReferencesToSolution(solution, project.Key);
        }
    }
}

