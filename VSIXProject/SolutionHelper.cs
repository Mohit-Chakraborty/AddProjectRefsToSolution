using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;

namespace VSIXProject
{
    internal static class SolutionHelper
    {
        internal static IEnumerable<string> GetPathsOfLoadedProjectsInSolution(IVsSolution solution)
        {
            Requires.NotNull(solution, nameof(solution));

            ThreadHelper.ThrowIfNotOnUIThread();

            var projectPaths = new List<string>();

            if (ErrorHandler.Succeeded(solution.GetSolutionInfo(out string solutionDir, out _, out _)) &&
                ErrorHandler.Succeeded(solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies)))
            {
                var hierarchy = new IVsHierarchy[1];

                while (ErrorHandler.Succeeded(enumHierarchies.Next(1, hierarchy, out uint countFetched)) && countFetched == 1)
                {
                    if (hierarchy[0] is IPersist hierPersist)
                    {
                        if (ErrorHandler.Succeeded(hierPersist.GetClassID(out Guid classID)) &&
                            (classID == VSConstants.CLSID.MiscellaneousFilesProject_guid || classID == VSConstants.CLSID.SolutionFolderProject_guid || classID == VSConstants.CLSID.SolutionItemsProject_guid))
                        {
                            continue;
                        }
                    }

                    int hr = solution.GetUniqueNameOfProject(hierarchy[0], out string projectUniqueName);

                    if (ErrorHandler.Succeeded(hr))
                    {
                        string projectPath = Path.Combine(solutionDir, projectUniqueName);
                        projectPaths.Add(projectPath);
                    }
                }
            }

            return projectPaths;
        }

        internal static int AddProjectToSolution(IVsSolution solution, string projectPath)
        {
            Requires.NotNull(solution, nameof(solution));

            ThreadHelper.ThrowIfNotOnUIThread();

            int hr = VSConstants.S_OK;
            if (ErrorHandler.Succeeded(solution.GetProjectOfUniqueName(projectPath, out _)))
            {
            }
            else
            {
                PackageHelper.WriteMessage("Adding project to solution: " + projectPath);
                hr = ((IVsSolution6)solution).AddExistingProject(projectPath, null, out _);

                if (ErrorHandler.Failed(hr))
                {
                    PackageHelper.WriteMessage("FAILED. hr = " + hr);
                }
            }

            return hr;
        }
    }
}
