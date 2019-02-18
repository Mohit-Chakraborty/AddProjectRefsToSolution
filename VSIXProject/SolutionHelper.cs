using Microsoft;
using Microsoft.VisualStudio;
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

            if (ErrorHandler.Succeeded(solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies)))
            {
                var hierarchy = new IVsHierarchy[1];

                while (ErrorHandler.Succeeded(enumHierarchies.Next(1, hierarchy, out uint countFetched)) && countFetched == 1)
                {
                    int hr1 = solution.GetSolutionInfo(out string solutionDir, out _, out _);
                    int hr2 = solution.GetUniqueNameOfProject(hierarchy[0], out string projectUniqueName);

                    if (ErrorHandler.Succeeded(hr1) && ErrorHandler.Succeeded(hr2))
                    {
                        string projectPath = Path.Combine((string)solutionDir, (string)projectUniqueName);
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
