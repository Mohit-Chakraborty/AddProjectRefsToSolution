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
        internal static IDictionary<string, IVsHierarchy> GetLoadedProjectsInSolution(IVsSolution solution)
        {
            Requires.NotNull(solution, nameof(solution));

            ThreadHelper.ThrowIfNotOnUIThread();

            var loadedProjects = new Dictionary<string, IVsHierarchy>();

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
                        loadedProjects.Add(projectPath, hierarchy[0]);
                    }
                }
            }

            return loadedProjects;
        }

        internal static IVsHierarchy AddProjectToSolution(IVsSolution solution, string projectPath)
        {
            Requires.NotNull(solution, nameof(solution));
            Requires.NotNullOrEmpty(projectPath, nameof(projectPath));

            ThreadHelper.ThrowIfNotOnUIThread();

            if (ErrorHandler.Succeeded(solution.GetProjectOfUniqueName(projectPath, out IVsHierarchy existingProject)))
            {
                return existingProject;
            }
            else
            {
                PackageHelper.WriteMessage("Adding project to solution: " + projectPath);
                int hr = ((IVsSolution6)solution).AddExistingProject(projectPath, null, out IVsHierarchy newProject);

                if (ErrorHandler.Succeeded(hr))
                {
                    return newProject;
                }
                else
                {
                    PackageHelper.WriteMessage("FAILED. hr = " + hr);
                    return null;
                }
            }
        }
    }
}
