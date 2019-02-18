using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSIXProject
{
    internal static class SharedProjectHelper
    {
        internal static void AddSharedProjectReferencesToSolution(IVsSolution solution, string projectPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Requires.NotNull(solution, nameof(solution));
            Requires.NotNull(projectPath, nameof(projectPath));

            // Get the shared project imports of the project.
            // Load all the Shared projects.
            var sharedProjectImports = GetSharedProjectImportPaths(solution, projectPath);

            foreach (var sharedProjectPath in sharedProjectImports)
            {
                if (ErrorHandler.Succeeded(solution.GetProjectOfUniqueName(sharedProjectPath, out IVsHierarchy sharedProjectHierarchy)))
                {
                    //if (ErrorHandler.Succeeded(sharedProjectHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID5.VSHPROPID_ProjectUnloadStatus, out object unloadStatus)))
                    //{
                    //    WriteMessage("Loading project in solution: " + sharedProjectPath);

                    //    int hr = projectHierarchy.GetGuidProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out Guid projectId);

                    //    var projectGuids = new Guid[1];
                    //    projectGuids[0] = projectId;

                    //    hr = ((IVsSolution8)solution).BatchProjectAction(
                    //        (uint)__VSBatchProjectAction.BPA_LOAD,
                    //        (uint)__VSBatchProjectActionFlags.BPAF_IGNORE_SELFRELOAD_PROJECTS,
                    //        1,
                    //        projectGuids,
                    //        out IVsBatchProjectActionContext actionContext);

                    //    if (ErrorHandler.Failed(hr))
                    //    {
                    //        WriteMessage("FAILED. hr = " + hr);
                    //    }
                    //}
                    //else
                    //{
                    // The Shared project is already loaded in the solution.
                    continue;
                    //}
                }
                else
                {
                    PackageHelper.WriteMessage("Adding project to solution: " + sharedProjectPath);
                    int hr = ((IVsSolution6)solution).AddExistingProject(sharedProjectPath, null, out IVsHierarchy addedProjectHierarchy);

                    if (ErrorHandler.Failed(hr))
                    {
                        PackageHelper.WriteMessage("FAILED. hr = " + hr);
                    }
                }
            }
        }

        private static List<string> GetSharedProjectImportPaths(IVsSolution solution, string projectUniqueName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Requires.NotNull(solution, nameof(solution));
            Requires.NotNull(projectUniqueName, nameof(projectUniqueName));

            // Assumption: The project is loaded.
            int hr = solution.GetProjectOfUniqueName(projectUniqueName, out IVsHierarchy projectHierarchy);

            if (ErrorHandler.Failed(hr) || (projectHierarchy == null))
            {
                throw new ArgumentException("Unknown project - " + projectUniqueName);
            }

            var sharedProjectImportPaths = new List<string>();

            // If the project is importing Shared projects, get the paths to the .projitems files using the 'VSHPROPID_SharedItemsImportFullPaths' property.
            if (ErrorHandler.Succeeded(projectHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID7.VSHPROPID_SharedItemsImportFullPaths, out object sharedItemImportsObject)) &&
                (sharedItemImportsObject is string sharedItemImports) &&
                !string.IsNullOrEmpty(sharedItemImports))
            {
                foreach (string sharedItemImportPath in sharedItemImports.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (ErrorHandler.Succeeded(projectHierarchy.ParseCanonicalName(sharedItemImportPath, out uint itemId)))
                    {
                        // If the Shared project is loaded, its hierarchy (represented by the .shproj file) can be obtained.
                        if (ErrorHandler.Succeeded(projectHierarchy.GetProperty(itemId, (int)__VSHPROPID7.VSHPROPID_SharedProjectHierarchy, out object sharedProjectHierarchyObject)) &&
                            sharedProjectHierarchyObject is IVsHierarchy sharedProjectHierarchy)
                        {
                            // The shared project is loaded, so we don't add it to the list.
                        }
                        else
                        {
                            PackageHelper.WriteMessage(projectUniqueName + "->" + sharedItemImportPath);

                            // If the Shared project is not loaded, the .shproj hierarchy is not available.
                            // The shared items import file (.projitems) is available only via .shproj hierarchy.
                            // Without that connection, we use the implementation detail that the name of the .shproj and its .projitems only varies by the file extension.
                            var sharedProjectPath = GetSharedProjectForImportsFile(new FileInfo(sharedItemImportPath).FullName);
                            sharedProjectImportPaths.Add(sharedProjectPath);
                        }
                    }
                }
            }

            return sharedProjectImportPaths;
        }

        private static string GetSharedProjectForImportsFile(string sharedItemImportPath)
        {
            var importsFileExtension = Path.GetExtension(sharedItemImportPath);

            if (".projitems".Equals(importsFileExtension, StringComparison.OrdinalIgnoreCase))
            {
                return Path.ChangeExtension(sharedItemImportPath, ".shproj");
            }

            return sharedItemImportPath;
        }
    }
}
