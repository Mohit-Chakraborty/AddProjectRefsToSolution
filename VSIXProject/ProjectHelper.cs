using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace VSIXProject
{
    internal static class ProjectHelper
    {
        /// <summary>
        /// Reads the ProjectReference elements in the project XML
        /// </summary>
        /// <param name="project"><code>KeyValuePair</code> containing the full path to the project file and the project hierarchy</param>
        /// <returns>Collection of project reference paths</returns>
        internal static List<string> GetProjectReferencePaths(KeyValuePair<string, IVsHierarchy> project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var projectReferencePaths = new List<string>();

            using (var fileStream = File.OpenRead(project.Key))
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
                                    PackageHelper.WriteMessage(System.Environment.NewLine + project.Key + "--->" + includeValue);

                                    var referencedProjectPath = ResolveMacrosInPath(project.Value, includeValue);

                                    if (!Path.IsPathRooted(referencedProjectPath))
                                    {
                                        referencedProjectPath = Path.Combine(Path.GetDirectoryName(project.Key), referencedProjectPath);
                                    }

                                    // Save the canonical path of the project.
                                    projectReferencePaths.Add(new FileInfo(referencedProjectPath).FullName);
                                }
                            }
                        }
                        while (xmlReader.ReadToNextSibling("ProjectReference"));
                    }
                }
            }

            return projectReferencePaths;
        }

        internal static string ResolveMacrosInPath(IVsHierarchy projectHierarchy, string path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if ((projectHierarchy == null) || string.IsNullOrEmpty(path))
            {
                return path;
            }

            int startIndex = 0;
            var i = path.IndexOf("$(", startIndex);

            if (i < 0)
            {
                // Path has no macros
                return path;
            }

            string resolvedPath = string.Empty;

            while (i >= 0)
            {
                var j = path.IndexOf(")", i);
                resolvedPath += path.Substring(startIndex, i - startIndex);

                var propertyName = path.Substring(i + 2, j - (i + 2));
                var propertyValue = GetProjectPropertyValue(projectHierarchy, propertyName);

                PackageHelper.WriteMessage(propertyName + " = " + propertyValue);

                resolvedPath += propertyValue;

                startIndex = j + 1;
                i = path.IndexOf("$(", startIndex);
            }

            if (startIndex < path.Length)
            {
                resolvedPath += path.Substring(startIndex, path.Length - startIndex);
            }

            return resolvedPath;
        }

        /// <summary>
        /// Gets the value of a property defined in the passed-in project
        /// </summary>
        /// <param name="projectHierarchy">Project</param>
        /// <param name="propertyName">Name of the property</param>
        /// <returns>Value of the property. If there is an error, returns <code>string.Empty</code></returns>
        private static string GetProjectPropertyValue(IVsHierarchy projectHierarchy, string propertyName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (projectHierarchy is IVsBuildPropertyStorage vsBuildPropertyStorage)
            {
                int hr = vsBuildPropertyStorage.GetPropertyValue(propertyName, null, (uint)_PersistStorageType.PST_PROJECT_FILE, out string propertyValue);

                if (ErrorHandler.Succeeded(hr))
                {
                    return propertyValue;
                }
            }

            PackageHelper.WriteMessage("ERROR reading project property - " + propertyName ?? "<null>");
            return string.Empty;
        }
    }
}
