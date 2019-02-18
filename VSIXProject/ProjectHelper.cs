using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace VSIXProject
{
    internal static class ProjectHelper
    {
        internal static void CollectProjectReferencePaths(string projectPath, List<string> allReferencedProjectPaths)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<string> referencedProjectPaths = new List<string>();

            try
            {
                referencedProjectPaths = GetProjectReferencePaths(projectPath);
            }
            catch (Exception e)
            {
                PackageHelper.WriteMessage("*** FAILED to read project references. ***\t" + e.Message);
            }

            foreach (var referencedProjectPath in referencedProjectPaths)
            {
                PackageHelper.WriteMessage(projectPath + "->" + referencedProjectPath);

                if (!allReferencedProjectPaths.Contains(referencedProjectPath))
                {
                    allReferencedProjectPaths.Add(referencedProjectPath);

                    CollectProjectReferencePaths(referencedProjectPath, allReferencedProjectPaths);
                }
            }
        }

        private static List<string> GetProjectReferencePaths(string projectPath)
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
    }
}
