using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSIXProject
{
    internal sealed class PackageHelper
    {
        private static VSIXProjectPackage ProjectPackage;
        private static IVsOutputWindowPane GeneralOutputWindowPane;

        internal PackageHelper(VSIXProjectPackage projectPackage)
        {
            Requires.NotNull(projectPackage, nameof(projectPackage));

            ProjectPackage = projectPackage;
        }

        internal static void WriteMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (GeneralOutputWindowPane == null)
            {
                GeneralOutputWindowPane = ProjectPackage?.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, "Add referenced projects to solution");
            }

            GeneralOutputWindowPane?.Activate();
            GeneralOutputWindowPane?.OutputStringThreadSafe(message + System.Environment.NewLine);
        }
    }
}
