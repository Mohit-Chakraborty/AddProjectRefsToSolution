using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSIXProject
{
    internal sealed class PackageHelper
    {
        private static IVsOutputWindowPane GeneralOutputWindowPane;

        internal PackageHelper(VSIXProjectPackage projectPackage)
        {
            Requires.NotNull(projectPackage, nameof(projectPackage));
            GeneralOutputWindowPane = projectPackage.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, "Add referenced projects to solution");
        }

        internal static void WriteMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            GeneralOutputWindowPane?.Activate();
            GeneralOutputWindowPane?.OutputString(message + System.Environment.NewLine);
        }
    }
}
