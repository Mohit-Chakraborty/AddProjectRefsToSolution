using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSIXProject
{
    internal sealed class SolutionEventsListener : IVsSolutionEvents
    {
        internal uint Cookie { private get; set; }

        private readonly IVsSolution solution;
        private readonly VSIXProjectPackage projectPackage;

        internal SolutionEventsListener(IVsSolution solution, VSIXProjectPackage projectPackage)
        {
            Requires.NotNull(solution, nameof(solution));
            Requires.NotNull(projectPackage, nameof(projectPackage));

            this.solution = solution;
            this.projectPackage = projectPackage;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (fNewSolution == 1)
            {
                // this.solution.AdviseSolutionEvents(this, out uint cookie);
                // this.Cookie = cookie;

                this.projectPackage.OnSolutionOpen(this.solution);
            }

            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (this.Cookie != 0)
            {
                // this.solution.UnadviseSolutionEvents(this.Cookie);
                // this.Cookie = 0;
            }

            return VSConstants.S_OK;
        }
    }
}
