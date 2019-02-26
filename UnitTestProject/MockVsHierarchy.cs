using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Validation;

namespace UnitTestProject
{
    internal sealed class MockVsHierarchy : IVsHierarchy, IVsBuildPropertyStorage
    {
        public int SetSite(Microsoft.VisualStudio.OLE.Interop.IServiceProvider psp)
        {
            return VSConstants.S_OK;
        }

        public int GetSite(out Microsoft.VisualStudio.OLE.Interop.IServiceProvider ppSP)
        {
            ppSP = null;
            return VSConstants.S_OK;
        }

        public int QueryClose(out int pfCanClose)
        {
            pfCanClose = 0;
            return VSConstants.S_OK;
        }

        public int Close()
        {
            return VSConstants.S_OK;
        }

        public int GetGuidProperty(uint itemid, int propid, out Guid pguid)
        {
            pguid = Guid.Empty;
            return VSConstants.S_OK;
        }

        public int SetGuidProperty(uint itemid, int propid, ref Guid rguid)
        {
            return VSConstants.S_OK;
        }

        public int GetProperty(uint itemid, int propid, out object pvar)
        {
            pvar = null;
            return VSConstants.S_OK;
        }

        public int SetProperty(uint itemid, int propid, object var)
        {
            return VSConstants.S_OK;
        }

        public int GetNestedHierarchy(uint itemid, ref Guid iidHierarchyNested, out IntPtr ppHierarchyNested, out uint pitemidNested)
        {
            ppHierarchyNested = IntPtr.Zero;
            pitemidNested = (uint)VSConstants.VSITEMID.Nil;
            return VSConstants.S_OK;
        }

        public int GetCanonicalName(uint itemid, out string pbstrName)
        {
            pbstrName = null;
            return VSConstants.S_OK;
        }

        public int ParseCanonicalName(string pszName, out uint pitemid)
        {
            pitemid = 0;
            return VSConstants.S_OK;
        }

        public int Unused0()
        {
            return VSConstants.S_OK;
        }

        public int AdviseHierarchyEvents(IVsHierarchyEvents pEventSink, out uint pdwCookie)
        {
            pdwCookie = 0;
            return VSConstants.S_OK;
        }

        public int UnadviseHierarchyEvents(uint dwCookie)
        {
            return VSConstants.S_OK;
        }

        public int Unused1()
        {
            return VSConstants.S_OK;
        }

        public int Unused2()
        {
            return VSConstants.S_OK;
        }

        public int Unused3()
        {
            return VSConstants.S_OK;
        }

        public int Unused4()
        {
            return VSConstants.S_OK;
        }

        public int GetPropertyValue(string pszPropName, string pszConfigName, uint storage, out string pbstrPropValue)
        {
            Requires.NotNull(pszPropName, nameof(pszPropName));

            pbstrPropValue = pszPropName;

            return VSConstants.S_OK;
        }

        public int SetPropertyValue(string pszPropName, string pszConfigName, uint storage, string pszPropValue)
        {
            return VSConstants.S_OK;
        }

        public int RemoveProperty(string pszPropName, string pszConfigName, uint storage)
        {
            return VSConstants.S_OK;
        }

        public int GetItemAttribute(uint item, string pszAttributeName, out string pbstrAttributeValue)
        {
            pbstrAttributeValue = null;
            return VSConstants.S_OK;
        }

        public int SetItemAttribute(uint item, string pszAttributeName, string pszAttributeValue)
        {
            return VSConstants.S_OK;
        }
    }
}
