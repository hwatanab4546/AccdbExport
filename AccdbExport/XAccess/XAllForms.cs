using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public sealed class XAllForms : IEnumerable<XAccessObject>, IDisposable
    {
        public XAllForms(MSAccess.Application application)
        {
            allForms = application.CurrentProject.AllForms;
        }

        public IEnumerator<XAccessObject> GetEnumerator()
        {
            if (enumerator == null)
            {
                enumerator = new XAccessObjectEnumerator(allForms);
            }
            return enumerator;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private MSAccess.AllObjects allForms = null;
        private XAccessObjectEnumerator enumerator = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (allForms != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    enumerator?.Dispose();
                    enumerator = null;
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(allForms);

                allForms = null;
            }
        }

        ~XAllForms()
        {
            Dispose(false);
        }
    }
}
