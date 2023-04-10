using System;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public class XAccessObject : IDisposable
    {
        public XAccessObject(MSAccess.AllObjects allObjects, int index)
        {
            accessObject = allObjects[index];

            Name = accessObject.Name;
        }

        public string Name { get; }

        private MSAccess.AccessObject accessObject;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (accessObject != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(accessObject);

                accessObject = null;
            }
        }

        ~XAccessObject()
        {
            Dispose(false);
        }
    }
}
