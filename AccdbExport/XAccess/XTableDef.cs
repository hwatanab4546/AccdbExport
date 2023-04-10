using System;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public class XTableDef : IDisposable
    {
        public XTableDef(MSAccess.Dao.TableDefs tabledefs, int index)
        {
            tabledef = tabledefs[index];

            Name = tabledef.Name;
        }

        public string Name { get; }

        private MSAccess.Dao.TableDef tabledef = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (tabledef != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(tabledef);

                tabledef = null;
            }
        }

        ~XTableDef()
        {
            Dispose(false);
        }
    }
}
