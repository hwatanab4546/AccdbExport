using System;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public class XQueryDef : IDisposable
    {
        public XQueryDef(MSAccess.Dao.QueryDefs querydefs, int index)
        {
            querydef = querydefs[index];

            Name = querydef.Name;
            SQL = querydef.SQL;
        }

        public string Name { get; }
        public string SQL { get; }

        private MSAccess.Dao.QueryDef querydef = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (querydef != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(querydef);

                querydef = null;
            }
        }

        ~XQueryDef()
        {
            Dispose(false);
        }
    }
}
