using System;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public sealed class XDatabase : IDisposable
    {
        public XDatabase(MSAccess.Application application)
        {
            database = application.CurrentDb();
        }

        private XTableDefs _TableDefs = null;
        public XTableDefs TableDefs
        {
            get
            {
                if (_TableDefs == null)
                {
                    _TableDefs = new XTableDefs(database);
                }
                return _TableDefs;
            }
        }

        private XQueryDefs _QueryDefs = null;
        public XQueryDefs QueryDefs
        {
            get
            {
                if (_QueryDefs == null)
                {
                    _QueryDefs = new XQueryDefs(database);
                }
                return _QueryDefs;
            }
        }

        private MSAccess.Dao.Database database = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (database != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    _TableDefs?.Dispose();
                    _TableDefs = null;

                    _QueryDefs?.Dispose();
                    _QueryDefs = null;
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(database);

                database = null;
            }
        }

        ~XDatabase()
        {
            Dispose(false);
        }
    }
}
