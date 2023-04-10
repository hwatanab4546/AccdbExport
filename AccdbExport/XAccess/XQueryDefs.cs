using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public class XQueryDefs : IEnumerable<XQueryDef>, IDisposable
    {
        public XQueryDefs(MSAccess.Dao.Database database)
        {
            querydefs = database.QueryDefs;
        }

        public IEnumerator<XQueryDef> GetEnumerator()
        {
            if (enumerator == null)
            {
                enumerator = new QeyrtDefEnumerator(querydefs);
            }
            return enumerator;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private MSAccess.Dao.QueryDefs querydefs = null;
        private QeyrtDefEnumerator enumerator = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (querydefs != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    enumerator?.Dispose();
                    enumerator = null;
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(querydefs);

                querydefs = null;
            }
        }

        ~XQueryDefs()
        {
            Dispose(false);
        }

        private class QeyrtDefEnumerator : IEnumerator<XQueryDef>
        {
            public QeyrtDefEnumerator(MSAccess.Dao.QueryDefs querydefs)
            {
                this.querydefs = querydefs;
            }

            public XQueryDef Current
            {
                get
                {
                    _Current?.Dispose();
                    _Current = new XQueryDef(querydefs, index);
                    return _Current;
                }
            }

            object IEnumerator.Current => Current;

            public bool MoveNext()
            {
                if (index < querydefs.Count - 1)
                {
                    _Current?.Dispose();
                    _Current = null;

                    ++index;
                    return true;
                }
                else
                {
                    return false;
                }
            }

            public void Reset()
            {
                _Current?.Dispose();
                _Current = null;

                index = -1;
            }

            private MSAccess.Dao.QueryDefs querydefs = null;
            private XQueryDef _Current = null;
            private int index = -1;

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            private void Dispose(bool disposing)
            {
                if (querydefs != null)
                {
                    if (disposing)
                    {
                        // Free managed resources here.
                        _Current?.Dispose();
                        _Current = null;
                    }

                    // Free unmanaged resources here.

                    querydefs = null;
                }
            }

            ~QeyrtDefEnumerator()
            {
                Dispose(false);
            }
        }
    }
}
