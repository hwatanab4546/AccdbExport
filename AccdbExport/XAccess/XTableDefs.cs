using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public class XTableDefs : IEnumerable<XTableDef>, IDisposable
    {
        public XTableDefs(MSAccess.Dao.Database database)
        {
            tabledefs = database.TableDefs;
        }

        public IEnumerator<XTableDef> GetEnumerator()
        {
            if (enumerator == null)
            {
                enumerator = new TableDefEnumerator(tabledefs);
            }
            return enumerator;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private MSAccess.Dao.TableDefs tabledefs = null;
        private TableDefEnumerator enumerator = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (tabledefs != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    enumerator?.Dispose();
                    enumerator = null;
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(tabledefs);

                tabledefs = null;
            }
        }

        ~XTableDefs()
        {
            Dispose(false);
        }

        private class TableDefEnumerator : IEnumerator<XTableDef>
        {
            public TableDefEnumerator(MSAccess.Dao.TableDefs tabledefs)
            {
                this.tabledefs = tabledefs;
            }

            public XTableDef Current
            {
                get
                {
                    _Current?.Dispose();
                    _Current = new XTableDef(tabledefs, index);
                    return _Current;
                }
            }

            object IEnumerator.Current => Current;

            public bool MoveNext()
            {
                if (index < tabledefs.Count - 1)
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

            private MSAccess.Dao.TableDefs tabledefs = null;
            private XTableDef _Current = null;
            private int index = -1;

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            private void Dispose(bool disposing)
            {
                if (tabledefs != null)
                {
                    if (disposing)
                    {
                        // Free managed resources here.
                        _Current?.Dispose();
                        _Current = null;
                    }

                    // Free unmanaged resources here.

                    tabledefs = null;
                }
            }

            ~TableDefEnumerator()
            {
                Dispose(false);
            }
        }
    }
}
