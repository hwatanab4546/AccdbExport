using System;
using System.Collections;
using System.Collections.Generic;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    class XAccessObjectEnumerator : IEnumerator<XAccessObject>
    {
        public XAccessObjectEnumerator(MSAccess.AllObjects allObjects)
        {
            this.allObjects = allObjects;
        }

        public XAccessObject Current
        {
            get
            {
                _Current?.Dispose();
                _Current = new XAccessObject(allObjects, index);
                return _Current;
            }
        }
        object IEnumerator.Current => Current;

        public bool MoveNext()
        {
            if (index < allObjects.Count - 1)
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

        private MSAccess.AllObjects allObjects = null;
        private XAccessObject _Current = null;
        private int index = -1;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (allObjects != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    _Current?.Dispose();
                    _Current = null;
                }

                // Free unmanaged resources here.

                allObjects = null;
            }
        }

        ~XAccessObjectEnumerator()
        {
            Dispose(false);
        }
    }
}
