using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public sealed class XVBComponents : IEnumerable<XVBComponent>, IDisposable
    {
        public XVBComponents(MSAccess.Application application)
        {
            vBComponents = application.VBE.ActiveVBProject.VBComponents;

            Count = vBComponents.Count;
        }

        public int Count { get; }

        public IEnumerator<XVBComponent> GetEnumerator()
        {
            if (enumerator == null)
            {
                enumerator = new XVBComponentEnumerator(vBComponents);
            }
            return enumerator;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private Microsoft.Vbe.Interop.VBComponents vBComponents = null;
        private XVBComponentEnumerator enumerator = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (vBComponents != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                    enumerator?.Dispose();
                    enumerator = null;
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(vBComponents);

                vBComponents = null;
            }
        }

        ~XVBComponents()
        {
            Dispose(false);
        }

        private class XVBComponentEnumerator : IEnumerator<XVBComponent>
        {
            public XVBComponentEnumerator(Microsoft.Vbe.Interop.VBComponents vBComponents)
            {
                this.vBComponents = vBComponents;
            }

            public XVBComponent Current
            {
                get
                {
                    _Current?.Dispose();
                    _Current = new XVBComponent(vBComponents, index);
                    return _Current;
                }
            }
            object IEnumerator.Current => Current;

            public bool MoveNext()
            {
                if (index < vBComponents.Count)
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

                index = 0;
            }

            private Microsoft.Vbe.Interop.VBComponents vBComponents = null;
            private XVBComponent _Current = null;
            private int index = 0;

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            private void Dispose(bool disposing)
            {
                if (vBComponents != null)
                {
                    if (disposing)
                    {
                        // Free managed resources here.
                        _Current?.Dispose();
                        _Current = null;
                    }

                    // Free unmanaged resources here.

                    vBComponents = null;
                }
            }
        }
    }
}
