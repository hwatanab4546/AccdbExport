using System;
using System.Runtime.InteropServices;

namespace XAccess
{
    public sealed class XVBComponent : IDisposable
    {
        public XVBComponent(Microsoft.Vbe.Interop.VBComponents vBComponents, int index)
        {
            vBComponent = vBComponents.Item(index);

            Name = vBComponent.Name;
            Type = (Xvbext_ComponentType)vBComponent.Type;
        }

        public string Name { get; }

        public Xvbext_ComponentType Type { get; }

        public void Export(string filename) => vBComponent.Export(filename);

        private Microsoft.Vbe.Interop.VBComponent vBComponent = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (vBComponent != null)
            {
                if (disposing)
                {
                    // Free managed resources here.
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(vBComponent);

                vBComponent = null;
            }
        }

        ~XVBComponent()
        {
            Dispose(false);
        }
    }
}
