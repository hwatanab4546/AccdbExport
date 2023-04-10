using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using MSAccess = Microsoft.Office.Interop.Access;

namespace XAccess
{
    public sealed class XApplication : IDisposable
    {
        public XApplication(string accdbPath)
        {
            application = new MSAccess.Application();

            application.OpenCurrentDatabase(accdbPath, true);
            application.Visible = false;
        }

        public bool IsQuit { get; private set; } = false;

        public void Quit()
        {
            if (!IsQuit)
            {
                while (disposables.Any())
                {
                    disposables.Pop().Dispose();
                }

                application.Quit(MSAccess.AcQuitOption.acQuitSaveNone);
            }
            IsQuit = true;
        }

        private XVBComponents _VBComponents = null;
        public XVBComponents VBComponents
        {
            get
            {
                if (_VBComponents == null)
                {
                    _VBComponents = new XVBComponents(application);
                    disposables.Push(_VBComponents);
                }
                return _VBComponents;
            }
        }

        public void SaveAsText(XAcObjectType objectType, string objectName, string fileName) => application.SaveAsText((MSAccess.AcObjectType)objectType, objectName, fileName);

        private XAllForms _AllForms = null;
        public XAllForms AllForms
        {
            get
            {
                if (_AllForms == null)
                {
                    _AllForms = new XAllForms(application);
                    disposables.Push(_AllForms);
                }
                return _AllForms;
            }
        }

        private XDatabase _CurrentDb = null;
        public XDatabase CurrentDb()
        {
            if (_CurrentDb == null)
            {
                _CurrentDb = new XDatabase(application);
                disposables.Push(_CurrentDb);
            }
            return _CurrentDb;
        }

        public void ExportXML(XAcExportXMLObjectType objectType, string dataSource, string schemaTarget = "")
            => application.ExportXML((MSAccess.AcExportXMLObjectType)objectType, dataSource, SchemaTarget: schemaTarget);

        private MSAccess.Application application = null;
        private readonly Stack<IDisposable> disposables = new Stack<IDisposable>();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (application != null)
            {
                if (disposing)
                {
                    // Free managed reosurces here.
                    Quit();
                }

                // Free unmanaged resources here.
                Marshal.FinalReleaseComObject(application);

                application = null;
            }
        }

        ~XApplication()
        {
            Dispose(false);
        }
    }
}
