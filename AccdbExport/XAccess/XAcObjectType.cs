using System.Runtime.InteropServices;

namespace XAccess
{
    public enum XAcObjectType
    {
        acDefault = -1,
        acTable,
        acQuery,
        acForm,
        acReport,
        acMacro,
        acModule,
        [TypeLibVar(64)]
        acDataAccessPage,
        acServerView,
        acDiagram,
        acStoredProcedure,
        acFunction,
        acDatabaseProperties,
        acTableDataMacro
    }
}
