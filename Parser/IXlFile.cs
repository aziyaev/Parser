using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public interface IXlFile
    {
        void OpenTable();
        void SaveTable();
        XlParser LoadTable(out string message);
        string UpdateTable();
    }
}
