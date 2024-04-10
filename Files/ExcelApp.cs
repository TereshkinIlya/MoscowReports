using Excel = Microsoft.Office.Interop.Excel;

namespace Files
{
    public static class ExcelApp
    {
        public static Excel.Application Run { get; set; } = new Excel.Application();
        static ExcelApp() 
        {
            Run.IgnoreRemoteRequests = true;
        }
    }
}