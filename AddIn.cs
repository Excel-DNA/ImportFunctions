using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

namespace ImportFunctions
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations()
                             .RegisterFunctions();

            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
