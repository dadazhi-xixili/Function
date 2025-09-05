global using ExcelDna.Integration;
global using ExcelDna.IntelliSense;
global using ExcelDna.Registration;
global using System.Diagnostics;
global using System.Linq.Expressions;
global using System.Text.Json;
#pragma warning disable CA2211

namespace Function;

public class AddIn : IExcelAddIn
{
    public Dictionary<string, Process> pythonProcess = new();
    public void AutoOpen()
    {
        Global.addIn = this;
        ExcelRegistration.GetExcelFunctions()
            .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
            .ProcessParameterConversions(GetAsyncFunctionConfig())
            .RegisterFunctions();
        IntelliSenseServer.Install();
    }
    public void AutoClose()
    {
        foreach (Process process in pythonProcess.Values)
        {
            process.Close();
            process.Dispose();
        }
        IntelliSenseServer.Uninstall();
        Global.addIn = null;
    }

    private static ParameterConversionConfiguration GetAsyncFunctionConfig()
    {
        const string rval = "loading...";
        ParameterConversionConfiguration pcc = new ParameterConversionConfiguration();
        pcc.AddReturnConversion((type, _) =>
            type != typeof(object)
                ? null
                : (Expression<Func<object, object>>)(retValue => retValue.Equals(ExcelError.ExcelErrorNA) ? rval : retValue));
        return pcc;
    }
}

public static class Global
{
    public static AddIn? addIn;
}
