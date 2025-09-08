namespace Function;
public partial class Function
{
    [ExcelFunction(
        Description = "异步执行 Python 命令，返回二维数组",
        IsMacroType = true,
        IsThreadSafe = false
    )]
    public static async Task<object[,]> RunPythonCode(
        [ExcelArgument(Description = "多行 Python 命令")] string pyCode,
        [ExcelArgument(Description = "解析维度：0 直接返回，1 一维，2 二维")] int dim,
        [ExcelArgument(Description = "python解释器路径：默认为：python，即全局变量中的python解释器")] string pyPath,
        [ExcelArgument(Description = "启用Debug：默认为：FALSE，即不启用Debug模式")] bool isDbug
    )
    {
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".py");
        File.WriteAllText(tempPath, pyCode);
        return await RunPtyon(tempPath, dim, pyPath, isDbug, true);
    }

    [ExcelFunction(
        Description = "异步执行 Python 文件，返回二维数组",
        IsMacroType = true,
        IsThreadSafe = false
    )]
    public static async Task<object[,]> RunPythonFile(
        [ExcelArgument(Description = "Python 文件路径")] string filePath,
        [ExcelArgument(Description = "解析维度：0 直接返回，1 一维，2 二维")] int dim,
        [ExcelArgument(Description = "python解释器路径：默认为：\"python\"，即全局变量中的python解释器")] string pyPath,
        [ExcelArgument(Description = "启用Debug：默认为：FALSE，即不启用Debug模式")] bool isDbug
    )
    {
        pyPath = string.IsNullOrWhiteSpace(pyPath) ? "python" : pyPath;
        return await RunPtyon(filePath, dim, pyPath, isDbug, false);
    }
    private static async Task<object[,]> RunPtyon(string filePath, int dim, string pyPath, bool isDebug = false, bool removeTempFile = false)
    {
        return await Task.Run(() =>
        {
            pyPath = string.IsNullOrWhiteSpace(pyPath) ? "python" : pyPath;
            if (!File.Exists(filePath) || string.IsNullOrWhiteSpace(filePath))
                return new object[,] { { "Python 文件不存在。" } };
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pyPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using Process? process = Process.Start(psi);
                if (process == null)
                    return new object[,] { { "无法启动 Python。" } };
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(error))
                    throw new Exception(error);
                return ParseJsonToExcelArray(output, dim, isDebug);
            }
            catch (Exception e)
            {
                if (isDebug) return new object[,] { { e.ToString() } };
                throw;
            }
            finally
            {
                if (removeTempFile & File.Exists(filePath)) File.Delete(filePath);
            }
        }
            );
    }

    private static async Task<object[,]> RunPythonGetVar(string filePath, string varName, int dim, string pyPath, bool isDebug, bool removeTempFile = false)
    {
        const string PyCodeHeader = "from json import dumps as _InjectedJsonDumps\n" +
                                    "_CSharpGetPythonVar = lambda varName : print(_InjectedJsonDumps(varName))\n";
        string PyCodeLast = $"\n_CSharpGetPythonVar({varName})";
        return await Task.Run(() =>
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pyPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
            }
            catch (Exception e)
            {
                if (isDebug) return new object[,] { { e.ToString() } };
                throw;
            }
            finally
            {
                if (removeTempFile & File.Exists(filePath)) File.Delete(filePath);
            }
            return new object[,] { { "" } };
        });
    }
}
