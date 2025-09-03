global using ExcelDna.Integration;
global using ExcelDna.IntelliSense;
global using ExcelDna.Registration;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Text.Json;

#pragma warning disable CS8600 // 将 null 字面量或可能为 null 的值转换为非 null 类型。

namespace Function
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelRegistration.GetExcelFunctions()
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                .ProcessParameterConversions(GetAsyncFunctionConfig())
                .RegisterFunctions();
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
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

    public class Function
    {
        #region Test
        [ExcelFunction(Description = "My first .NET function")]
        public static async Task<string> SayHello(
            [ExcelArgument(Description = "name")] string name,
            [ExcelArgument(Description = "delay")] int time
        )
        {
            await Task.Delay(time * 1000);
            return "Hello " + name;
        }
        #endregion

        #region RunPython
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
                    using Process process = Process.Start(psi);
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
                    if (removeTempFile & File.Exists(filePath))
                        File.Delete(filePath);
                }
            }
                );
        }
        #endregion

        #region GetFilesPath
        [ExcelFunction(
                    Description = "获取指定路径下的文件列表，支持通配符，可选择返回文件夹名",
                    IsMacroType = true,
                    IsThreadSafe = false
                )]
        public static async Task<object> GetFilesPath(
                    [ExcelArgument(Description = "文件夹路径")] string folderPath,
                    [ExcelArgument(Description = "匹配模式：0 文件，1 文件夹, 3 所有")] int matchType = 0,
                    [ExcelArgument(Description = "通配符筛选，默认 *")] string filter = "*",
                    [ExcelArgument(Description = "是否完整路径：TRUE 全部路径，FALSE 文件名")] bool fullPath = false
                )
        {
            return await Task.Run(() =>
            {
                if (!Directory.Exists(folderPath))
                    return new object[,] { { "目录不存在" } };
                filter = string.IsNullOrEmpty(filter) ? "*" : filter;
                string[] files;
                switch (matchType)
                {
                    case 0:
                        files = Directory.GetFiles(folderPath, filter);
                        break;
                    case 1:
                        files = Directory.GetDirectories(folderPath, filter);
                        break;
                    case 2:
                        files = Directory.GetFileSystemEntries(folderPath, filter);
                        break;
                    default:
                        return new object[,] { { "模式错误" } };
                }
                if (files.Length == 0)
                    return new object[,] { { "无文件" } };
                object[,] result = new object[files.Length, 1];
                for (int i = 0; i < files.Length; i++)
                    result[i, 0] = fullPath ? files[i] : Path.GetFileName(files[i]);
                return result;
            });
        }
        #endregion

        #region 内部方法
        /// <summary>
        /// 解析JsonElement
        /// </summary>
        /// <param name="elem">JsonElement 元素</param>
        /// <param name="isDebug">是否返回错误信息</param>
        /// <returns></returns>
        private static object ConvertObjectType(JsonElement elem, bool isDebug)
        {
            try
            {
                switch (elem.ValueKind)
                {
                    case JsonValueKind.Number: return elem.GetDouble();
                    case JsonValueKind.Array: return elem.GetRawText();
                    case JsonValueKind.String: return elem.ToString();
                    case JsonValueKind.True:
                    case JsonValueKind.False: return elem.GetBoolean();
                    case JsonValueKind.Object: return elem.GetRawText();
                    case JsonValueKind.Undefined:
                    case JsonValueKind.Null: return ExcelError.ExcelErrorNA;
                    default: return ExcelError.ExcelErrorValue;
                }
            }
            catch (Exception e)
            {
                if (isDebug) return new object[,] { { e.ToString() } };
                throw;
            }
        }
        /// <summary>
        /// 将 JSON 字符串解析为 Excel 可识别的二维数组
        /// dim: 0 直接返回字符串, 1 一维数组, 2 二维数组（支持不规则补齐 #N/A）
        /// </summary>
        /// <param name="json">JSON 字符串</param>
        /// <param name="dim">返回维度</param>
        /// <param name="isDebug">是否返回错误信息</param>
        public static object[,] ParseJsonToExcelArray(string json, int dim, bool isDebug)
        {
            try
            {
                object[,] result;
                switch (dim)
                {
                    case 0:
                        result = new object[1, 1];
                        result[0, 0] = json;
                        break;
                    case 1:
                        var json1 = JsonSerializer.Deserialize<JsonElement[]>(json);
                        if (json1 == null) return new object[,] { { ExcelError.ExcelErrorNA } };
                        result = new object[1, json1.Length];
                        for (int i = 0; i < json1.Length; i++)
                            result[0, i] = ConvertObjectType(json1[i], isDebug);
                        break;
                    case 2:
                        var jsonOuter = JsonSerializer.Deserialize<JsonElement[]>(json);
                        if (jsonOuter == null) return new object[,] { { ExcelError.ExcelErrorNA } };
                        int rows = jsonOuter.Length;
                        int cols = 0;
                        foreach (var rowElem in jsonOuter)
                        {
                            if (rowElem.ValueKind == JsonValueKind.Array)
                            {
                                int len = rowElem.GetArrayLength();
                                if (len > cols) cols = len;
                            }
                            else
                            {
                                cols = Math.Max(cols, 1);
                            }
                        }
                        result = new object[rows, cols];
                        for (int i = 0; i < rows; i++)
                        {
                            if (jsonOuter[i].ValueKind == JsonValueKind.Array)
                            {
                                var row = jsonOuter[i].EnumerateArray().ToArray();
                                for (int j = 0; j < cols; j++)
                                {
                                    result[i, j] = j < row.Length ? ConvertObjectType(row[j], isDebug) : ExcelError.ExcelErrorNA;
                                }
                            }
                            else
                            {
                                result[i, 0] = ConvertObjectType(jsonOuter[i], isDebug);
                                for (int j = 1; j < cols; j++)
                                    result[i, j] = ExcelError.ExcelErrorNA;
                            }
                        }
                        break;
                    default:
                        if (isDebug) return new object[,] { { "传入的解析模式错误" } };
                        return new object[,] { { ExcelError.ExcelErrorValue } };
                }
                return result;
            }
            catch (Exception e)
            {
                if (isDebug) return new object[,] { { e.ToString() } }; throw;
            }
        }
        #endregion
    }

}
