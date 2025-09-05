#pragma warning disable CS8602 // 解引用可能出现空引用。
#pragma warning disable CS8600 // 将 null 字面量或可能为 null 的值转换为非 null 类型。
namespace Function;
public partial class Function
{
    #region JSON
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
    private static object[,] ParseJsonToExcelArray(string json, int dim, bool isDebug)
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

    #region Process
    
    public static Process AddProcess(Dictionary<string, Process> addInPrcDict, string processId)
    {
        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = "python",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        Process process = Process.Start(psi) ?? throw new InvalidOperationException("无法启动Python");
        addInPrcDict[processId] = process;
        return process;
    }

    public static Process GetProcess(Dictionary<string, Process> addInPrcDict, string processId)
    {
        if (Global.addIn == null) throw new InvalidOperationException("无法获取进程");
        if (addInPrcDict.TryGetValue(processId, out Process process) || process.HasExited)
        { 
            return process ;
        }
        return AddProcess(addInPrcDict, processId);
        
    }
    private static Process GetPythonProcess(string processId)
    {
        return GetProcess(Global.addIn.pythonProcess, processId);
    }

    #endregion
}

