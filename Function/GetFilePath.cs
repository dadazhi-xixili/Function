namespace Function;
public partial class Function
{
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
}

