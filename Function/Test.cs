using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Function;
public partial class Function
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
}




