//get currently generated quick data (for TODAY)
using QuickAutomatorService.Utils;
using QuickAutomatorService.Models;
using System.Threading.Tasks;

var Cfg = new Helper() { };

Console.WriteLine("Getting data from data store");

var configParam = await Cfg.GetConfigurationParamAsync(@"QuickRawDataFilePath");
var quick_data = await Cfg.GetQuickData(DateTime.Now);

if ((quick_data.Count() > 0) && (configParam.ConfigValue.Length > 0))
{
    Console.WriteLine("Data exist: generating excel workbook");

    if (XL.Create(quick_data, string.Format("{0}{1}.{2}",configParam.ConfigValue,@"Quick_Raw_Data_test",@"xlsx")))
    {
        Console.WriteLine("Excel Generation finished");
    }
}

Console.ReadKey();
