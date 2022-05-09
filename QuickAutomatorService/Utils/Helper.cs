using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using QuickAutomatorService.Data;
using QuickAutomatorService.Models;
using Microsoft.EntityFrameworkCore;

using System.Diagnostics;

namespace QuickAutomatorService.Utils
{
    public class Helper
    {
        public DautomateContext config = null;

        public Helper()
        {
            config = new DautomateContext();
        }

        public async Task<List<Quick>> GetQuickData(DateTime dt)
        {
            //method is responsible for fetching quick data generated for the data
            List<Quick> data = new List<Quick>();

            try
            {
                data = await config.Quicks.Where(q => q.DateUploaded == dt).ToListAsync();
                return data;
            }
            catch(Exception ex)
            {
                Debug.Print(ex.Message);
                return data;
            }
        }

        public async Task<ConfigParam> GetConfigurationParamAsync(string key)
        {
            try
            {
                return await config.ConfigParams.Where(cfg => cfg.ConfigKey == key).FirstOrDefaultAsync();
            }
            catch(Exception ex)
            {
                Debug.Print(ex.Message);
                return null;
            }
        }

    }
}
