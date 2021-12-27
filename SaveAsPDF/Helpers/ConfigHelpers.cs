using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public class ConfigHelpers
    {
        public static void EditAppSetting(string key, string value)
        {
            try
            {
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings[key].Value = value;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error update file App.config: " + ex.Message, "Thông báo kiểm tra !", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
        }
    }
}
