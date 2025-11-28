

using SaveAsPDF.Models;

namespace SaveAsPDF
{
    public interface ISettingsRequester
    {
        void SettingsComplete(SettingsModel model);
    }
}