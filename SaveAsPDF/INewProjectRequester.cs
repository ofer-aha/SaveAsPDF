using SaveAsPDF.Models;

namespace SaveAsPDF
{
    public interface INewProjectRequester
    {

        void NewProjectComplete(ProjectModel model);

    }
}
