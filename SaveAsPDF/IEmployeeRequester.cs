using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SaveAsPDF.Models;
namespace SaveAsPDF
{
    public interface IEmployeeRequester
    {
        void EmployeeComplete(EmployeeModel model);
    }
}
