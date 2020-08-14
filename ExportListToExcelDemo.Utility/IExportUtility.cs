using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExportListToExcelDemo.Utility
{
    public interface IExportUtility
    {
        IWorkbook WriteExcelWithNPOI<T>(List<T> data, string extension);
    }
}
