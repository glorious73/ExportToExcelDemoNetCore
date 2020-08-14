using ExportListToExcelDemo.Entity;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExportListToExcelDemo.Service
{
    public interface IUserService
    {
        List<User> GetAllUsers();
        IWorkbook ExportUsersToExcel();
    }
}
