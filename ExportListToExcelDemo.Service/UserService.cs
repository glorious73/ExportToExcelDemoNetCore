using ExportListToExcelDemo.Entity;
using ExportListToExcelDemo.Utility;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExportListToExcelDemo.Service
{
    public class UserService : IUserService
    {
        private IExportUtility _exportUtility { get; set; }

        public UserService(IExportUtility exportUtility)
        {
            _exportUtility = exportUtility;
        }
        public List<User> GetAllUsers()
        {
            List<User> users = new List<User>();
            // Example user list
            for (int i = 0; i < 50; i++)
            {
                users.Add(new User()
                {
                    Id = i+1,
                    FirstName = "Amjad" + (i+1),
                    LastName  = "Aj" + (i+1),
                    EmailAddress = "someemail@example.com",
                    PhoneNumber = "2025550168"
                });
            }
            return users;
        }

        public IWorkbook ExportUsersToExcel()
        {
            // 1. Get all users
            List<User> users = GetAllUsers();
            // 2. Return users Excel workbook
            return _exportUtility.WriteExcelWithNPOI(users, "xlsx");
        }
    }
}
