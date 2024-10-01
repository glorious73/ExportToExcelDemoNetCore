using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ExportListToExcelDemo.Models;
using ExportListToExcelDemo.Service;
using System.Collections.Generic;
using ExportListToExcelDemo.Entity;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using System;

namespace ExportListToExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        private IUserService _userService { get; set; }

        public HomeController(IUserService userService)
        {
            _userService = userService;
        }
        public IActionResult Index()
        {
            List<User> users = _userService.GetAllUsers();
            return View(users);
        }

        public IActionResult ExportToExcel()
        {
            IWorkbook workbook = _userService.ExportUsersToExcel();
            string contentType = ""; // Scope
            // Credit for two stream since workbook.write() closes the first one: https://stackoverflow.com/a/36584861/6336270 
            MemoryStream tempStream = null;
            MemoryStream stream = null;
            try
            {
                // 1. Write the workbook to a temporary stream
                tempStream = new MemoryStream();
                workbook.Write(tempStream);
                // 2. Convert the tempStream to byteArray and copy to another stream
                var byteArray = tempStream.ToArray();
                stream = new MemoryStream();
                stream.Write(byteArray, 0, byteArray.Length);
                stream.Seek(0, SeekOrigin.Begin);
                // 3. Set file content type
                contentType = workbook.GetType() == typeof(XSSFWorkbook) ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel";
                // 4. Return file
                return File(
                    fileContents: stream.ToArray(),
                    contentType: contentType,
                    fileDownloadName: "Demo Users " + DateTime.Now.ToString() + ((workbook.GetType() == typeof(XSSFWorkbook)) ? ".xlsx" : ".xls"));
            }
            finally
            {
                if (tempStream != null) tempStream.Dispose();
                if (stream != null) stream.Dispose();
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
