using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using ExcelDemo.Models;
using System.Text;
using System.Collections.ObjectModel;
using System.Security.Cryptography.Xml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace ExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        private List<Employee> employees = new List<Employee>
       {
           new Employee{EmpId=1, EmpName="Sanjsigan", JoinDate="1-jan-2020"},
           new Employee{EmpId=2, EmpName="Antony", JoinDate="1-jan-2020"},
           new Employee{EmpId=3, EmpName="Mathu", JoinDate="1-jan-2020"},
            new Employee{EmpId=4, EmpName="Vijay", JoinDate="1-jan-2020"}
       };

        public IActionResult Index()
        {
            // return CSV();
            return Excel();
        }

  /*     public IActionResult CSV()
        {
            var builder = new StringBuilder();
            builder.AppendLine("EmpId,EmpName,JoinDate");
            foreach(var emp in employees)
            {
                builder.AppendLine($"{emp.EmpId},{emp.EmpName},{emp.JoinDate}");
            }
            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "employeeInfo.csv");


        }*/

        public IActionResult Excel()
        {
            using (var workbook= new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employees");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "EmpId";
                worksheet.Cell(currentRow, 2).Value = "EmpName";
                worksheet.Cell(currentRow, 3).Value = "JoinDate";


                foreach(var emp in employees)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = emp.EmpId;
                    worksheet.Cell(currentRow, 2).Value = emp.EmpName;
                    worksheet.Cell(currentRow, 3).Value = emp.JoinDate;
                    
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content,
                        "application/vnd.openxmlformats-officedocuments.spreadsheetml.sheet",
                        "EmployeeInfo.xlsx");
                }
                    
            }
        }
    }
}
