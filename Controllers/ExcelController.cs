using ExcelExport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using System.Reflection.Metadata.Ecma335;

namespace ExcelExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private List<ExcelSheetModel> GetExcelData()
        {
            List<ExcelSheetModel> lstData = new List<ExcelSheetModel>
            {
                new ExcelSheetModel { slno = 1, EmpName = "yamu", Department = "IT", Designation = "Manager" },
                new ExcelSheetModel { slno=2,EmpName="hhhh",Department="Software",Designation="Sr" },
                new ExcelSheetModel { slno=3,EmpName="jjj",Department="Hardware",Designation="Developer" },
                 new ExcelSheetModel { slno=4,EmpName="PPp",Department="Helpdesk",Designation="Assistant Mg" }


            };
            return lstData;
        }
        [HttpGet("GenerateExcel")]
        public async Task<ActionResult> GenerateExcel()
        {
            string strpath = Path.Combine(Directory.GetCurrentDirectory(), "Upload\\files", "TestReport" + DateTime.Now.Ticks.ToString() + ".xls");

            List<ExcelSheetModel> lstexcelSheetModels = GetExcelData();

            string strhtml = "<table style ='width:800px;border:solid;border-width:1px;'><thead><tr> ";

            strhtml += "<th style='width:10%;text-align:left'>slno</th>";
            strhtml += "<th style='width:30%;text-align:left'>Employee Name</th>";
            strhtml += "<th style='width:30%;text-align:left'>Department</th>";
            strhtml += "<th style='width:3s0%;text-align:left'>Designation</th>";

            strhtml += "</tr></thead>  <tbody>";

            foreach(ExcelSheetModel obj in lstexcelSheetModels)
            {
                strhtml += "<tr><td style='width:10%;text-align:left'>" + obj.slno.ToString() + " </td>";
                strhtml += "<td style='width:30%;text-align:left'>" + obj.EmpName + " </td>";
                strhtml += "<td style='width:30%;text-align:left'>" + obj.Department + " </td>";
                strhtml += "<td style='width:30%;text-align:left'>" + obj.Designation + " </td>";
            }

            strhtml += "</tbody> </table>";
            System.IO.File.AppendAllText(strpath, strhtml);

            var provider = new FileExtensionContentTypeProvider();

            if(!provider.TryGetContentType(strpath, out var contentType))
            {
                contentType = "application/octet-stream";
            }
            var bytes = await System.IO.File.ReadAllBytesAsync(strpath);
            return File(bytes, contentType,Path.GetFileName(strpath));
        }

    }
}
