using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using PrintExcelToPdf.Services;
using System.Diagnostics;

namespace PrintExcelToPdf.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelToPrintController : ControllerBase
    {
        private readonly ConvertXlsxToPdf converterPdf;
        public ExcelToPrintController(ConvertXlsxToPdf converterPdf) {
            this.converterPdf = converterPdf;
        }

        [HttpPost("convert-excel-to-pdf")]
        public IActionResult Report([FromBody] Dictionary<string, object> jsonData)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string filePath = "./Templates/template_sunu.xlsx";

            if (jsonData == null || jsonData.Count == 0)
                return BadRequest("Données JSON invalides");

            if (System.IO.File.Exists(filePath))
            {
                var fileInfo = new FileInfo(filePath);

                if (fileInfo.IsReadOnly)
                {

                    return Ok("Le fichier est en lecture seule");
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

                using (var package = new ExcelPackage(new MemoryStream(fileBytes)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells["A6"].Value = jsonData["product"];
                    worksheet.Cells["B13"].Value = jsonData["transDate"];
                    worksheet.Cells["B15"].Value = jsonData["agencyName"];
                    worksheet.Cells["C17"].Value = jsonData["remarks"];
                    worksheet.Cells["C19"].Value = jsonData["dateStart"];
                    worksheet.Cells["F19"].Value = jsonData["dateEnd"];
                    worksheet.Cells["C21"].Value = jsonData["amount"];

                    worksheet.Calculate();

                    var stream = new MemoryStream();

                    package.SaveAs(stream);
                    stream.Position = 0;

                    var pdfStream = this.converterPdf.Convert(stream);

                    return File(pdfStream, "application/pdf", "converted.pdf");
                }

            }

            return BadRequest();
        }
    }
}
