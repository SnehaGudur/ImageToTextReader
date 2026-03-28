using Microsoft.AspNetCore.Mvc;
using Tesseract;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace ImageToTextReader.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
        }

        [HttpGet]
        public IActionResult Index() => View();

        [HttpPost]
        public IActionResult Index(IFormFile imageFile)
        {
            if (imageFile == null || imageFile.Length == 0)
            {
                ViewBag.Result = "Please upload an image.";
                return View();
            }

            // Save uploaded image
            string uploadsFolder = Path.Combine(_env.WebRootPath, "uploads");
            if (!Directory.Exists(uploadsFolder))
                Directory.CreateDirectory(uploadsFolder);

            string filePath = Path.Combine(uploadsFolder, Path.GetFileName(imageFile.FileName));
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                imageFile.CopyTo(stream);
            }

            try
            {
                // Initialize Tesseract
                string tessdataPath = Path.Combine(_env.ContentRootPath, "tessdata");

                if (!Directory.Exists(tessdataPath))
                    throw new DirectoryNotFoundException($"tessdata not found at: {tessdataPath}");

                string trainedData = Path.Combine(tessdataPath, "eng.traineddata");
                if (!System.IO.File.Exists(trainedData))
                    throw new FileNotFoundException($"eng.traineddata missing at: {trainedData}");

                string extractedText = "";

                using (var engine = new TesseractEngine(tessdataPath, "eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(filePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            extractedText = page.GetText();
                            ViewBag.Result = extractedText;
                        }
                    }
                }

                // =========================
                // CREATE EXCEL FILE (HORIZONTAL)
                // =========================

                ExcelPackage.License.SetNonCommercialPersonal("OCR App");

                string excelPath = Path.Combine(uploadsFolder, "OCR_Result.xlsx");

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("OCR Result");

                    // Clean and split lines
                    var lines = extractedText
                        .Split('\n')
                        .Where(line => !string.IsNullOrWhiteSpace(line))
                        .ToArray();

                    // Write horizontally (Row 1, Columns increasing)
                    for (int i = 0; i < lines.Length; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = lines[i];
                    }

                    System.IO.File.WriteAllBytes(excelPath, package.GetAsByteArray());
                }
            }
            catch (Exception ex)
            {
                ViewBag.Result = "OCR Error: " + ex.Message;
            }

            return View();
        }

        public IActionResult DownloadExcel()
        {
            string filePath = Path.Combine(_env.WebRootPath, "uploads", "OCR_Result.xlsx");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Excel file not found.");
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            return File(fileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "OCR_Result.xlsx");
        }
    }
}