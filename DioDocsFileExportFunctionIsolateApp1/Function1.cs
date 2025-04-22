using GrapeCity.Documents.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace DioDocsFileExportFunctionIsolateApp1
{
    public class Function1
    {
        private readonly ILogger<Function1> _logger;

        public Function1(ILogger<Function1> logger)
        {
            _logger = logger;
        }

        [Function("Function1")]
        public IActionResult Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            string? name = req.Query["name"].FirstOrDefault();

            string requestBody = new StreamReader(req.Body).ReadToEndAsync().Result;
            dynamic? data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            string Message = string.IsNullOrEmpty(name)
                ? "Hello, World!!"
                : $"Hello, {name}!!";

            Workbook workbook = new();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            byte[] output;

            using (MemoryStream ms = new())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);
                output = ms.ToArray();
            }

            return new FileContentResult(output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "Result.xlsx"
            };
        }
    }
}
