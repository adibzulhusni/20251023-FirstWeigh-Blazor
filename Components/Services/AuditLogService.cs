using FirstWeigh.Models;
using ClosedXML.Excel;

namespace FirstWeigh.Services
{
    public class AuditLogService
    {
        private readonly IConfiguration _configuration;
        private readonly string _auditLogFilePath;
        private static readonly SemaphoreSlim _fileLock = new(1, 1);

        public AuditLogService(IConfiguration configuration)
        {
            _configuration = configuration;
            _auditLogFilePath = _configuration["DataStorage:AuditLogFilePath"] ?? "Data/AuditLog.xlsx";

            var directory = Path.GetDirectoryName(_auditLogFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(_auditLogFilePath))
            {
                InitializeFile();
            }
        }

        private void InitializeFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("AuditLog");

            worksheet.Cell(1, 1).Value = "Timestamp";
            worksheet.Cell(1, 2).Value = "Username";
            worksheet.Cell(1, 3).Value = "Action";
            worksheet.Cell(1, 4).Value = "Details";
            worksheet.Cell(1, 5).Value = "IpAddress";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_auditLogFilePath);
        }

        public async Task LogAsync(string username, string action, string details, string ipAddress = "")
        {
            await _fileLock.WaitAsync();
            try
            {
                var logs = await GetAllLogsAsync();
                logs.Add(new AuditLog
                {
                    Timestamp = DateTime.Now,
                    Username = username,
                    Action = action,
                    Details = details,
                    IpAddress = ipAddress
                });

                await SaveLogsAsync(logs);
            }
            finally
            {
                _fileLock.Release();
            }
        }

        public async Task<List<AuditLog>> GetAllLogsAsync()
        {
            await _fileLock.WaitAsync();
            try
            {
                if (!File.Exists(_auditLogFilePath))
                {
                    return new List<AuditLog>();
                }

                var logs = new List<AuditLog>();

                using var workbook = new XLWorkbook(_auditLogFilePath);
                var worksheet = workbook.Worksheet(1);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                if (lastRow <= 1) return logs;

                for (int row = 2; row <= lastRow; row++)
                {
                    logs.Add(new AuditLog
                    {
                        Timestamp = worksheet.Cell(row, 1).GetDateTime(),
                        Username = worksheet.Cell(row, 2).GetString(),
                        Action = worksheet.Cell(row, 3).GetString(),
                        Details = worksheet.Cell(row, 4).GetString(),
                        IpAddress = worksheet.Cell(row, 5).GetString()
                    });
                }

                return logs.OrderByDescending(l => l.Timestamp).ToList();
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private async Task SaveLogsAsync(List<AuditLog> logs)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("AuditLog");

            worksheet.Cell(1, 1).Value = "Timestamp";
            worksheet.Cell(1, 2).Value = "Username";
            worksheet.Cell(1, 3).Value = "Action";
            worksheet.Cell(1, 4).Value = "Details";
            worksheet.Cell(1, 5).Value = "IpAddress";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

            int rowIndex = 2;
            foreach (var log in logs)
            {
                worksheet.Cell(rowIndex, 1).Value = log.Timestamp;
                worksheet.Cell(rowIndex, 2).Value = log.Username;
                worksheet.Cell(rowIndex, 3).Value = log.Action;
                worksheet.Cell(rowIndex, 4).Value = log.Details;
                worksheet.Cell(rowIndex, 5).Value = log.IpAddress;
                rowIndex++;
            }

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(_auditLogFilePath);
        }
    }
}