using FirstWeigh.Models;
using ClosedXML.Excel;

namespace FirstWeigh.Services
{
    public class LoginAttemptService
    {
        private readonly IConfiguration _configuration;
        private readonly string _attemptsFilePath;
        private static readonly SemaphoreSlim _fileLock = new(1, 1);
        private const int MAX_ATTEMPTS = 5;
        private const int LOCKOUT_MINUTES = 30;

        public LoginAttemptService(IConfiguration configuration)
        {
            _configuration = configuration;
            _attemptsFilePath = _configuration["DataStorage:LoginAttemptsFilePath"] ?? "Data/LoginAttempts.xlsx";

            var directory = Path.GetDirectoryName(_attemptsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(_attemptsFilePath))
            {
                InitializeFile();
            }
        }

        private void InitializeFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("LoginAttempts");

            worksheet.Cell(1, 1).Value = "Username";
            worksheet.Cell(1, 2).Value = "IpAddress";
            worksheet.Cell(1, 3).Value = "Timestamp";
            worksheet.Cell(1, 4).Value = "Reason";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_attemptsFilePath);
        }

        public async Task RecordFailedAttemptAsync(string username, string ipAddress, string reason)
        {
            await _fileLock.WaitAsync();
            try
            {
                var attempts = await GetAllAttemptsAsync();
                attempts.Add(new LoginAttempt
                {
                    Username = username,
                    IpAddress = ipAddress,
                    Timestamp = DateTime.Now,
                    Reason = reason
                });

                await SaveAttemptsAsync(attempts);
            }
            finally
            {
                _fileLock.Release();
            }
        }

        public async Task<bool> IsAccountLockedAsync(string username)
        {
            var attempts = await GetRecentAttemptsAsync(username);

            if (attempts.Count >= MAX_ATTEMPTS)
            {
                var lastAttempt = attempts.OrderByDescending(a => a.Timestamp).First();
                var lockoutExpiry = lastAttempt.Timestamp.AddMinutes(LOCKOUT_MINUTES);

                if (DateTime.Now < lockoutExpiry)
                {
                    return true;
                }
                else
                {
                    // Lockout expired, clear old attempts
                    await ClearAttemptsForUserAsync(username);
                    return false;
                }
            }

            return false;
        }

        public async Task<DateTime?> GetLockoutExpiryAsync(string username)
        {
            var attempts = await GetRecentAttemptsAsync(username);

            if (attempts.Count >= MAX_ATTEMPTS)
            {
                var lastAttempt = attempts.OrderByDescending(a => a.Timestamp).First();
                return lastAttempt.Timestamp.AddMinutes(LOCKOUT_MINUTES);
            }

            return null;
        }

        public async Task ClearAttemptsForUserAsync(string username)
        {
            await _fileLock.WaitAsync();
            try
            {
                var allAttempts = await GetAllAttemptsAsync();
                var remaining = allAttempts.Where(a => !a.Username.Equals(username, StringComparison.OrdinalIgnoreCase)).ToList();
                await SaveAttemptsAsync(remaining);
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private async Task<List<LoginAttempt>> GetRecentAttemptsAsync(string username)
        {
            var allAttempts = await GetAllAttemptsAsync();
            var cutoffTime = DateTime.Now.AddMinutes(-LOCKOUT_MINUTES);

            return allAttempts
                .Where(a => a.Username.Equals(username, StringComparison.OrdinalIgnoreCase) && a.Timestamp > cutoffTime)
                .ToList();
        }

        private async Task<List<LoginAttempt>> GetAllAttemptsAsync()
        {
            await _fileLock.WaitAsync();
            try
            {
                if (!File.Exists(_attemptsFilePath))
                {
                    return new List<LoginAttempt>();
                }

                var attempts = new List<LoginAttempt>();

                using var workbook = new XLWorkbook(_attemptsFilePath);
                var worksheet = workbook.Worksheet(1);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                if (lastRow <= 1) return attempts;

                for (int row = 2; row <= lastRow; row++)
                {
                    attempts.Add(new LoginAttempt
                    {
                        Username = worksheet.Cell(row, 1).GetString(),
                        IpAddress = worksheet.Cell(row, 2).GetString(),
                        Timestamp = worksheet.Cell(row, 3).GetDateTime(),
                        Reason = worksheet.Cell(row, 4).GetString()
                    });
                }

                return attempts;
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private async Task SaveAttemptsAsync(List<LoginAttempt> attempts)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("LoginAttempts");

            worksheet.Cell(1, 1).Value = "Username";
            worksheet.Cell(1, 2).Value = "IpAddress";
            worksheet.Cell(1, 3).Value = "Timestamp";
            worksheet.Cell(1, 4).Value = "Reason";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

            int rowIndex = 2;
            foreach (var attempt in attempts)
            {
                worksheet.Cell(rowIndex, 1).Value = attempt.Username;
                worksheet.Cell(rowIndex, 2).Value = attempt.IpAddress;
                worksheet.Cell(rowIndex, 3).Value = attempt.Timestamp;
                worksheet.Cell(rowIndex, 4).Value = attempt.Reason;
                rowIndex++;
            }

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(_attemptsFilePath);
        }
    }
}