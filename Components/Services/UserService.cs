using FirstWeigh.Models;
using ClosedXML.Excel;
using System.Text.Json;

namespace FirstWeigh.Services
{
    public class UserService
    {
        private readonly IConfiguration _configuration;
        private readonly string _usersFilePath;
        private readonly string _backupFolder;
        private static readonly SemaphoreSlim _fileLock = new(1, 1);

        public UserService(IConfiguration configuration)
        {
            _configuration = configuration;
            _usersFilePath = _configuration["DataStorage:UsersFilePath"] ?? "Data/Users.xlsx";
            _backupFolder = _configuration["DataStorage:BackupFolder"] ?? "Backups/Users";

            // Ensure directories exist
            var directory = Path.GetDirectoryName(_usersFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!Directory.Exists(_backupFolder))
            {
                Directory.CreateDirectory(_backupFolder);
            }

            // Create default admin if file doesn't exist
            if (!File.Exists(_usersFilePath))
            {
                InitializeWithDefaultAdmin();
            }
        }

        private void InitializeWithDefaultAdmin()
        {
            var defaultAdmin = new User
            {
                UserId = "USER001",
                Username = "admin",
                Password = "admin123",
                FullName = "System Administrator",
                Role = UserRoles.Admin,
                IsActive = true,
                CreatedDate = DateTime.Now,
                LastModifiedDate = DateTime.Now,
                LastModifiedBy = "system"
            };

            var users = new List<User> { defaultAdmin };
            SaveUsers(users);
        }

        public async Task<List<User>> GetAllUsersAsync()
        {
            await _fileLock.WaitAsync();
            try
            {
                if (!File.Exists(_usersFilePath))
                {
                    return new List<User>();
                }

                var users = new List<User>();

                using var workbook = new XLWorkbook(_usersFilePath);
                var worksheet = workbook.Worksheet(1); // First worksheet

                // Find last used row
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                if (lastRow <= 1) // Only header or empty
                {
                    return users;
                }

                // Read from row 2 (skip header)
                for (int row = 2; row <= lastRow; row++)
                {
                    var user = new User
                    {
                        UserId = worksheet.Cell(row, 1).GetString(),
                        Username = worksheet.Cell(row, 2).GetString(),
                        Password = worksheet.Cell(row, 3).GetString(),
                        FullName = worksheet.Cell(row, 4).GetString(),
                        Role = worksheet.Cell(row, 5).GetString(),
                        IsActive = worksheet.Cell(row, 6).GetBoolean(),
                        CreatedDate = worksheet.Cell(row, 7).GetDateTime(),
                        LastModifiedDate = worksheet.Cell(row, 8).GetDateTime(),
                        LastModifiedBy = worksheet.Cell(row, 9).GetString(),
                        LastLoginDate = worksheet.Cell(row, 10).IsEmpty() ? null : worksheet.Cell(row, 10).GetDateTime()
                    };

                    users.Add(user);
                }

                return users;
            }
            finally
            {
                _fileLock.Release();
            }
        }

        public async Task SaveUsers(List<User> users)
        {
            await _fileLock.WaitAsync();
            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Users");

                // Create header row
                worksheet.Cell(1, 1).Value = "UserId";
                worksheet.Cell(1, 2).Value = "Username";
                worksheet.Cell(1, 3).Value = "Password";
                worksheet.Cell(1, 4).Value = "FullName";
                worksheet.Cell(1, 5).Value = "Role";
                worksheet.Cell(1, 6).Value = "IsActive";
                worksheet.Cell(1, 7).Value = "CreatedDate";
                worksheet.Cell(1, 8).Value = "LastModifiedDate";
                worksheet.Cell(1, 9).Value = "LastModifiedBy";
                worksheet.Cell(1, 10).Value = "LastLoginDate";

                // Style header row
                var headerRow = worksheet.Row(1);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Write data rows
                int rowIndex = 2;
                foreach (var user in users)
                {
                    worksheet.Cell(rowIndex, 1).Value = user.UserId;
                    worksheet.Cell(rowIndex, 2).Value = user.Username;
                    worksheet.Cell(rowIndex, 3).Value = user.Password;
                    worksheet.Cell(rowIndex, 4).Value = user.FullName;
                    worksheet.Cell(rowIndex, 5).Value = user.Role;
                    worksheet.Cell(rowIndex, 6).Value = user.IsActive;
                    worksheet.Cell(rowIndex, 7).Value = user.CreatedDate;
                    worksheet.Cell(rowIndex, 8).Value = user.LastModifiedDate;
                    worksheet.Cell(rowIndex, 9).Value = user.LastModifiedBy;

                    if (user.LastLoginDate.HasValue)
                    {
                        worksheet.Cell(rowIndex, 10).Value = user.LastLoginDate.Value;
                    }

                    rowIndex++;
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(_usersFilePath);
            }
            finally
            {
                _fileLock.Release();
            }
        }

        public async Task<User?> GetUserByIdAsync(string userId)
        {
            var users = await GetAllUsersAsync();
            return users.FirstOrDefault(u => u.UserId == userId);
        }

        public async Task<User?> ValidateLoginAsync(string username, string password)
        {
            var users = await GetAllUsersAsync();

            var user = users.FirstOrDefault(u =>
                u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) &&
                u.Password == password &&
                u.IsActive == true
            );

            if (user != null)
            {
                // Update last login date
                user.LastLoginDate = DateTime.Now;
                await SaveUsers(users);
            }

            return user;
        }

        public async Task<bool> AddUserAsync(User newUser, string currentUsername)
        {
            var users = await GetAllUsersAsync();

            // Check if username already exists
            if (users.Any(u => u.Username.Equals(newUser.Username, StringComparison.OrdinalIgnoreCase)))
            {
                return false; // Username already exists
            }

            // Generate next UserId
            newUser.UserId = GenerateNextUserId(users);
            newUser.CreatedDate = DateTime.Now;
            newUser.LastModifiedDate = DateTime.Now;
            newUser.LastModifiedBy = currentUsername;

            users.Add(newUser);
            await SaveUsers(users);

            return true;
        }

        public async Task<bool> UpdateUserAsync(User updatedUser, string currentUsername)
        {
            var users = await GetAllUsersAsync();
            var existingUser = users.FirstOrDefault(u => u.UserId == updatedUser.UserId);

            if (existingUser == null)
            {
                return false;
            }

            // Check if new username conflicts with another user
            if (users.Any(u => u.UserId != updatedUser.UserId &&
                              u.Username.Equals(updatedUser.Username, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            // Update fields
            existingUser.Username = updatedUser.Username;
            existingUser.Password = updatedUser.Password;
            existingUser.FullName = updatedUser.FullName;
            existingUser.Role = updatedUser.Role;
            existingUser.IsActive = updatedUser.IsActive;
            existingUser.LastModifiedDate = DateTime.Now;
            existingUser.LastModifiedBy = currentUsername;

            await SaveUsers(users);
            return true;
        }

        public async Task<bool> DeleteUserAsync(string userId)
        {
            var users = await GetAllUsersAsync();
            var userToDelete = users.FirstOrDefault(u => u.UserId == userId);

            if (userToDelete == null)
            {
                return false;
            }

            // Prevent deleting the last admin
            if (userToDelete.Role == UserRoles.Admin &&
                users.Count(u => u.Role == UserRoles.Admin) <= 1)
            {
                return false; // Cannot delete last admin
            }

            users.Remove(userToDelete);
            await SaveUsers(users);

            return true;
        }

        private string GenerateNextUserId(List<User> users)
        {
            if (users.Count == 0)
                return "USER001";

            // Get highest number
            var maxId = users
                .Select(u => u.UserId)
                .Where(id => id.StartsWith("USER") && id.Length == 7)
                .Select(id =>
                {
                    if (int.TryParse(id.Substring(4), out int num))
                        return num;
                    return 0;
                })
                .DefaultIfEmpty(0)
                .Max();

            var nextId = maxId + 1;
            return $"USER{nextId:D3}"; // USER001, USER002, etc.
        }

        public async Task<string> CreateBackupAsync(string currentUsername)
        {
            var users = await GetAllUsersAsync();

            var backup = new
            {
                backupDate = DateTime.Now,
                backupBy = currentUsername,
                recordCount = users.Count,
                users = users
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            var json = JsonSerializer.Serialize(backup, options);

            var fileName = $"Users_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            var filePath = Path.Combine(_backupFolder, fileName);

            await File.WriteAllTextAsync(filePath, json);

            // Clean up old backups (keep last 10)
            CleanupOldBackups();

            return fileName;
        }

        private void CleanupOldBackups()
        {
            var backupFiles = Directory.GetFiles(_backupFolder, "Users_Backup_*.json")
                                       .OrderByDescending(f => f)
                                       .Skip(10)
                                       .ToList();

            foreach (var file in backupFiles)
            {
                try
                {
                    File.Delete(file);
                }
                catch
                {
                    // Ignore errors when deleting old backups
                }
            }
        }

        public async Task<bool> RestoreFromBackupAsync(string backupFileName)
        {
            var backupPath = Path.Combine(_backupFolder, backupFileName);

            if (!File.Exists(backupPath))
            {
                return false;
            }

            try
            {
                var json = await File.ReadAllTextAsync(backupPath);
                var backup = JsonSerializer.Deserialize<BackupData>(json);

                if (backup?.users != null)
                {
                    await SaveUsers(backup.users);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private class BackupData
        {
            public List<User> users { get; set; } = new();
        }
    }
}