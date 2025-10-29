using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using FirstWeigh.Models;
using System.Diagnostics;

namespace FirstWeigh.Services
{
    public class BowlService
    {
        private readonly string _filePath;

        public BowlService(IConfiguration configuration)
        {
            // Get the base directory where the app is running
            var baseDirectory = AppContext.BaseDirectory;

            // Get configured path (relative)
            var configuredPath = configuration["DataStorage:BowlsFilePath"] ?? "Data/Bowls.xlsx";

            // Combine to get absolute path
            _filePath = Path.Combine(baseDirectory, configuredPath);

            Console.WriteLine($"🔧 Bowl Service initialized");
            Console.WriteLine($"📂 Base Directory: {baseDirectory}");
            Console.WriteLine($"📄 Bowl File Path: {_filePath}");

            // Ensure directory exists
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
                Console.WriteLine($"✅ Created directory: {directory}");
            }

            // Create file if it doesn't exist
            if (!File.Exists(_filePath))
            {
                Console.WriteLine($"📝 Creating new Excel file...");
                CreateEmptyExcelFile();
                Console.WriteLine($"✅ Excel file created: {_filePath}");
            }
            else
            {
                Console.WriteLine($"✅ Excel file exists: {_filePath}");
            }
        }

        private void CreateEmptyExcelFile()
        {
            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Bowls");

                // Headers
                worksheet.Cell(1, 1).Value = "BowlId";
                worksheet.Cell(1, 2).Value = "BowlCode";
                worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(1, 4).Value = "BowlType";
                worksheet.Cell(1, 5).Value = "Weight";
                worksheet.Cell(1, 6).Value = "Status";
                worksheet.Cell(1, 7).Value = "CurrentLocation";
                worksheet.Cell(1, 8).Value = "CreatedDate";
                worksheet.Cell(1, 9).Value = "LastModifiedDate";
                worksheet.Cell(1, 10).Value = "LastModifiedBy";
                worksheet.Cell(1, 11).Value = "Remarks";

                // Style headers
                var headerRange = worksheet.Range(1, 1, 1, 11);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                workbook.SaveAs(_filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating Excel file: {ex.Message}");
                throw;
            }
        }

        public async Task<List<Bowl>> GetAllBowlsAsync()
        {
            return await Task.Run(() =>
            {
                var bowls = new List<Bowl>();

                if (!File.Exists(_filePath))
                {
                    Console.WriteLine($"⚠️ Excel file not found, creating new one: {_filePath}");
                    CreateEmptyExcelFile();
                    return bowls;
                }

                try
                {
                    using var workbook = new XLWorkbook(_filePath);
                    var worksheet = workbook.Worksheet("Bowls");

                    var rows = worksheet.RowsUsed().Skip(1); // Skip header

                    foreach (var row in rows)
                    {
                        try
                        {
                            var bowl = new Bowl
                            {
                                BowlId = row.Cell(1).GetString(),
                                BowlCode = row.Cell(2).GetString(),
                                Category = row.Cell(3).GetString(),
                                BowlType = row.Cell(4).GetString(),
                                Weight = row.Cell(5).GetValue<decimal>(),
                                Status = row.Cell(6).GetString(),
                                CurrentLocation = row.Cell(7).GetString(),
                                CreatedDate = row.Cell(8).GetDateTime(),
                                LastModifiedDate = row.Cell(9).GetDateTime(),
                                LastModifiedBy = row.Cell(10).GetString(),
                                Remarks = row.Cell(11).GetString()
                            };

                            bowls.Add(bowl);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error reading bowl row: {ex.Message}");
                        }
                    }

                    Console.WriteLine($"✅ Loaded {bowls.Count} bowls from Excel");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading bowls: {ex.Message}");
                }

                return bowls;
            });
        }

        public async Task SaveBowlAsync(Bowl bowl)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_filePath);
                    var worksheet = workbook.Worksheet("Bowls");

                    // Find next empty row
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    var newRow = lastRow + 1;

                    // Write bowl data
                    worksheet.Cell(newRow, 1).Value = bowl.BowlId;
                    worksheet.Cell(newRow, 2).Value = bowl.BowlCode;
                    worksheet.Cell(newRow, 3).Value = bowl.Category;
                    worksheet.Cell(newRow, 4).Value = bowl.BowlType;
                    worksheet.Cell(newRow, 5).Value = bowl.Weight;
                    worksheet.Cell(newRow, 6).Value = bowl.Status;
                    worksheet.Cell(newRow, 7).Value = bowl.CurrentLocation;
                    worksheet.Cell(newRow, 8).Value = bowl.CreatedDate;
                    worksheet.Cell(newRow, 9).Value = bowl.LastModifiedDate;
                    worksheet.Cell(newRow, 10).Value = bowl.LastModifiedBy;
                    worksheet.Cell(newRow, 11).Value = bowl.Remarks;

                    workbook.Save();
                    Console.WriteLine($"✅ Bowl {bowl.BowlCode} saved to Excel (Row {newRow})");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error saving bowl: {ex.Message}");
                    throw;
                }
            });
        }

        public async Task DeleteBowlAsync(string bowlId)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_filePath);
                    var worksheet = workbook.Worksheet("Bowls");

                    // Find the row with matching BowlId
                    var rows = worksheet.RowsUsed().Skip(1); // Skip header

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString() == bowlId)
                        {
                            row.Delete();
                            workbook.Save();
                            Console.WriteLine($"✅ Bowl {bowlId} deleted from Excel");
                            return;
                        }
                    }

                    Console.WriteLine($"⚠️ Bowl {bowlId} not found in Excel");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error deleting bowl: {ex.Message}");
                    throw;
                }
            });
        }

        public void OpenExcelFile()
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    Console.WriteLine($"📂 Opening Excel file: {_filePath}");
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = _filePath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    Console.WriteLine($"❌ Excel file not found: {_filePath}");
                    throw new FileNotFoundException($"Excel file not found: {_filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error opening Excel file: {ex.Message}");
                throw;
            }
        }

    }
}
