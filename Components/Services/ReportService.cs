// ReportService.txt

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using FirstWeigh.Models;
using System.Runtime.Intrinsics.Arm;
using System.Text.Json;

namespace FirstWeigh.Services
{
    public class ReportService
    {
        private readonly string _recordsFilePath;
        private readonly string _detailsFilePath;
        private readonly string _jsonBackupFolder;
        private static readonly SemaphoreSlim _fileLock = new(1, 1);

        public ReportService(IConfiguration configuration)
        {
            var baseDirectory = AppContext.BaseDirectory;

            _recordsFilePath = Path.Combine(baseDirectory,
                configuration["DataStorage:WeighingRecordsFilePath"] ?? "Data/WeighingRecords.xlsx");

            _detailsFilePath = Path.Combine(baseDirectory,
                configuration["DataStorage:WeighingDetailsFilePath"] ?? "Data/WeighingDetails.xlsx");

            _jsonBackupFolder = Path.Combine(baseDirectory,
                configuration["DataStorage:JsonBackupFolder"] ?? "Data/JsonBackups");

            Console.WriteLine($"🔧 Report Service initialized");
            Console.WriteLine($"📄 Records File: {_recordsFilePath}");
            Console.WriteLine($"📄 Details File: {_detailsFilePath}");
            Console.WriteLine($"📁 JSON Backup Folder: {_jsonBackupFolder}");

            EnsureFilesExist();
        }

        private void EnsureFilesExist()
        {
            // Ensure directories exist
            var recordsDir = Path.GetDirectoryName(_recordsFilePath);
            if (!string.IsNullOrEmpty(recordsDir) && !Directory.Exists(recordsDir))
            {
                Directory.CreateDirectory(recordsDir);
            }

            if (!Directory.Exists(_jsonBackupFolder))
            {
                Directory.CreateDirectory(_jsonBackupFolder);
            }

            // Create WeighingRecords file if not exists
            if (!File.Exists(_recordsFilePath))
            {
                CreateWeighingRecordsFile();
            }

            // Create WeighingDetails file if not exists
            if (!File.Exists(_detailsFilePath))
            {
                CreateWeighingDetailsFile();
            }
        }

        private void CreateWeighingRecordsFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("WeighingRecords");

            // Headers
            worksheet.Cell(1, 1).Value = "RecordId";
            worksheet.Cell(1, 2).Value = "BatchId";
            worksheet.Cell(1, 3).Value = "RecipeId";
            worksheet.Cell(1, 4).Value = "RecipeCode";
            worksheet.Cell(1, 5).Value = "RecipeName";
            worksheet.Cell(1, 6).Value = "OperatorName";
            worksheet.Cell(1, 7).Value = "SessionStartTime";
            worksheet.Cell(1, 8).Value = "SessionEndTime";
            worksheet.Cell(1, 9).Value = "TotalRepetitions";
            worksheet.Cell(1, 10).Value = "CompletedRepetitions";
            worksheet.Cell(1, 11).Value = "Status";
            worksheet.Cell(1, 12).Value = "AbortReason";
            worksheet.Cell(1, 13).Value = "AbortedBy";
            worksheet.Cell(1, 14).Value = "AbortedDate";
            worksheet.Cell(1, 15).Value = "TotalIngredientsWeighed";
            worksheet.Cell(1, 16).Value = "IngredientsWithinTolerance";
            worksheet.Cell(1, 17).Value = "IngredientsOutOfTolerance";
            worksheet.Cell(1, 18).Value = "AverageDeviation";
            worksheet.Cell(1, 19).Value = "MaxDeviation";
            worksheet.Cell(1, 20).Value = "CreatedDate";
            worksheet.Cell(1, 21).Value = "CreatedBy";

            // Style headers
            var headerRange = worksheet.Range(1, 1, 1, 21);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_recordsFilePath);
            Console.WriteLine($"✅ Created WeighingRecords file: {_recordsFilePath}");
        }

        private void CreateWeighingDetailsFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("WeighingDetails");

            // Headers
            worksheet.Cell(1, 1).Value = "DetailId";
            worksheet.Cell(1, 2).Value = "RecordId";
            worksheet.Cell(1, 3).Value = "BatchId";
            worksheet.Cell(1, 4).Value = "RepetitionNumber";
            worksheet.Cell(1, 5).Value = "IngredientSequence";
            worksheet.Cell(1, 6).Value = "IngredientId";
            worksheet.Cell(1, 7).Value = "IngredientCode";
            worksheet.Cell(1, 8).Value = "IngredientName";
            worksheet.Cell(1, 9).Value = "TargetWeight";
            worksheet.Cell(1, 10).Value = "ActualWeight";
            worksheet.Cell(1, 11).Value = "Deviation";
            worksheet.Cell(1, 12).Value = "MinWeight";
            worksheet.Cell(1, 13).Value = "MaxWeight";
            worksheet.Cell(1, 14).Value = "ToleranceValue";
            worksheet.Cell(1, 15).Value = "IsWithinTolerance";
            worksheet.Cell(1, 16).Value = "BowlCode";
            worksheet.Cell(1, 17).Value = "BowlType";
            worksheet.Cell(1, 18).Value = "ScaleNumber";
            worksheet.Cell(1, 19).Value = "Unit";
            worksheet.Cell(1, 20).Value = "Timestamp";

            // Style headers
            var headerRange = worksheet.Range(1, 1, 1, 20);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_detailsFilePath);
            Console.WriteLine($"✅ Created WeighingDetails file: {_detailsFilePath}");
        }

        // ✅ START NEW WEIGHING RECORD
        // Line 138-172
        // ✅ Complete method - Line 138-172 in ReportService.cs
        public async Task<string> StartWeighingRecordAsync(WeighingSession session)
        {
            await _fileLock.WaitAsync();
            try
            {
                // ✅ Read existing records (inside lock)
                var existingRecords = new List<WeighingRecord>();

                using (var workbook = new XLWorkbook(_recordsFilePath))
                {
                    var worksheet = workbook.Worksheet("WeighingRecords");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        try
                        {
                            existingRecords.Add(ParseRecordRow(row));
                        }
                        catch { }
                    }
                }

                // ✅ Generate next ID (inside lock)
                var recordId = GenerateNextRecordId(existingRecords);

                // ✅ Create record
                var record = new WeighingRecord
                {
                    RecordId = recordId,
                    BatchId = session.BatchId,
                    RecipeId = session.RecipeId,
                    RecipeCode = session.RecipeId,
                    RecipeName = session.RecipeName,  // ← Fixed from before!
                    OperatorName = session.OperatorName,
                    SessionStartTime = session.SessionStarted ?? DateTime.Now,
                    TotalRepetitions = session.TotalRepetitions,
                    CompletedRepetitions = 0,
                    Status = WeighingRecordStatus.InProgress,
                    CreatedDate = DateTime.Now,
                    CreatedBy = session.OperatorName
                };

                // ✅ Save immediately (inside lock)
                using (var workbook = new XLWorkbook(_recordsFilePath))
                {
                    var worksheet = workbook.Worksheet("WeighingRecords");
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = record.RecordId;
                    worksheet.Cell(newRow, 2).Value = record.BatchId;
                    worksheet.Cell(newRow, 3).Value = record.RecipeId;
                    worksheet.Cell(newRow, 4).Value = record.RecipeCode;
                    worksheet.Cell(newRow, 5).Value = record.RecipeName;
                    worksheet.Cell(newRow, 6).Value = record.OperatorName;
                    worksheet.Cell(newRow, 7).Value = record.SessionStartTime;
                    worksheet.Cell(newRow, 8).Value = record.SessionEndTime;
                    worksheet.Cell(newRow, 9).Value = record.TotalRepetitions;
                    worksheet.Cell(newRow, 10).Value = record.CompletedRepetitions;
                    worksheet.Cell(newRow, 11).Value = record.Status;
                    worksheet.Cell(newRow, 12).Value = record.AbortReason ?? "";
                    worksheet.Cell(newRow, 13).Value = record.AbortedBy ?? "";
                    worksheet.Cell(newRow, 14).Value = record.AbortedDate;
                    worksheet.Cell(newRow, 15).Value = record.TotalIngredientsWeighed;
                    worksheet.Cell(newRow, 16).Value = record.IngredientsWithinTolerance;
                    worksheet.Cell(newRow, 17).Value = record.IngredientsOutOfTolerance;
                    worksheet.Cell(newRow, 18).Value = record.AverageDeviation;
                    worksheet.Cell(newRow, 19).Value = record.MaxDeviation;
                    worksheet.Cell(newRow, 20).Value = record.CreatedDate;
                    worksheet.Cell(newRow, 21).Value = record.CreatedBy;

                    workbook.Save();
                }

                Console.WriteLine($"✅ Started weighing record: {recordId}");
                return recordId;
            }
            finally
            {
                _fileLock.Release();
            }
        }

        // ✅ Add internal method that doesn't use the lock (called from within locked sections)
        private async Task<List<WeighingRecord>> GetAllRecordsInternalAsync()
        {
            // Same code as GetAllRecordsAsync but without taking the lock
            return await Task.Run(() =>
            {
                var records = new List<WeighingRecord>();
                try
                {
                    using var workbook = new XLWorkbook(_recordsFilePath);
                    var worksheet = workbook.Worksheet("WeighingRecords");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        try
                        {
                            records.Add(ParseRecordRow(row));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error parsing record row: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading records: {ex.Message}");
                }
                return records;
            });
        }

        // ✅ SAVE INGREDIENT DETAIL (called after each ingredient)
        public async Task SaveIngredientDetailAsync(string recordId, WeighingDetail detail)
        {
            await _fileLock.WaitAsync();
            try
            {
                // Generate DetailId
                var existingDetails = await GetDetailsByRecordIdAsync(recordId);
                detail.DetailId = $"DETAIL{(existingDetails.Count + 1):D4}";
                detail.RecordId = recordId;

                using var workbook = new XLWorkbook(_detailsFilePath);
                var worksheet = workbook.Worksheet("WeighingDetails");
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                var newRow = lastRow + 1;

                worksheet.Cell(newRow, 1).Value = detail.DetailId;
                worksheet.Cell(newRow, 2).Value = detail.RecordId;
                worksheet.Cell(newRow, 3).Value = detail.BatchId;
                worksheet.Cell(newRow, 4).Value = detail.RepetitionNumber;
                worksheet.Cell(newRow, 5).Value = detail.IngredientSequence;
                worksheet.Cell(newRow, 6).Value = detail.IngredientId;
                worksheet.Cell(newRow, 7).Value = detail.IngredientCode;
                worksheet.Cell(newRow, 8).Value = detail.IngredientName;
                worksheet.Cell(newRow, 9).Value = detail.TargetWeight;
                worksheet.Cell(newRow, 10).Value = detail.ActualWeight;
                worksheet.Cell(newRow, 11).Value = detail.Deviation;
                worksheet.Cell(newRow, 12).Value = detail.MinWeight;
                worksheet.Cell(newRow, 13).Value = detail.MaxWeight;
                worksheet.Cell(newRow, 14).Value = detail.ToleranceValue;
                worksheet.Cell(newRow, 15).Value = detail.IsWithinTolerance;
                worksheet.Cell(newRow, 16).Value = detail.BowlCode;
                worksheet.Cell(newRow, 17).Value = detail.BowlType;
                worksheet.Cell(newRow, 18).Value = detail.ScaleNumber;
                worksheet.Cell(newRow, 19).Value = detail.Unit;
                worksheet.Cell(newRow, 20).Value = detail.Timestamp;

                workbook.Save();
                Console.WriteLine($"✅ Saved ingredient detail: {detail.IngredientCode} for {recordId}");
            }
            finally
            {
                _fileLock.Release();
            }
        }

        // ✅ FINALIZE REPORT (called when batch completes)
        public async Task<bool> FinalizeReportAsync(string recordId, bool isAborted = false,
            string? abortReason = null, string? abortedBy = null)
        {
            await _fileLock.WaitAsync();
            try
            {
                var record = await GetRecordByIdAsync(recordId);
                if (record == null) return false;

                // Load all details for quality metrics
                var details = await GetDetailsByRecordIdAsync(recordId);

                // Update record with final data
                record.SessionEndTime = DateTime.Now;
                record.Status = isAborted ? WeighingRecordStatus.Aborted : WeighingRecordStatus.Completed;
                record.CompletedRepetitions = details.Select(d => d.RepetitionNumber).Distinct().Count();

                if (isAborted)
                {
                    record.AbortReason = abortReason;
                    record.AbortedBy = abortedBy;
                    record.AbortedDate = DateTime.Now;
                }

                // Calculate quality metrics
                record.TotalIngredientsWeighed = details.Count;
                record.IngredientsWithinTolerance = details.Count(d => d.IsWithinTolerance);
                record.IngredientsOutOfTolerance = details.Count(d => !d.IsWithinTolerance);
                record.AverageDeviation = details.Any()
                    ? details.Average(d => Math.Abs(d.Deviation))
                    : 0;
                record.MaxDeviation = details.Any()
                    ? details.Max(d => Math.Abs(d.Deviation))
                    : 0;

                // Update record in Excel
                await UpdateRecordAsync(record);

                // Create JSON backup
                await CreateJsonBackupAsync(record, details);

                Console.WriteLine($"✅ Finalized report: {recordId} - Status: {record.Status}");
                return true;
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private async Task SaveRecordAsync(WeighingRecord record)
        {
            using var workbook = new XLWorkbook(_recordsFilePath);
            var worksheet = workbook.Worksheet("WeighingRecords");
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var newRow = lastRow + 1;

            worksheet.Cell(newRow, 1).Value = record.RecordId;
            worksheet.Cell(newRow, 2).Value = record.BatchId;
            worksheet.Cell(newRow, 3).Value = record.RecipeId;
            worksheet.Cell(newRow, 4).Value = record.RecipeCode;
            worksheet.Cell(newRow, 5).Value = record.RecipeName;
            worksheet.Cell(newRow, 6).Value = record.OperatorName;
            worksheet.Cell(newRow, 7).Value = record.SessionStartTime;
            worksheet.Cell(newRow, 8).Value = record.SessionEndTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
            worksheet.Cell(newRow, 9).Value = record.TotalRepetitions;
            worksheet.Cell(newRow, 10).Value = record.CompletedRepetitions;
            worksheet.Cell(newRow, 11).Value = record.Status;
            worksheet.Cell(newRow, 12).Value = record.AbortReason ?? "";
            worksheet.Cell(newRow, 13).Value = record.AbortedBy ?? "";
            worksheet.Cell(newRow, 14).Value = record.AbortedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
            worksheet.Cell(newRow, 15).Value = record.TotalIngredientsWeighed;
            worksheet.Cell(newRow, 16).Value = record.IngredientsWithinTolerance;
            worksheet.Cell(newRow, 17).Value = record.IngredientsOutOfTolerance;
            worksheet.Cell(newRow, 18).Value = record.AverageDeviation;
            worksheet.Cell(newRow, 19).Value = record.MaxDeviation;
            worksheet.Cell(newRow, 20).Value = record.CreatedDate;
            worksheet.Cell(newRow, 21).Value = record.CreatedBy;

            workbook.Save();
        }

        private async Task UpdateRecordAsync(WeighingRecord record)
        {
            await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(_recordsFilePath);
                var worksheet = workbook.Worksheet("WeighingRecords");
                var rows = worksheet.RowsUsed().Skip(1);

                foreach (var row in rows)
                {
                    if (row.Cell(1).GetString() == record.RecordId)
                    {
                        row.Cell(8).Value = record.SessionEndTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
                        row.Cell(10).Value = record.CompletedRepetitions;
                        row.Cell(11).Value = record.Status;
                        row.Cell(12).Value = record.AbortReason ?? "";
                        row.Cell(13).Value = record.AbortedBy ?? "";
                        row.Cell(14).Value = record.AbortedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
                        row.Cell(15).Value = record.TotalIngredientsWeighed;
                        row.Cell(16).Value = record.IngredientsWithinTolerance;
                        row.Cell(17).Value = record.IngredientsOutOfTolerance;
                        row.Cell(18).Value = record.AverageDeviation;
                        row.Cell(19).Value = record.MaxDeviation;

                        workbook.Save();
                        return;
                    }
                }
            });
        }

        // ✅ GET METHODS
        public async Task<List<WeighingRecord>> GetAllRecordsAsync()
        {
            return await Task.Run(() =>
            {
                var records = new List<WeighingRecord>();

                try
                {
                    using var workbook = new XLWorkbook(_recordsFilePath);
                    var worksheet = workbook.Worksheet("WeighingRecords");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        try
                        {
                            records.Add(ParseRecordRow(row));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error parsing record row: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading records: {ex.Message}");
                }

                return records;
            });
        }

        public async Task<WeighingRecord?> GetRecordByIdAsync(string recordId)
        {
            var records = await GetAllRecordsAsync();
            return records.FirstOrDefault(r => r.RecordId == recordId);
        }

        public async Task<List<WeighingDetail>> GetDetailsByRecordIdAsync(string recordId)
        {
            return await Task.Run(() =>
            {
                var details = new List<WeighingDetail>();

                try
                {
                    using var workbook = new XLWorkbook(_detailsFilePath);
                    var worksheet = workbook.Worksheet("WeighingDetails");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(2).GetString() == recordId)
                        {
                            try
                            {
                                details.Add(ParseDetailRow(row));
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"⚠️ Error parsing detail row: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading details: {ex.Message}");
                }

                return details;
            });
        }

        public async Task<List<WeighingRecord>> GetRecordsByDateRangeAsync(DateTime startDate, DateTime endDate)
        {
            var allRecords = await GetAllRecordsAsync();
            return allRecords
                .Where(r => r.SessionStartTime.Date >= startDate.Date &&
                           r.SessionStartTime.Date <= endDate.Date)
                .OrderByDescending(r => r.SessionStartTime)
                .ToList();
        }

        // ✅ JSON BACKUP
        private async Task CreateJsonBackupAsync(WeighingRecord record, List<WeighingDetail> details)
        {
            try
            {
                var backup = new
                {
                    Record = record,
                    Details = details,
                    BackupDate = DateTime.Now
                };

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                var json = JsonSerializer.Serialize(backup, options);
                var fileName = $"{record.RecordId}_{record.BatchId}.json";
                var filePath = Path.Combine(_jsonBackupFolder, fileName);

                await File.WriteAllTextAsync(filePath, json);
                Console.WriteLine($"✅ JSON backup created: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Failed to create JSON backup: {ex.Message}");
            }
        }

        // ✅ HELPER METHODS
        private string GenerateNextRecordId(List<WeighingRecord> existingRecords)
        {
            if (!existingRecords.Any())
                return "RECORD001";

            var maxId = existingRecords
                .Select(r => r.RecordId)
                .Where(id => id.StartsWith("RECORD"))
                .Select(id =>
                {
                    if (int.TryParse(id.Substring(6), out int num))
                        return num;
                    return 0;
                })
                .DefaultIfEmpty(0)
                .Max();

            return $"RECORD{(maxId + 1):D3}";
        }

        private WeighingRecord ParseRecordRow(IXLRow row)
        {
            return new WeighingRecord
            {
                RecordId = row.Cell(1).GetString(),
                BatchId = row.Cell(2).GetString(),
                RecipeId = row.Cell(3).GetString(),
                RecipeCode = row.Cell(4).GetString(),
                RecipeName = row.Cell(5).GetString(),
                OperatorName = row.Cell(6).GetString(),
                SessionStartTime = row.Cell(7).GetDateTime(),
                SessionEndTime = row.Cell(8).IsEmpty() ? null : row.Cell(8).GetDateTime(),
                TotalRepetitions = row.Cell(9).GetValue<int>(),
                CompletedRepetitions = row.Cell(10).GetValue<int>(),
                Status = row.Cell(11).GetString(),
                AbortReason = row.Cell(12).IsEmpty() ? null : row.Cell(12).GetString(),
                AbortedBy = row.Cell(13).IsEmpty() ? null : row.Cell(13).GetString(),
                AbortedDate = row.Cell(14).IsEmpty() ? null : row.Cell(14).GetDateTime(),
                TotalIngredientsWeighed = row.Cell(15).GetValue<int>(),
                IngredientsWithinTolerance = row.Cell(16).GetValue<int>(),
                IngredientsOutOfTolerance = row.Cell(17).GetValue<int>(),
                AverageDeviation = row.Cell(18).GetValue<decimal>(),
                MaxDeviation = row.Cell(19).GetValue<decimal>(),
                CreatedDate = row.Cell(20).GetDateTime(),
                CreatedBy = row.Cell(21).GetString()
            };
        }

        private WeighingDetail ParseDetailRow(IXLRow row)
        {
            return new WeighingDetail
            {
                DetailId = row.Cell(1).GetString(),
                RecordId = row.Cell(2).GetString(),
                BatchId = row.Cell(3).GetString(),
                RepetitionNumber = row.Cell(4).GetValue<int>(),
                IngredientSequence = row.Cell(5).GetValue<int>(),
                IngredientId = row.Cell(6).GetString(),
                IngredientCode = row.Cell(7).GetString(),
                IngredientName = row.Cell(8).GetString(),
                TargetWeight = row.Cell(9).GetValue<decimal>(),
                ActualWeight = row.Cell(10).GetValue<decimal>(),
                MinWeight = row.Cell(12).GetValue<decimal>(),
                MaxWeight = row.Cell(13).GetValue<decimal>(),
                ToleranceValue = row.Cell(14).GetValue<decimal>(),
                BowlCode = row.Cell(16).GetString(),
                BowlType = row.Cell(17).GetString(),
                ScaleNumber = row.Cell(18).GetValue<int>(),
                Unit = row.Cell(19).GetString(),
                Timestamp = row.Cell(20).GetDateTime()
            };
        }
        // ✅ COMPLETE WEIGHING RECORD
        public async Task CompleteWeighingRecordAsync(string recordId, int completedRepetitions)
        {
            await _fileLock.WaitAsync();
            try
            {
                using (var workbook = new XLWorkbook(_recordsFilePath))
                {
                    var worksheet = workbook.Worksheet("WeighingRecords");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString() == recordId)
                        {
                            // Update to Completed
                            row.Cell(8).Value = DateTime.Now; // SessionEndTime
                            row.Cell(10).Value = completedRepetitions; // ← ADD THIS! CompletedRepetitions
                            row.Cell(11).Value = WeighingRecordStatus.Completed; // Status = "Completed"
                            break;
                        }
                    }

                    workbook.Save();
                }

                Console.WriteLine($"✅ Completed weighing record: {recordId}");
            }
            finally
            {
                _fileLock.Release();
            }
        }
    }
}