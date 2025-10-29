using ClosedXML.Excel;
using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    public class BatchService : IBatchService
    {
        private readonly string _filePath = "Data/Batches.xlsx";
        private readonly RecipeService _recipeService;

        public BatchService(RecipeService recipeService)
        {
            _recipeService = recipeService;
            EnsureFileExists();
        }

        private void EnsureFileExists()
        {
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(_filePath))
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Batches");

                // Create headers
                worksheet.Cell(1, 1).Value = "BatchId";
                worksheet.Cell(1, 2).Value = "RecipeId";
                worksheet.Cell(1, 3).Value = "RecipeName";
                worksheet.Cell(1, 4).Value = "TotalRepetitions";
                worksheet.Cell(1, 5).Value = "CurrentRepetition";
                worksheet.Cell(1, 6).Value = "Status";
                worksheet.Cell(1, 7).Value = "PlannedStartTime";
                worksheet.Cell(1, 8).Value = "PlannedEndTime";
                worksheet.Cell(1, 9).Value = "CreatedBy";
                worksheet.Cell(1, 10).Value = "CreatedDate";
                worksheet.Cell(1, 11).Value = "StartedBy";
                worksheet.Cell(1, 12).Value = "StartedDate";
                worksheet.Cell(1, 13).Value = "CompletedDate";
                worksheet.Cell(1, 14).Value = "AbortReason";
                worksheet.Cell(1, 15).Value = "AbortedBy";
                worksheet.Cell(1, 16).Value = "AbortedDate";
                worksheet.Cell(1, 17).Value = "Notes";

                // Style headers
                var headerRange = worksheet.Range(1, 1, 1, 17);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                workbook.SaveAs(_filePath);
            }
        }

        public async Task<List<Batch>> GetAllBatchesAsync()
        {
            return await Task.Run(() =>
            {
                var batches = new List<Batch>();

                try
                {
                    using (var workbook = new XLWorkbook(_filePath))
                    {
                        var worksheet = workbook.Worksheet("Batches");
                        if (worksheet == null)
                        {
                            Console.WriteLine("Batches worksheet not found");
                            return batches;
                        }

                        var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1);
                        if (rows == null) return batches;

                        foreach (var row in rows)
                        {
                            try
                            {
                                batches.Add(new Batch
                                {
                                    BatchId = row.Cell(1).GetString(),
                                    RecipeId = row.Cell(2).GetString(),
                                    RecipeName = row.Cell(3).GetString(),
                                    TotalRepetitions = ParseInt(row.Cell(4)),
                                    CurrentRepetition = ParseInt(row.Cell(5)),
                                    Status = row.Cell(6).GetString(),
                                    PlannedStartTime = ParseNullableDateTime(row.Cell(7)),
                                    PlannedEndTime = ParseNullableDateTime(row.Cell(8)),
                                    CreatedBy = row.Cell(9).GetString(),
                                    CreatedDate = ParseDateTime(row.Cell(10)),
                                    StartedBy = row.Cell(11).IsEmpty() ? null : row.Cell(11).GetString(),
                                    StartedDate = ParseNullableDateTime(row.Cell(12)),
                                    CompletedDate = ParseNullableDateTime(row.Cell(13)),
                                    AbortReason = row.Cell(14).IsEmpty() ? null : row.Cell(14).GetString(),
                                    AbortedBy = row.Cell(15).IsEmpty() ? null : row.Cell(15).GetString(),
                                    AbortedDate = ParseNullableDateTime(row.Cell(16)),
                                    Notes = row.Cell(17).IsEmpty() ? null : row.Cell(17).GetString()
                                });
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error parsing batch row: {ex.Message}");
                            }
                        }
                    } // Workbook is disposed here
                }
                catch (IOException ioEx)
                {
                    Console.WriteLine($"File access error: {ioEx.Message}");
                    throw new InvalidOperationException("Cannot access the batch file. Make sure it's not open in Excel.", ioEx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading batches: {ex.Message}");
                    throw;
                }

                return batches;
            });
        }

        // Helper methods for safe parsing
        private int ParseInt(IXLCell cell)
        {
            if (cell.IsEmpty()) return 0;

            var value = cell.GetString();
            if (int.TryParse(value, out int result))
                return result;

            // Try getting as double first (Excel sometimes stores numbers as doubles)
            if (cell.TryGetValue(out double doubleValue))
                return (int)doubleValue;

            return 0;
        }

        private DateTime ParseDateTime(IXLCell cell)
        {
            if (cell.IsEmpty()) return DateTime.Now;

            // Try to get as DateTime directly
            if (cell.TryGetValue(out DateTime dateValue))
                return dateValue;

            // Try parsing string
            var value = cell.GetString();
            if (DateTime.TryParse(value, out DateTime result))
                return result;

            return DateTime.Now;
        }

        private DateTime? ParseNullableDateTime(IXLCell cell)
        {
            if (cell.IsEmpty()) return null;

            // Try to get as DateTime directly
            if (cell.TryGetValue(out DateTime dateValue))
                return dateValue;

            // Try parsing string
            var value = cell.GetString();
            if (DateTime.TryParse(value, out DateTime result))
                return result;

            return null;
        }

        public async Task<Batch?> GetBatchByIdAsync(string batchId)
        {
            var batches = await GetAllBatchesAsync();
            return batches.FirstOrDefault(b => b.BatchId == batchId);
        }

        public async Task<List<Batch>> GetBatchesByStatusAsync(string status)
        {
            var batches = await GetAllBatchesAsync();
            return batches.Where(b => b.Status == status).ToList();
        }

        public async Task<List<Batch>> GetActiveBatchesAsync()
        {
            return await GetBatchesByStatusAsync("InProgress");
        }

        public async Task<List<Batch>> GetPendingBatchesAsync()
        {
            return await GetBatchesByStatusAsync("Pending");
        }

        public async Task<string> CreateBatchAsync(Batch batch)
        {
            return await Task.Run(async () =>
            {
                // Generate new Batch ID
                var batches = await GetAllBatchesAsync();
                int maxId = 0;
                foreach (var b in batches)
                {
                    if (b.BatchId.StartsWith("BATCH"))
                    {
                        var numPart = b.BatchId.Substring(5);
                        if (int.TryParse(numPart, out int id))
                        {
                            maxId = Math.Max(maxId, id);
                        }
                    }
                }
                batch.BatchId = $"BATCH{(maxId + 1).ToString("D3")}";

                // Get recipe name
                var recipe = await _recipeService.GetRecipeByIdAsync(batch.RecipeId);
                if (recipe != null)
                {
                    batch.RecipeName = recipe.RecipeName;
                }

                // Set defaults
                batch.Status = "Pending";
                batch.CurrentRepetition = 0;
                batch.CreatedDate = DateTime.Now;

                using var workbook = new XLWorkbook(_filePath);
                var worksheet = workbook.Worksheet("Batches");
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                var newRow = lastRow + 1;

                worksheet.Cell(newRow, 1).Value = batch.BatchId;
                worksheet.Cell(newRow, 2).Value = batch.RecipeId;
                worksheet.Cell(newRow, 3).Value = batch.RecipeName;
                worksheet.Cell(newRow, 4).Value = batch.TotalRepetitions;
                worksheet.Cell(newRow, 5).Value = batch.CurrentRepetition;
                worksheet.Cell(newRow, 6).Value = batch.Status;
                worksheet.Cell(newRow, 7).Value = batch.PlannedStartTime.HasValue ? batch.PlannedStartTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cell(newRow, 8).Value = batch.PlannedEndTime.HasValue ? batch.PlannedEndTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cell(newRow, 9).Value = batch.CreatedBy;
                worksheet.Cell(newRow, 10).Value = batch.CreatedDate;
                worksheet.Cell(newRow, 11).Value = batch.StartedBy ?? "";
                worksheet.Cell(newRow, 12).Value = batch.StartedDate.HasValue ? batch.StartedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cell(newRow, 13).Value = batch.CompletedDate.HasValue ? batch.CompletedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cell(newRow, 14).Value = batch.AbortReason ?? "";
                worksheet.Cell(newRow, 15).Value = batch.AbortedBy ?? "";
                worksheet.Cell(newRow, 16).Value = batch.AbortedDate.HasValue ? batch.AbortedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cell(newRow, 17).Value = batch.Notes ?? "";

                workbook.Save();

                return batch.BatchId;
            });
        }

        public async Task<bool> UpdateBatchAsync(Batch batch)
        {
            return await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(_filePath);
                var worksheet = workbook.Worksheet("Batches");
                var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1);

                if (rows == null) return false;

                foreach (var row in rows)
                {
                    if (row.Cell(1).GetString() == batch.BatchId)
                    {
                        row.Cell(2).Value = batch.RecipeId;
                        row.Cell(3).Value = batch.RecipeName;
                        row.Cell(4).Value = batch.TotalRepetitions;
                        row.Cell(5).Value = batch.CurrentRepetition;
                        row.Cell(6).Value = batch.Status;
                        row.Cell(7).Value = batch.PlannedStartTime.HasValue ? batch.PlannedStartTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                        row.Cell(8).Value = batch.PlannedEndTime.HasValue ? batch.PlannedEndTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                        row.Cell(9).Value = batch.CreatedBy;
                        row.Cell(10).Value = batch.CreatedDate;
                        row.Cell(11).Value = batch.StartedBy ?? "";
                        row.Cell(12).Value = batch.StartedDate.HasValue ? batch.StartedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                        row.Cell(13).Value = batch.CompletedDate.HasValue ? batch.CompletedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                        row.Cell(14).Value = batch.AbortReason ?? "";
                        row.Cell(15).Value = batch.AbortedBy ?? "";
                        row.Cell(16).Value = batch.AbortedDate.HasValue ? batch.AbortedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                        row.Cell(17).Value = batch.Notes ?? "";

                        workbook.Save();
                        return true;
                    }
                }

                return false;
            });
        }

        public async Task<bool> DeleteBatchAsync(string batchId)
        {
            return await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(_filePath);
                var worksheet = workbook.Worksheet("Batches");
                var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1);

                if (rows == null) return false;

                foreach (var row in rows)
                {
                    if (row.Cell(1).GetString() == batchId)
                    {
                        row.Delete();
                        workbook.Save();
                        return true;
                    }
                }

                return false;
            });
        }

        public async Task<bool> StartBatchAsync(string batchId, string startedBy)
        {
            var batch = await GetBatchByIdAsync(batchId);
            if (batch == null || batch.Status != "Pending") return false;

            // Check if we can start (max 5 active batches)
            if (!await CanStartBatch()) return false;

            batch.Status = "InProgress";
            batch.StartedBy = startedBy;
            batch.StartedDate = DateTime.Now;
            batch.CurrentRepetition = 0;

            return await UpdateBatchAsync(batch);
        }

        public async Task<bool> CompleteBatchAsync(string batchId)
        {
            var batch = await GetBatchByIdAsync(batchId);
            if (batch == null || batch.Status != "InProgress") return false;

            batch.Status = "Completed";
            batch.CompletedDate = DateTime.Now;
            batch.CurrentRepetition = batch.TotalRepetitions;

            return await UpdateBatchAsync(batch);
        }

        public async Task<bool> AbortBatchAsync(string batchId, string abortedBy, string abortReason)
        {
            var batch = await GetBatchByIdAsync(batchId);
            if (batch == null) return false;

            batch.Status = "Aborted";
            batch.AbortedBy = abortedBy;
            batch.AbortedDate = DateTime.Now;
            batch.AbortReason = abortReason;

            return await UpdateBatchAsync(batch);
        }

        public async Task<bool> UpdateRepetitionProgressAsync(string batchId, int currentRepetition)
        {
            var batch = await GetBatchByIdAsync(batchId);
            if (batch == null) return false;

            batch.CurrentRepetition = currentRepetition;

            // Auto-complete if all repetitions done
            if (currentRepetition >= batch.TotalRepetitions)
            {
                batch.Status = "Completed";
                batch.CompletedDate = DateTime.Now;
            }

            return await UpdateBatchAsync(batch);
        }

        public async Task<int> GetActiveBatchCountAsync()
        {
            var activeBatches = await GetActiveBatchesAsync();
            return activeBatches.Count;
        }

        public async Task<bool> CanStartBatch()
        {
            var activeCount = await GetActiveBatchCountAsync();
            return activeCount < 5;
        }
    }
}