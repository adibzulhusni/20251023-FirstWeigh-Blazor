using FirstWeigh.Models;
using ClosedXML.Excel;

namespace FirstWeigh.Services
{
    public class IngredientService
    {
        private readonly IConfiguration _configuration;
        private readonly string _ingredientsFilePath;
        private static readonly SemaphoreSlim _fileLock = new(1, 1);
        private DateTime _lastLoadedTime;

        public IngredientService(IConfiguration configuration)
        {
            _configuration = configuration;
            _ingredientsFilePath = _configuration["DataStorage:IngredientsFilePath"] ?? "Data/Ingredients.xlsx";

            // Ensure directory exists
            var directory = Path.GetDirectoryName(_ingredientsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        public async Task<List<Ingredient>> GetAllIngredientsAsync()
        {
            await _fileLock.WaitAsync();
            try
            {
                if (!File.Exists(_ingredientsFilePath))
                {
                    return new List<Ingredient>();
                }

                var ingredients = new List<Ingredient>();

                using var workbook = new XLWorkbook(_ingredientsFilePath);
                var worksheet = workbook.Worksheet(1);

                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                if (lastRow <= 1)
                {
                    return ingredients;
                }

                // Read from row 2 (skip header)
                for (int row = 2; row <= lastRow; row++)
                {
                    // Skip completely empty rows
                    if (worksheet.Cell(row, 1).IsEmpty() && worksheet.Cell(row, 2).IsEmpty())
                        continue;

                    var ingredient = new Ingredient
                    {
                        IngredientId = worksheet.Cell(row, 1).GetString().Trim(),
                        IngredientCode = worksheet.Cell(row, 2).GetString().Trim(),
                        IngredientName = worksheet.Cell(row, 3).GetString().Trim(),
                        PackingType = worksheet.Cell(row, 4).GetValue<string>(),  // Changed this line
                        UnitOfMeasure = worksheet.Cell(row, 5).GetString().Trim(),
                        CreatedDate = worksheet.Cell(row, 6).IsEmpty() ? DateTime.Now : worksheet.Cell(row, 6).GetDateTime(),
                        LastModifiedDate = worksheet.Cell(row, 7).IsEmpty() ? DateTime.Now : worksheet.Cell(row, 7).GetDateTime(),
                        LastModifiedBy = worksheet.Cell(row, 8).GetString().Trim()
                    };

                    ingredients.Add(ingredient);
                }

                _lastLoadedTime = DateTime.Now;
                return ingredients;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel: {ex.Message}");
                throw;
            }
            finally
            {
                _fileLock.Release();
            }
        }

        public string GetFilePath()
        {
            return Path.GetFullPath(_ingredientsFilePath);
        }

        public DateTime GetLastLoadedTime()
        {
            return _lastLoadedTime;
        }

        public void OpenExcelFile()
        {
            // Get the absolute path to the actual file location
            var fullPath = Path.GetFullPath(_ingredientsFilePath);

            if (File.Exists(fullPath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = fullPath,
                    UseShellExecute = true
                });
            }
            else
            {
                throw new FileNotFoundException($"Excel file not found at: {fullPath}");
            }
        }
    }
}