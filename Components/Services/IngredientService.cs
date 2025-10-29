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

            // Create sample Excel file if it doesn't exist
            if (!File.Exists(_ingredientsFilePath))
            {
                CreateSampleIngredientsFile();
            }
        }

        private void CreateSampleIngredientsFile()
        {
            var sampleIngredients = new List<Ingredient>
            {
                new Ingredient
                {
                    IngredientId = "ING001",
                    IngredientCode = "ING001",
                    IngredientName = "Flour",
                    PackingType = "25kg Bag",
                    UnitOfMeasure = "kg",
                    CreatedDate = DateTime.Now,
                    LastModifiedDate = DateTime.Now,
                    LastModifiedBy = "system"
                },
                new Ingredient
                {
                    IngredientId = "ING002",
                    IngredientCode = "ING002",
                    IngredientName = "Sugar",
                    PackingType = "50kg Bag",
                    UnitOfMeasure = "kg",
                    CreatedDate = DateTime.Now,
                    LastModifiedDate = DateTime.Now,
                    LastModifiedBy = "system"
                },
                new Ingredient
                {
                    IngredientId = "ING003",
                    IngredientCode = "ING003",
                    IngredientName = "Salt",
                    PackingType = "25kg Bag",
                    UnitOfMeasure = "kg",
                    CreatedDate = DateTime.Now,
                    LastModifiedDate = DateTime.Now,
                    LastModifiedBy = "system"
                },
                new Ingredient
                {
                    IngredientId = "ING004",
                    IngredientCode = "ING004",
                    IngredientName = "Water",
                    PackingType = "1L Bottle",
                    UnitOfMeasure = "L",
                    CreatedDate = DateTime.Now,
                    LastModifiedDate = DateTime.Now,
                    LastModifiedBy = "system"
                },
                new Ingredient
                {
                    IngredientId = "ING005",
                    IngredientCode = "ING005",
                    IngredientName = "Cooking Oil",
                    PackingType = "5L Bottle",
                    UnitOfMeasure = "L",
                    CreatedDate = DateTime.Now,
                    LastModifiedDate = DateTime.Now,
                    LastModifiedBy = "system"
                }
            };

            SaveIngredients(sampleIngredients);
        }

        private void SaveIngredients(List<Ingredient> ingredients)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Ingredients");

            // Create header row
            worksheet.Cell(1, 1).Value = "IngredientId";
            worksheet.Cell(1, 2).Value = "IngredientCode";
            worksheet.Cell(1, 3).Value = "IngredientName";
            worksheet.Cell(1, 4).Value = "PackingType";
            worksheet.Cell(1, 5).Value = "UnitOfMeasure";
            worksheet.Cell(1, 6).Value = "CreatedDate";
            worksheet.Cell(1, 7).Value = "LastModifiedDate";
            worksheet.Cell(1, 8).Value = "LastModifiedBy";

            // Style header row
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Write data rows
            int rowIndex = 2;
            foreach (var ingredient in ingredients)
            {
                worksheet.Cell(rowIndex, 1).Value = ingredient.IngredientId;
                worksheet.Cell(rowIndex, 2).Value = ingredient.IngredientCode;
                worksheet.Cell(rowIndex, 3).Value = ingredient.IngredientName;
                worksheet.Cell(rowIndex, 4).Value = ingredient.PackingType;
                worksheet.Cell(rowIndex, 5).Value = ingredient.UnitOfMeasure;
                worksheet.Cell(rowIndex, 6).Value = ingredient.CreatedDate;
                worksheet.Cell(rowIndex, 7).Value = ingredient.LastModifiedDate;
                worksheet.Cell(rowIndex, 8).Value = ingredient.LastModifiedBy;

                rowIndex++;
            }

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(_ingredientsFilePath);
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
                    var ingredient = new Ingredient
                    {
                        IngredientId = worksheet.Cell(row, 1).GetString(),
                        IngredientCode = worksheet.Cell(row, 2).GetString(),
                        IngredientName = worksheet.Cell(row, 3).GetString(),
                        PackingType = worksheet.Cell(row, 4).GetString(),
                        UnitOfMeasure = worksheet.Cell(row, 5).GetString(),
                        CreatedDate = worksheet.Cell(row, 6).IsEmpty() ? DateTime.Now : worksheet.Cell(row, 6).GetDateTime(),
                        LastModifiedDate = worksheet.Cell(row, 7).IsEmpty() ? DateTime.Now : worksheet.Cell(row, 7).GetDateTime(),
                        LastModifiedBy = worksheet.Cell(row, 8).GetString()
                    };

                    ingredients.Add(ingredient);
                }

                _lastLoadedTime = DateTime.Now;
                return ingredients;
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