using ClosedXML.Excel;
using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    // Add this class for import results
    public class ImportResult
    {
        public bool Success { get; set; }
        public int RecipesImported { get; set; }
        public int IngredientsImported { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public List<string> Warnings { get; set; } = new();
    }

    public class RecipeService
    {
        private readonly string _recipesFilePath;
        private readonly string _recipeIngredientsFilePath;

        public RecipeService(IConfiguration configuration)
        {
            var baseDirectory = AppContext.BaseDirectory;

            _recipesFilePath = Path.Combine(baseDirectory,
                configuration["DataStorage:RecipesFilePath"] ?? "Data/Recipes.xlsx");

            _recipeIngredientsFilePath = Path.Combine(baseDirectory,
                configuration["DataStorage:RecipeIngredientsFilePath"] ?? "Data/RecipeIngredients.xlsx");

            Console.WriteLine($"🔧 Recipe Service initialized");
            Console.WriteLine($"📄 Recipes File: {_recipesFilePath}");
            Console.WriteLine($"📄 Recipe Ingredients File: {_recipeIngredientsFilePath}");

            EnsureFilesExist();
        }

        private void EnsureFilesExist()
        {
            // Ensure directory exists
            var directory = Path.GetDirectoryName(_recipesFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Create Recipes file if not exists
            if (!File.Exists(_recipesFilePath))
            {
                CreateRecipesFile();
            }

            // Create Recipe Ingredients file if not exists
            if (!File.Exists(_recipeIngredientsFilePath))
            {
                CreateRecipeIngredientsFile();
            }
        }

        private void CreateRecipesFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Recipes");

            // Headers
            worksheet.Cell(1, 1).Value = "RecipeId";
            worksheet.Cell(1, 2).Value = "RecipeCode";
            worksheet.Cell(1, 3).Value = "RecipeName";
            worksheet.Cell(1, 4).Value = "Description";
            worksheet.Cell(1, 5).Value = "Status";
            worksheet.Cell(1, 6).Value = "CreatedDate";
            worksheet.Cell(1, 7).Value = "LastModifiedDate";
            worksheet.Cell(1, 8).Value = "CreatedBy";
            worksheet.Cell(1, 9).Value = "LastModifiedBy";

            // Style headers
            var headerRange = worksheet.Range(1, 1, 1, 9);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_recipesFilePath);
            Console.WriteLine($"✅ Created Recipes file: {_recipesFilePath}");
        }

        private void CreateRecipeIngredientsFile()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("RecipeIngredients");

            // Headers
            worksheet.Cell(1, 1).Value = "RecipeId";
            worksheet.Cell(1, 2).Value = "Sequence";
            worksheet.Cell(1, 3).Value = "IngredientId";
            worksheet.Cell(1, 4).Value = "IngredientCode";
            worksheet.Cell(1, 5).Value = "IngredientName";
            worksheet.Cell(1, 6).Value = "TargetWeight";
            worksheet.Cell(1, 7).Value = "TolerancePercentage";
            worksheet.Cell(1, 8).Value = "ScaleNumber";
            worksheet.Cell(1, 9).Value = "Unit";
            worksheet.Cell(1, 10).Value = "BowlSize";

            // Style headers
            var headerRange = worksheet.Range(1, 1, 1, 10);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            workbook.SaveAs(_recipeIngredientsFilePath);
            Console.WriteLine($"✅ Created Recipe Ingredients file: {_recipeIngredientsFilePath}");
        }

        public async Task<List<Recipe>> GetAllRecipesAsync()
        {
            return await Task.Run(() =>
            {
                var recipes = new List<Recipe>();

                try
                {
                    using var workbook = new XLWorkbook(_recipesFilePath);
                    var worksheet = workbook.Worksheet("Recipes");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        try
                        {
                            var recipe = new Recipe
                            {
                                RecipeId = row.Cell(1).GetString(),
                                RecipeCode = row.Cell(2).GetString(),
                                RecipeName = row.Cell(3).GetString(),
                                Description = row.Cell(4).GetString(),
                                Status = row.Cell(5).GetString(),
                                CreatedDate = row.Cell(6).GetDateTime(),
                                LastModifiedDate = row.Cell(7).GetDateTime(),
                                CreatedBy = row.Cell(8).GetString(),
                                LastModifiedBy = row.Cell(9).GetString()
                            };

                            recipes.Add(recipe);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error reading recipe row: {ex.Message}");
                        }
                    }

                    Console.WriteLine($"✅ Loaded {recipes.Count} recipes");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading recipes: {ex.Message}");
                }

                return recipes;
            });
        }

        public async Task<Recipe?> GetRecipeByIdAsync(string recipeId)
        {
            var recipes = await GetAllRecipesAsync();
            var recipe = recipes.FirstOrDefault(r => r.RecipeId == recipeId);

            if (recipe != null)
            {
                recipe.Ingredients = await GetRecipeIngredientsAsync(recipeId);
            }

            return recipe;
        }

        public async Task<List<RecipeIngredient>> GetRecipeIngredientsAsync(string recipeId)
        {
            return await Task.Run(() =>
            {
                var ingredients = new List<RecipeIngredient>();

                try
                {
                    using var workbook = new XLWorkbook(_recipeIngredientsFilePath);
                    var worksheet = workbook.Worksheet("RecipeIngredients");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        try
                        {
                            var recId = row.Cell(1).GetString();
                            if (recId == recipeId)
                            {
                                var ingredient = new RecipeIngredient
                                {
                                    RecipeId = recId,
                                    Sequence = row.Cell(2).GetValue<int>(),
                                    IngredientId = row.Cell(3).GetString(),
                                    IngredientCode = row.Cell(4).GetString(),
                                    IngredientName = row.Cell(5).GetString(),
                                    TargetWeight = row.Cell(6).GetValue<decimal>(),
                                    TolerancePercentage = row.Cell(7).GetValue<decimal>(),
                                    ScaleNumber = row.Cell(8).GetValue<int>(),
                                    Unit = row.Cell(9).GetString(),
                                    BowlSize = row.Cell(10).IsEmpty() ? "Medium" : row.Cell(10).GetString()
                                };

                                ingredients.Add(ingredient);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error reading ingredient row: {ex.Message}");
                        }
                    }

                    Console.WriteLine($"✅ Loaded {ingredients.Count} ingredients for recipe {recipeId}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error loading recipe ingredients: {ex.Message}");
                }

                return ingredients.OrderBy(i => i.Sequence).ToList();
            });
        }

        public async Task SaveRecipeAsync(Recipe recipe)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_recipesFilePath);
                    var worksheet = workbook.Worksheet("Recipes");

                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = recipe.RecipeId;
                    worksheet.Cell(newRow, 2).Value = recipe.RecipeCode;
                    worksheet.Cell(newRow, 3).Value = recipe.RecipeName;
                    worksheet.Cell(newRow, 4).Value = recipe.Description;
                    worksheet.Cell(newRow, 5).Value = recipe.Status;
                    worksheet.Cell(newRow, 6).Value = recipe.CreatedDate;
                    worksheet.Cell(newRow, 7).Value = recipe.LastModifiedDate;
                    worksheet.Cell(newRow, 8).Value = recipe.CreatedBy;
                    worksheet.Cell(newRow, 9).Value = recipe.LastModifiedBy;

                    workbook.Save();
                    Console.WriteLine($"✅ Recipe {recipe.RecipeCode} saved");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error saving recipe: {ex.Message}");
                    throw;
                }
            });
        }

        public async Task SaveRecipeIngredientAsync(RecipeIngredient ingredient)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_recipeIngredientsFilePath);
                    var worksheet = workbook.Worksheet("RecipeIngredients");

                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = ingredient.RecipeId;
                    worksheet.Cell(newRow, 2).Value = ingredient.Sequence;
                    worksheet.Cell(newRow, 3).Value = ingredient.IngredientId;
                    worksheet.Cell(newRow, 4).Value = ingredient.IngredientCode;
                    worksheet.Cell(newRow, 5).Value = ingredient.IngredientName;
                    worksheet.Cell(newRow, 6).Value = ingredient.TargetWeight;
                    worksheet.Cell(newRow, 7).Value = ingredient.TolerancePercentage;
                    worksheet.Cell(newRow, 8).Value = ingredient.ScaleNumber;
                    worksheet.Cell(newRow, 9).Value = ingredient.Unit;
                    worksheet.Cell(newRow, 10).Value = ingredient.BowlSize;

                    workbook.Save();
                    Console.WriteLine($"✅ Ingredient added to recipe {ingredient.RecipeId}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error saving recipe ingredient: {ex.Message}");
                    throw;
                }
            });
        }

        public async Task UpdateRecipeAsync(Recipe recipe)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_recipesFilePath);
                    var worksheet = workbook.Worksheet("Recipes");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString() == recipe.RecipeId)
                        {
                            row.Cell(2).Value = recipe.RecipeCode;
                            row.Cell(3).Value = recipe.RecipeName;
                            row.Cell(4).Value = recipe.Description;
                            row.Cell(5).Value = recipe.Status;
                            row.Cell(7).Value = recipe.LastModifiedDate;
                            row.Cell(9).Value = recipe.LastModifiedBy;

                            workbook.Save();
                            Console.WriteLine($"✅ Recipe {recipe.RecipeCode} updated");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error updating recipe: {ex.Message}");
                    throw;
                }
            });
        }

        public async Task DeleteRecipeAsync(string recipeId)
        {
            await Task.Run(() =>
            {
                try
                {
                    // Delete recipe
                    using (var workbook = new XLWorkbook(_recipesFilePath))
                    {
                        var worksheet = workbook.Worksheet("Recipes");
                        var rows = worksheet.RowsUsed().Skip(1);

                        foreach (var row in rows)
                        {
                            if (row.Cell(1).GetString() == recipeId)
                            {
                                row.Delete();
                                workbook.Save();
                                break;
                            }
                        }
                    }

                    // Delete all ingredients for this recipe
                    using (var workbook = new XLWorkbook(_recipeIngredientsFilePath))
                    {
                        var worksheet = workbook.Worksheet("RecipeIngredients");
                        var rows = worksheet.RowsUsed().Skip(1).ToList();
                        rows.Reverse();

                        foreach (var row in rows)
                        {
                            if (row.Cell(1).GetString() == recipeId)
                            {
                                row.Delete();
                            }
                        }

                        workbook.Save();
                    }

                    Console.WriteLine($"✅ Recipe {recipeId} deleted");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error deleting recipe: {ex.Message}");
                    throw;
                }
            });
        }

        public async Task DeleteRecipeIngredientAsync(string recipeId, int sequence)
        {
            await Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(_recipeIngredientsFilePath);
                    var worksheet = workbook.Worksheet("RecipeIngredients");
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString() == recipeId &&
                            row.Cell(2).GetValue<int>() == sequence)
                        {
                            row.Delete();
                            workbook.Save();
                            Console.WriteLine($"✅ Ingredient removed from recipe {recipeId}");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error deleting recipe ingredient: {ex.Message}");
                    throw;
                }
            });
        }

        // ==================== NEW: IMPORT TEMPLATE CREATION ====================
        public async Task<string> CreateImportTemplateAsync(List<Ingredient> availableIngredients)
        {
            return await Task.Run(() =>
            {
                var templatePath = Path.Combine(Path.GetDirectoryName(_recipesFilePath) ?? "", "RecipeImportTemplate.xlsx");

                using var workbook = new XLWorkbook();

                // ==================== CREATE INSTRUCTIONS SHEET ====================
                var instructionsSheet = workbook.Worksheets.Add("Instructions");
                instructionsSheet.Column(1).Width = 80;

                instructionsSheet.Cell(1, 1).Value = "FirstWeigh Recipe Import Template - Instructions";
                instructionsSheet.Cell(1, 1).Style.Font.Bold = true;
                instructionsSheet.Cell(1, 1).Style.Font.FontSize = 16;
                instructionsSheet.Cell(1, 1).Style.Font.FontColor = XLColor.White;
                instructionsSheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(74, 158, 255);
                instructionsSheet.Row(1).Height = 30;

                var row = 3;
                instructionsSheet.Cell(row++, 1).Value = "HOW TO USE THIS TEMPLATE:";
                instructionsSheet.Cell(row++, 1).Value = "1. Check the 'Available Ingredients' sheet to see what ingredients you can use";
                instructionsSheet.Cell(row++, 1).Value = "2. Fill in the 'Recipes' sheet with your recipe information";
                instructionsSheet.Cell(row++, 1).Value = "3. Fill in the 'RecipeIngredients' sheet with ingredient details";
                instructionsSheet.Cell(row++, 1).Value = "4. Use the dropdown lists for Status, BowlSize, and ScaleNumber";
                instructionsSheet.Cell(row++, 1).Value = "5. Make sure RecipeId matches between both sheets";
                instructionsSheet.Cell(row++, 1).Value = "6. Ingredient codes MUST match exactly what's in 'Available Ingredients' sheet";
                instructionsSheet.Cell(row++, 1).Value = "7. Save and import the file into FirstWeigh";

                row++;
                instructionsSheet.Cell(row++, 1).Value = "IMPORTANT NOTES:";
                instructionsSheet.Cell(row++, 1).Value = "• RecipeId and RecipeCode should be unique (e.g., RCP001, RCP002)";
                instructionsSheet.Cell(row++, 1).Value = "• Sequence numbers should start at 1 and increment by 1";
                instructionsSheet.Cell(row++, 1).Value = "• TargetWeight should be in kilograms (kg)";
                instructionsSheet.Cell(row++, 1).Value = "• TolerancePercentage is a percentage value (e.g., 2.0 = 2%)";
                instructionsSheet.Cell(row++, 1).Value = "• Status options: Active, Draft, Archived";
                instructionsSheet.Cell(row++, 1).Value = "• BowlSize options: Small, Medium, Large";
                instructionsSheet.Cell(row++, 1).Value = "• ScaleNumber options: 1 or 2";
                instructionsSheet.Cell(row++, 1).Value = "• CRITICAL: Only use ingredient codes from the 'Available Ingredients' sheet!";

                row++;
                instructionsSheet.Cell(row++, 1).Value = "STEP-BY-STEP PROCESS:";
                instructionsSheet.Cell(row++, 1).Value = "1. Go to the 'Available Ingredients' sheet";
                instructionsSheet.Cell(row++, 1).Value = "2. Find the ingredient you want to use";
                instructionsSheet.Cell(row++, 1).Value = "3. Copy the IngredientCode (e.g., ING001)";
                instructionsSheet.Cell(row++, 1).Value = "4. Paste it into both IngredientId AND IngredientCode columns in RecipeIngredients sheet";
                instructionsSheet.Cell(row++, 1).Value = "5. The system will automatically use the correct ingredient name and unit during import";

                // ==================== CREATE AVAILABLE INGREDIENTS SHEET ====================
                var availIngSheet = workbook.Worksheets.Add("Available Ingredients");

                availIngSheet.Column(1).Width = 14;
                availIngSheet.Column(2).Width = 14;
                availIngSheet.Column(3).Width = 30;
                availIngSheet.Column(4).Width = 12;
                availIngSheet.Column(5).Width = 35;

                // Headers
                availIngSheet.Cell(1, 1).Value = "IngredientId";
                availIngSheet.Cell(1, 2).Value = "IngredientCode";
                availIngSheet.Cell(1, 3).Value = "IngredientName";
                availIngSheet.Cell(1, 4).Value = "Unit";
                availIngSheet.Cell(1, 5).Value = "PackingType";

                for (int i = 1; i <= 5; i++)
                {
                    var cell = availIngSheet.Cell(1, i);
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Fill.BackgroundColor = XLColor.FromArgb(72, 187, 120); // Green
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                availIngSheet.Row(1).Height = 25;

                // Add all available ingredients
                int ingredientRow = 2;
                foreach (var ingredient in availableIngredients.OrderBy(i => i.IngredientCode))
                {
                    availIngSheet.Cell(ingredientRow, 1).Value = ingredient.IngredientId;
                    availIngSheet.Cell(ingredientRow, 2).Value = ingredient.IngredientCode;
                    availIngSheet.Cell(ingredientRow, 3).Value = ingredient.IngredientName;
                    availIngSheet.Cell(ingredientRow, 4).Value = ingredient.UnitOfMeasure;
                    availIngSheet.Cell(ingredientRow, 5).Value = ingredient.PackingType;

                    // Add borders
                    availIngSheet.Range(ingredientRow, 1, ingredientRow, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    ingredientRow++;
                }

                // Freeze header row
                availIngSheet.SheetView.FreezeRows(1);

                // Add note if no ingredients
                if (availableIngredients.Count == 0)
                {
                    availIngSheet.Cell(2, 1).Value = "No ingredients found in system!";
                    availIngSheet.Cell(2, 1).Style.Font.Bold = true;
                    availIngSheet.Cell(2, 1).Style.Font.FontColor = XLColor.Red;
                }

                // ==================== CREATE REFERENCE DATA SHEET (HIDDEN) ====================
                var refSheet = workbook.Worksheets.Add("RefData");

                // Status options
                refSheet.Cell(1, 1).Value = "Status";
                refSheet.Cell(2, 1).Value = "Active";
                refSheet.Cell(3, 1).Value = "Draft";
                refSheet.Cell(4, 1).Value = "Archived";

                // Bowl size options
                refSheet.Cell(1, 2).Value = "BowlSize";
                refSheet.Cell(2, 2).Value = "Small";
                refSheet.Cell(3, 2).Value = "Medium";
                refSheet.Cell(4, 2).Value = "Large";

                // Scale number options
                refSheet.Cell(1, 3).Value = "ScaleNumber";
                refSheet.Cell(2, 3).Value = "1";
                refSheet.Cell(3, 3).Value = "2";

                refSheet.Hide();

                // ==================== CREATE RECIPES SHEET ====================
                var recipesSheet = workbook.Worksheets.Add("Recipes");

                // Set column widths
                recipesSheet.Column(1).Width = 12;
                recipesSheet.Column(2).Width = 12;
                recipesSheet.Column(3).Width = 25;
                recipesSheet.Column(4).Width = 35;
                recipesSheet.Column(5).Width = 12;
                recipesSheet.Column(6).Width = 15;
                recipesSheet.Column(7).Width = 15;
                recipesSheet.Column(8).Width = 15;
                recipesSheet.Column(9).Width = 15;

                // Headers
                var recipeHeaders = new[] { "RecipeId", "RecipeCode", "RecipeName", "Description", "Status",
                                     "CreatedDate", "LastModifiedDate", "CreatedBy", "LastModifiedBy" };
                for (int i = 0; i < recipeHeaders.Length; i++)
                {
                    var cell = recipesSheet.Cell(1, i + 1);
                    cell.Value = recipeHeaders[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Fill.BackgroundColor = XLColor.FromArgb(74, 158, 255);
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                recipesSheet.Row(1).Height = 25;

                // Example row
                recipesSheet.Cell(2, 1).Value = "RCP999";
                recipesSheet.Cell(2, 2).Value = "RCP999";
                recipesSheet.Cell(2, 3).Value = "Example Recipe";
                recipesSheet.Cell(2, 4).Value = "This is an example recipe - replace with your data";
                recipesSheet.Cell(2, 5).Value = "Active";
                recipesSheet.Cell(2, 6).Value = DateTime.Now;
                recipesSheet.Cell(2, 6).Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                recipesSheet.Cell(2, 7).Value = DateTime.Now;
                recipesSheet.Cell(2, 7).Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                recipesSheet.Cell(2, 8).Value = "system";
                recipesSheet.Cell(2, 9).Value = "system";

                // Add dropdown for Status
                var statusValidation = recipesSheet.Range("E2:E100").CreateDataValidation();
                statusValidation.List("RefData!$A$2:$A$4", true);
                statusValidation.IgnoreBlanks = true;
                statusValidation.InCellDropdown = true;

                recipesSheet.SheetView.FreezeRows(1);

                // ==================== CREATE RECIPE INGREDIENTS SHEET ====================
                var ingredientsSheet = workbook.Worksheets.Add("RecipeIngredients");

                ingredientsSheet.Column(1).Width = 12;
                ingredientsSheet.Column(2).Width = 10;
                ingredientsSheet.Column(3).Width = 14;
                ingredientsSheet.Column(4).Width = 14;
                ingredientsSheet.Column(5).Width = 25;
                ingredientsSheet.Column(6).Width = 14;
                ingredientsSheet.Column(7).Width = 18;
                ingredientsSheet.Column(8).Width = 13;
                ingredientsSheet.Column(9).Width = 10;
                ingredientsSheet.Column(10).Width = 12;

                // Headers
                var ingredientHeaders = new[] { "RecipeId", "Sequence", "IngredientId", "IngredientCode",
                                         "IngredientName", "TargetWeight", "TolerancePercentage",
                                         "ScaleNumber", "Unit", "BowlSize" };
                for (int i = 0; i < ingredientHeaders.Length; i++)
                {
                    var cell = ingredientsSheet.Cell(1, i + 1);
                    cell.Value = ingredientHeaders[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Fill.BackgroundColor = XLColor.FromArgb(74, 158, 255);
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                ingredientsSheet.Row(1).Height = 25;

                // Example rows - use real ingredients if available
                if (availableIngredients.Count > 0)
                {
                    var exampleIngredient = availableIngredients.First();

                    ingredientsSheet.Cell(2, 1).Value = "RCP999";
                    ingredientsSheet.Cell(2, 2).Value = 1;
                    ingredientsSheet.Cell(2, 3).Value = exampleIngredient.IngredientId;
                    ingredientsSheet.Cell(2, 4).Value = exampleIngredient.IngredientCode;
                    ingredientsSheet.Cell(2, 5).Value = exampleIngredient.IngredientName;
                    ingredientsSheet.Cell(2, 6).Value = 10.5;
                    ingredientsSheet.Cell(2, 6).Style.NumberFormat.Format = "0.000";
                    ingredientsSheet.Cell(2, 7).Value = 2.0;
                    ingredientsSheet.Cell(2, 7).Style.NumberFormat.Format = "0.00";
                    ingredientsSheet.Cell(2, 8).Value = 1;
                    ingredientsSheet.Cell(2, 9).Value = exampleIngredient.UnitOfMeasure;
                    ingredientsSheet.Cell(2, 10).Value = "Medium";
                }
                else
                {
                    // Placeholder if no ingredients
                    for (int i = 2; i <= 3; i++)
                    {
                        ingredientsSheet.Cell(i, 1).Value = "RCP999";
                        ingredientsSheet.Cell(i, 2).Value = i - 1;
                        ingredientsSheet.Cell(i, 3).Value = $"ING00{i - 1}";
                        ingredientsSheet.Cell(i, 4).Value = $"ING00{i - 1}";
                        ingredientsSheet.Cell(i, 5).Value = $"Example Ingredient {i - 1}";
                        ingredientsSheet.Cell(i, 6).Value = 10.5;
                        ingredientsSheet.Cell(i, 6).Style.NumberFormat.Format = "0.000";
                        ingredientsSheet.Cell(i, 7).Value = 2.0;
                        ingredientsSheet.Cell(i, 7).Style.NumberFormat.Format = "0.00";
                        ingredientsSheet.Cell(i, 8).Value = 1;
                        ingredientsSheet.Cell(i, 9).Value = "kg";
                        ingredientsSheet.Cell(i, 10).Value = "Medium";
                    }
                }

                // Add dropdown for BowlSize
                var bowlValidation = ingredientsSheet.Range("J2:J100").CreateDataValidation();
                bowlValidation.List("RefData!$B$2:$B$4", true);
                bowlValidation.IgnoreBlanks = true;
                bowlValidation.InCellDropdown = true;

                // Add dropdown for ScaleNumber
                var scaleValidation = ingredientsSheet.Range("H2:H100").CreateDataValidation();
                scaleValidation.List("RefData!$C$2:$C$3", true);
                scaleValidation.IgnoreBlanks = true;
                scaleValidation.InCellDropdown = true;

                // Set number formats
                ingredientsSheet.Range("F2:F100").Style.NumberFormat.Format = "0.000";
                ingredientsSheet.Range("G2:G100").Style.NumberFormat.Format = "0.00";

                ingredientsSheet.SheetView.FreezeRows(1);

                // Set active sheet to Instructions
                instructionsSheet.SetTabActive();

                workbook.SaveAs(templatePath);
                Console.WriteLine($"✅ Enhanced template created with {availableIngredients.Count} available ingredients");
                return templatePath;
            });
        }

        // ==================== NEW: IMPORT FROM EXCEL ====================
        public async Task<ImportResult> ImportRecipesFromExcelAsync(Stream excelStream, List<Ingredient> availableIngredients)
        {
            var result = new ImportResult();

            try
            {
                using var workbook = new XLWorkbook(excelStream);

                // Check if required worksheets exist
                if (!workbook.Worksheets.Contains("Recipes"))
                {
                    result.ErrorMessage = "Excel file must contain a 'Recipes' worksheet";
                    return result;
                }

                if (!workbook.Worksheets.Contains("RecipeIngredients"))
                {
                    result.ErrorMessage = "Excel file must contain a 'RecipeIngredients' worksheet";
                    return result;
                }

                var recipesSheet = workbook.Worksheet("Recipes");
                var ingredientsSheet = workbook.Worksheet("RecipeIngredients");

                // Validate headers
                if (!ValidateRecipeHeaders(recipesSheet) || !ValidateIngredientHeaders(ingredientsSheet))
                {
                    result.ErrorMessage = "Excel file format is invalid. Please use the template format.";
                    return result;
                }

                // Import recipes
                var existingRecipes = await GetAllRecipesAsync();
                var recipeRows = recipesSheet.RowsUsed().Skip(1);

                foreach (var row in recipeRows)
                {
                    try
                    {
                        var recipeCode = row.Cell(2).GetString();

                        // Skip if recipe already exists
                        if (existingRecipes.Any(r => r.RecipeCode == recipeCode))
                        {
                            result.Warnings.Add($"Recipe {recipeCode} already exists - skipped");
                            continue;
                        }

                        var recipe = new Recipe
                        {
                            RecipeId = row.Cell(1).GetString(),
                            RecipeCode = recipeCode,
                            RecipeName = row.Cell(3).GetString(),
                            Description = row.Cell(4).GetString(),
                            Status = row.Cell(5).GetString(),
                            CreatedDate = row.Cell(6).GetDateTime(),
                            LastModifiedDate = row.Cell(7).GetDateTime(),
                            CreatedBy = row.Cell(8).GetString(),
                            LastModifiedBy = row.Cell(9).GetString()
                        };

                        await SaveRecipeAsync(recipe);
                        result.RecipesImported++;
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"Error importing recipe at row {row.RowNumber()}: {ex.Message}");
                    }
                }

                // Import recipe ingredients
                var ingredientRows = ingredientsSheet.RowsUsed().Skip(1);

                foreach (var row in ingredientRows)
                {
                    try
                    {
                        var recipeId = row.Cell(1).GetString();
                        var ingredientCode = row.Cell(4).GetString();

                        // Find matching ingredient in system
                        var ingredient = availableIngredients.FirstOrDefault(i =>
                            i.IngredientCode.Equals(ingredientCode, StringComparison.OrdinalIgnoreCase));

                        if (ingredient == null)
                        {
                            result.Warnings.Add($"Ingredient code '{ingredientCode}' not found in system - skipped at row {row.RowNumber()}");
                            continue;
                        }

                        var recipeIngredient = new RecipeIngredient
                        {
                            RecipeId = recipeId,
                            Sequence = row.Cell(2).GetValue<int>(),
                            IngredientId = ingredient.IngredientId,
                            IngredientCode = ingredient.IngredientCode,
                            IngredientName = ingredient.IngredientName,
                            TargetWeight = row.Cell(6).GetValue<decimal>(),
                            TolerancePercentage = row.Cell(7).GetValue<decimal>(),
                            ScaleNumber = row.Cell(8).GetValue<int>(),
                            Unit = ingredient.UnitOfMeasure,
                            BowlSize = row.Cell(10).IsEmpty() ? "Medium" : row.Cell(10).GetString()
                        };

                        await SaveRecipeIngredientAsync(recipeIngredient);
                        result.IngredientsImported++;
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"Error importing ingredient at row {row.RowNumber()}: {ex.Message}");
                    }
                }

                result.Success = result.RecipesImported > 0;
                if (!result.Success && result.RecipesImported == 0)
                {
                    result.ErrorMessage = "No recipes were imported";
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        private bool ValidateRecipeHeaders(IXLWorksheet worksheet)
        {
            try
            {
                return worksheet.Cell(1, 1).GetString() == "RecipeId" &&
                       worksheet.Cell(1, 2).GetString() == "RecipeCode" &&
                       worksheet.Cell(1, 3).GetString() == "RecipeName";
            }
            catch
            {
                return false;
            }
        }

        private bool ValidateIngredientHeaders(IXLWorksheet worksheet)
        {
            try
            {
                return worksheet.Cell(1, 1).GetString() == "RecipeId" &&
                       worksheet.Cell(1, 2).GetString() == "Sequence" &&
                       worksheet.Cell(1, 3).GetString() == "IngredientId";
            }
            catch
            {
                return false;
            }
        }

        public void OpenRecipesFile()
        {
            if (File.Exists(_recipesFilePath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = _recipesFilePath,
                    UseShellExecute = true
                });
            }
        }

        public void OpenRecipeIngredientsFile()
        {
            if (File.Exists(_recipeIngredientsFilePath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = _recipeIngredientsFilePath,
                    UseShellExecute = true
                });
            }
        }
    }
}