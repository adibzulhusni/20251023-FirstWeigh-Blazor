using ClosedXML.Excel;
using FirstWeigh.Models;

namespace FirstWeigh.Services
{
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
            worksheet.Cell(1, 10).Value = "BowlSize"; // ✅ NEW COLUMN

            // Style headers
            var headerRange = worksheet.Range(1, 1, 1, 10); // ✅ Changed from 9 to 10
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
                                    BowlSize = row.Cell(10).IsEmpty() ? "Medium" : row.Cell(10).GetString() // ✅ NEW
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
                    worksheet.Cell(newRow, 10).Value = ingredient.BowlSize; // ✅ NEW

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