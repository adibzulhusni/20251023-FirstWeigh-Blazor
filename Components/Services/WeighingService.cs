using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    public class WeighingService : IWeighingService
    {
        private readonly IBatchService _batchService;
        private readonly RecipeService _recipeService;
        private WeighingSession? _activeSession;

        public WeighingService(IBatchService batchService, RecipeService recipeService)
        {
            _batchService = batchService;
            _recipeService = recipeService;
        }

        public async Task<WeighingSession?> StartWeighingSessionAsync(string batchId)
        {
            var batch = await _batchService.GetBatchByIdAsync(batchId);
            if (batch == null || batch.Status != "InProgress")
                return null;

            var recipe = await _recipeService.GetRecipeByIdAsync(batch.RecipeId);
            if (recipe == null)
                return null;

            var ingredients = await _recipeService.GetRecipeIngredientsAsync(batch.RecipeId);
            if (ingredients == null || !ingredients.Any())
                return null;

            _activeSession = new WeighingSession
            {
                BatchId = batchId,
                RecipeId = batch.RecipeId,
                CurrentRepetition = batch.CurrentRepetition + 1,
                TotalRepetitions = batch.TotalRepetitions,
                CurrentIngredientIndex = 0,
                Ingredients = ingredients.OrderBy(i => i.Sequence).ToList(),
                OperatorName = batch.StartedBy ?? "Operator",
                SessionStarted = DateTime.Now
            };

            return _activeSession;
        }

        public bool SelectBowls(string batchId, string ingredientBowlCode, decimal ingredientBowlWeight,
                                string mixingBowlCode, decimal mixingBowlWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            _activeSession.SelectedIngredientBowlCode = ingredientBowlCode;
            _activeSession.SelectedIngredientBowlWeight = ingredientBowlWeight;
            _activeSession.SelectedMixingBowlCode = mixingBowlCode;
            _activeSession.SelectedMixingBowlWeight = mixingBowlWeight;
            _activeSession.MixingBowlWeightBefore = mixingBowlWeight;

            Console.WriteLine($"✅ Bowls selected - Ingredient: {ingredientBowlCode} ({ingredientBowlWeight:F3} kg), Mixing: {mixingBowlCode} ({mixingBowlWeight:F3} kg)");
            return true;
        }

        public (bool isValid, string message) VerifyBowlWeight(decimal actualWeight, decimal recordedWeight, string bowlCode, decimal tolerance = 0.01m)
        {
            var difference = Math.Abs(actualWeight - recordedWeight);

            if (difference <= tolerance)
            {
                return (true, $"✓ Bowl {bowlCode} verified: {actualWeight:F3} kg");
            }
            else
            {
                return (false, $"⚠ Bowl {bowlCode} weight mismatch!\nExpected: {recordedWeight:F3} kg\nActual: {actualWeight:F3} kg\nDifference: {difference:F3} kg");
            }
        }

        public decimal GetNetIngredientWeight(decimal currentScale1Weight)
        {
            if (_activeSession == null)
                return 0;

            return currentScale1Weight - _activeSession.SelectedIngredientBowlWeight;
        }

        public bool RecordBowlWeights(string batchId, decimal ingredientBowlWeight, decimal mixingBowlWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            _activeSession.IngredientBowlWeight = ingredientBowlWeight;
            _activeSession.MixingBowlWeightBefore = mixingBowlWeight;
            _activeSession.CurrentStage = WeighingStage.WeighIngredient;

            Console.WriteLine($"✅ Bowls recorded - Ingredient bowl: {ingredientBowlWeight:F3} kg, Mixing bowl: {mixingBowlWeight:F3} kg");
            return true;
        }

        public (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatusByNet(decimal netWeight, RecipeIngredient ingredient)
        {
            if (netWeight < ingredient.MinWeight)
            {
                return ("red", "⬇", "Keep adding material - Under target", false);
            }
            else if (netWeight >= ingredient.MinWeight && netWeight <= ingredient.MaxWeight)
            {
                return ("green", "✓", "GOOD - Target reached!", true);
            }
            else
            {
                return ("red", "⚠", "OVER TARGET - Stop adding!", false);
            }
        }

        // ✅ FIXED: Don't change CurrentStage
        public async Task<bool> ReadyToTransferAsync(string batchId, decimal netWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            _activeSession.NetIngredientWeight = netWeight;
            // ❌ REMOVED: _activeSession.CurrentStage = WeighingStage.Transfer;

            Console.WriteLine($"✅ Ready to transfer - Net weight: {netWeight:F3} kg");
            return true;
        }

        public async Task<bool> ConfirmTransferAsync(string batchId, decimal currentScale2Weight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            // Verify transfer
            var expectedScale2 = _activeSession.MixingBowlWeightBefore + _activeSession.NetIngredientWeight;
            var actualScale2 = currentScale2Weight;
            var difference = Math.Abs(expectedScale2 - actualScale2);

            Console.WriteLine($"📊 Transfer verification:");
            Console.WriteLine($"   Expected Scale 2: {expectedScale2:F3} kg");
            Console.WriteLine($"   Actual Scale 2: {actualScale2:F3} kg");
            Console.WriteLine($"   Difference: {difference:F3} kg");

            // Move to next ingredient
            _activeSession.CurrentIngredientIndex++;

            // Check if all ingredients for this repetition are complete
            if (_activeSession.CurrentIngredientIndex >= _activeSession.Ingredients.Count)
            {
                // All ingredients complete for this repetition
                Console.WriteLine($"✅ Repetition {_activeSession.CurrentRepetition} complete!");

                // ✅ CHECK COMPLETION BEFORE INCREMENTING
                if (_activeSession.CurrentRepetition >= _activeSession.TotalRepetitions)
                {
                    // ✅ ALL REPETITIONS COMPLETE - Batch done!
                    Console.WriteLine($"🎉 All {_activeSession.TotalRepetitions} repetitions complete! Batch finished!");

                    // Update to mark last repetition complete
                    await _batchService.UpdateRepetitionProgressAsync(
                        batchId,
                        _activeSession.CurrentRepetition
                    );

                    await _batchService.CompleteBatchAsync(batchId);
                    _activeSession = null;
                    return true;
                }

                // ✅ MORE REPETITIONS TO DO - Move to next repetition
                _activeSession.CurrentRepetition++;
                Console.WriteLine($"🔄 Starting repetition {_activeSession.CurrentRepetition} of {_activeSession.TotalRepetitions}");

                // Update batch progress for completed repetition
                await _batchService.UpdateRepetitionProgressAsync(
                    batchId,
                    _activeSession.CurrentRepetition - 1
                );

                _activeSession.CurrentIngredientIndex = 0;
                _activeSession.CurrentStage = WeighingStage.PlaceBowls;

                // ✅ RESET BOTH BOWLS for next repetition
                _activeSession.SelectedIngredientBowlCode = null;
                _activeSession.SelectedIngredientBowlWeight = 0;
                _activeSession.SelectedMixingBowlCode = null;
                _activeSession.SelectedMixingBowlWeight = 0;
                _activeSession.IngredientBowlWeight = 0;
                _activeSession.MixingBowlWeightBefore = 0;
                _activeSession.NetIngredientWeight = 0;

                return true;
            }

            // ✅ NOT LAST INGREDIENT - Just reset ingredient bowl, mixing bowl stays
            Console.WriteLine($"➡️ Moving to next ingredient: {_activeSession.CurrentIngredient?.IngredientCode}");
            _activeSession.CurrentStage = WeighingStage.PlaceBowls;

            // Only reset INGREDIENT bowl
            _activeSession.SelectedIngredientBowlCode = null;
            _activeSession.SelectedIngredientBowlWeight = 0;
            _activeSession.IngredientBowlWeight = 0;
            _activeSession.NetIngredientWeight = 0;

            return true;
        }

        public async Task<bool> CompleteIngredientAsync(string batchId, decimal actualWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            return true;
        }

        public Task<bool> PauseSessionAsync(string batchId)
        {
            _activeSession = null;
            return Task.FromResult(true);
        }

        public async Task<bool> AbortSessionAsync(string batchId, string reason, string abortedBy)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            await _batchService.AbortBatchAsync(batchId, abortedBy, reason);
            _activeSession = null;

            return true;
        }

        public WeighingSession? GetActiveSession()
        {
            return _activeSession;
        }

        public void ClearActiveSession()
        {
            _activeSession = null;
        }

        public (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatus(decimal currentWeight, RecipeIngredient ingredient)
        {
            var targetWeight = ingredient.TargetWeight;
            var tolerance = ingredient.TolerancePercentage;

            var minWeight = targetWeight - tolerance;
            var maxWeight = targetWeight + tolerance;

            var yellowMin = targetWeight - (tolerance * 0.5m);
            var yellowMax = targetWeight + (tolerance * 0.5m);

            if (currentWeight > maxWeight)
            {
                return ("red", "❌", "OVER TARGET - Stop adding immediately!", false);
            }

            if (currentWeight >= yellowMin && currentWeight <= yellowMax)
            {
                return ("green", "✓", "GOOD - Target reached!", true);
            }

            if (currentWeight > minWeight && currentWeight < yellowMin)
            {
                return ("yellow", "⚠️", "Slow down - approaching target", false);
            }

            if (currentWeight < minWeight)
            {
                var percentage = (currentWeight / targetWeight) * 100;
                return ("red", "⬆️", $"Keep adding material ({percentage:F0}%)", false);
            }

            if (currentWeight > yellowMax && currentWeight <= maxWeight)
            {
                return ("yellow", "⚠️", "Caution - near maximum limit", true);
            }

            return ("yellow", "⚠️", "Continue adding", false);
        }
    }
}