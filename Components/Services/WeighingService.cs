using FirstWeigh.Models;
using System.Text;

namespace FirstWeigh.Services
{
    public class WeighingService : IWeighingService
    {
        private readonly IBatchService _batchService;
        private readonly RecipeService _recipeService;
        private readonly ReportService _reportService;
        private WeighingSession? _activeSession;

        // Configuration constants
        private const decimal BASE_TRANSFER_TOLERANCE = 0.050m;  // 50g base tolerance
        private const decimal PER_INGREDIENT_TOLERANCE = 0.015m; // 15g per ingredient cumulative
        private const decimal BOWL_VERIFICATION_TOLERANCE = 0.050m; // 50g for bowl weight verification
        private const decimal SCALE_STABILITY_TOLERANCE = 0.005m; // 5g for stability check

        public WeighingService(
            IBatchService batchService,
            RecipeService recipeService,
            ReportService reportService)
        {
            _batchService = batchService;
            _recipeService = recipeService;
            _reportService = reportService;
        }
        public async Task<bool> UpdateSessionOperatorAsync(string batchId, string operatorName)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            // Update session
            _activeSession.OperatorName = operatorName;

            // ✅ Update WeighingRecord in database
            if (!string.IsNullOrEmpty(_activeSession.WeighingRecordId))
            {
                var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
                if (record != null)
                {
                    record.OperatorName = operatorName;
                    await _reportService.UpdateWeighingRecordAsync(record);
                    Console.WriteLine($"✅ WeighingRecord operator updated to: {operatorName}");
                    return true;
                }
            }

            return false;
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

            // ✅ Create WeighingRecord at session start
            var record = new WeighingRecord
            {
                RecordId = await GenerateRecordIdAsync(),
                BatchId = batchId,
                RecipeId = batch.RecipeId,
                RecipeCode = recipe.RecipeCode,
                RecipeName = recipe.RecipeName,
                OperatorName = batch.StartedBy ?? "Operator",
                SessionStartTime = DateTime.Now,
                PlannedStartTime = batch.PlannedStartTime,
                PlannedEndTime = batch.PlannedEndTime,
                TotalRepetitions = batch.TotalRepetitions,
                CompletedRepetitions = 0,
                Status = WeighingRecordStatus.InProgress,
                CreatedDate = DateTime.Now,
                CreatedBy = batch.StartedBy ?? "System"
            };

            await _reportService.SaveWeighingRecordAsync(record);

            _activeSession = new WeighingSession
            {
                BatchId = batchId,
                RecipeId = batch.RecipeId,
                RecipeName = recipe.RecipeName,
                RecipeCode = recipe.RecipeCode,
                CurrentRepetition = batch.CurrentRepetition + 1,
                TotalRepetitions = batch.TotalRepetitions,
                CurrentIngredientIndex = 0,
                Ingredients = ingredients.OrderBy(i => i.Sequence).ToList(),
                OperatorName = batch.StartedBy ?? "Operator",
                SessionStarted = DateTime.Now,
                PlannedStartTime = batch.PlannedStartTime,
                PlannedEndTime = batch.PlannedEndTime,
                WeighingRecordId = record.RecordId
            };

            Console.WriteLine($"✅ Weighing session started - Record: {record.RecordId}");
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

        public (bool isValid, string message) VerifyBowlWeight(
            decimal actualWeight,
            decimal recordedWeight,
            string bowlCode,
            decimal tolerance = BOWL_VERIFICATION_TOLERANCE)
        {
            var difference = Math.Abs(actualWeight - recordedWeight);

            if (difference <= tolerance)
            {
                return (true, $"✓ Bowl {bowlCode} verified: {actualWeight:F3} kg");
            }
            else
            {
                return (false,
                    $"⚠ Bowl {bowlCode} weight mismatch!\n" +
                    $"Expected: {recordedWeight:F3} kg\n" +
                    $"Actual: {actualWeight:F3} kg\n" +
                    $"Difference: {difference:F3} kg (max: ±{tolerance:F3} kg)");
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
                var percentage = ingredient.TargetWeight > 0
                    ? (netWeight / ingredient.TargetWeight) * 100
                    : 0;
                return ("red", "⬇", $"Keep adding material ({percentage:F0}%)", false);
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

        public async Task<bool> ReadyToTransferAsync(string batchId, decimal netWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            _activeSession.NetIngredientWeight = netWeight;

            Console.WriteLine($"✅ Ready to transfer - Net weight: {netWeight:F3} kg");
            return true;
        }

        public async Task<(bool success, string message, decimal deviation)> ConfirmTransferAsync(string batchId,decimal currentScale2Weight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return (false, "No active session", 0);

            var ingredient = _activeSession.CurrentIngredient;
            if (ingredient == null)
                return (false, "No current ingredient", 0);

            // ✅ STEP 1: Calculate expected cumulative from ACTUAL transfers
            decimal expectedCumulative = _activeSession.TransferredIngredients
                .Where(t => t.RepetitionNumber == _activeSession.CurrentRepetition)
                .Sum(t => t.ActualNetWeight);

            // Add current ingredient's actual net weight
            expectedCumulative += _activeSession.NetIngredientWeight;

            // ✅ STEP 2: Get net Scale 2 reading (minus bowl tare)
            decimal scale2Before = _activeSession.MixingBowlWeightBefore;
            decimal actualScale2Net = currentScale2Weight - scale2Before;

            // ✅ STEP 3: Calculate deviation
            decimal deviation = actualScale2Net - expectedCumulative;
            decimal absDeviation = Math.Abs(deviation);

            // ✅ STEP 4: Dynamic tolerance based on number of ingredients
            int ingredientsTransferred = _activeSession.CurrentIngredientIndex + 1;
            decimal allowedTolerance = CalculateDynamicTransferTolerance(ingredientsTransferred);

            Console.WriteLine($"📊 Transfer Verification:");
            Console.WriteLine($"   Expected Cumulative (Actuals): {expectedCumulative:F3} kg");
            Console.WriteLine($"   Actual Scale 2 Net: {actualScale2Net:F3} kg");
            Console.WriteLine($"   Deviation: {deviation:F3} kg");
            Console.WriteLine($"   Allowed Tolerance: ±{allowedTolerance:F3} kg");

            // ✅ STEP 5: Verify within tolerance
            if (absDeviation > allowedTolerance)
            {
                var message = $"⚠️ Scale 2 weight mismatch!\n" +
                    $"Expected: {expectedCumulative:F3} kg\n" +
                    $"Actual: {actualScale2Net:F3} kg\n" +
                    $"Deviation: {deviation:F3} kg (max: ±{allowedTolerance:F3} kg)\n\n" +
                    $"Possible causes:\n" +
                    $"- Material not fully transferred\n" +
                    $"- Material spilled\n" +
                    $"- Scale drift or calibration issue";

                return (false, message, deviation);
            }

            // ✅ STEP 6: Calculate tolerance values
            decimal toleranceValue = (ingredient.TargetWeight * ingredient.TolerancePercentage) / 100;

            // ✅ STEP 7: Record the successful transfer
            var transferRecord = new TransferredIngredient
            {
                RepetitionNumber = _activeSession.CurrentRepetition,
                IngredientSequence = _activeSession.CurrentIngredientIndex + 1,
                IngredientId = ingredient.IngredientId,
                IngredientCode = ingredient.IngredientCode,
                IngredientName = ingredient.IngredientName,
                TargetWeight = ingredient.TargetWeight,
                ActualNetWeight = _activeSession.NetIngredientWeight,
                Scale2WeightBefore = scale2Before,
                Scale2WeightAfter = currentScale2Weight,
                TransferDeviation = deviation,
                TransferredAt = DateTime.Now,
                BowlCode = _activeSession.SelectedIngredientBowlCode ?? "",
                BowlType = ingredient.BowlSize,
                // ✅ CALCULATE tolerance range from target and tolerance value
                MinWeight = ingredient.TargetWeight - toleranceValue,  // ✅ FIXED!
                MaxWeight = ingredient.TargetWeight + toleranceValue,  // ✅ FIXED!
                ToleranceValue = toleranceValue
            };

            _activeSession.TransferredIngredients.Add(transferRecord);

            // ✅ STEP 8: Create WeighingDetail record
            await SaveWeighingDetailAsync(transferRecord);

            Console.WriteLine($"✅ Transfer verified and recorded!");
            Console.WriteLine($"   Ingredient: {ingredient.IngredientCode}");
            Console.WriteLine($"   Target: {ingredient.TargetWeight:F3} kg");
            Console.WriteLine($"   Actual Net: {_activeSession.NetIngredientWeight:F3} kg");
            Console.WriteLine($"   Scale 2 Cumulative: {actualScale2Net:F3} kg");
            Console.WriteLine($"   Within Tolerance: {transferRecord.IsWithinTolerance}");

            // ✅ STEP 9: Move to next ingredient
            _activeSession.CurrentIngredientIndex++;

            // ✅ STEP 10: Check if all ingredients for this repetition are complete
            if (_activeSession.CurrentIngredientIndex >= _activeSession.Ingredients.Count)
            {
                // All ingredients complete for this repetition
                await CompleteRepetitionAsync(batchId);

                // Check if all repetitions complete
                if (_activeSession == null)
                {
                    // Batch completed
                    return (true, "Batch completed!", deviation);
                }

                // More repetitions to do
                return (true, $"Repetition {_activeSession.CurrentRepetition - 1} complete! Starting repetition {_activeSession.CurrentRepetition}", deviation);
            }

            // ✅ NOT LAST INGREDIENT - Just reset ingredient bowl, mixing bowl stays
            Console.WriteLine($"➡️ Moving to next ingredient: {_activeSession.CurrentIngredient?.IngredientCode}");
            _activeSession.CurrentStage = WeighingStage.PlaceBowls;

            // Only reset INGREDIENT bowl
            _activeSession.SelectedIngredientBowlCode = null;
            _activeSession.SelectedIngredientBowlWeight = 0;
            _activeSession.IngredientBowlWeight = 0;
            _activeSession.NetIngredientWeight = 0;

            return (true, "Transfer completed successfully", deviation);
        }

        public async Task<bool> CompleteIngredientAsync(
            string batchId,
            decimal actualWeight,
            string bowlCode,
            string bowlType)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            var ingredient = _activeSession.CurrentIngredient;
            if (ingredient == null)
                return false;

            // This is called when ingredient weighing is complete
            // The actual saving happens in ConfirmTransferAsync
            return true;
        }

        public async Task<bool> CompleteRepetitionAsync(string batchId)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            Console.WriteLine($"✅ Repetition {_activeSession.CurrentRepetition} complete!");

            // Update batch progress
            await _batchService.UpdateRepetitionProgressAsync(
                batchId,
                _activeSession.CurrentRepetition
            );

            // ✅ CHECK COMPLETION BEFORE INCREMENTING
            if (_activeSession.CurrentRepetition >= _activeSession.TotalRepetitions)
            {
                // ✅ ALL REPETITIONS COMPLETE - Batch done!
                Console.WriteLine($"🎉 All {_activeSession.TotalRepetitions} repetitions complete! Batch finished!");
                await CompleteBatchAsync(batchId);
                return true;
            }

            // ✅ MORE REPETITIONS TO DO - Move to next repetition
            _activeSession.CurrentRepetition++;
            Console.WriteLine($"🔄 Starting repetition {_activeSession.CurrentRepetition} of {_activeSession.TotalRepetitions}");

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

        public async Task<bool> CompleteBatchAsync(string batchId)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
                return false;

            // ✅ Update WeighingRecord first
            await UpdateWeighingRecordOnCompletion();

            // ✅ DEFENSIVE: Get the most reliable operator name
            string completedByOperator = _activeSession.OperatorName;

            // If session operator is empty, try to get from WeighingRecord
            if (string.IsNullOrEmpty(completedByOperator))
            {
                Console.WriteLine("⚠️ WARNING: Session operator is empty! Checking WeighingRecord...");

                if (!string.IsNullOrEmpty(_activeSession.WeighingRecordId))
                {
                    var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
                    if (record != null && !string.IsNullOrEmpty(record.OperatorName))
                    {
                        completedByOperator = record.OperatorName;
                        Console.WriteLine($"✅ Using operator from WeighingRecord: {completedByOperator}");
                    }
                }
            }

            // If still empty, try to get from batch
            if (string.IsNullOrEmpty(completedByOperator))
            {
                Console.WriteLine("⚠️ WARNING: Still no operator! Checking Batch.StartedBy...");
                var batch = await _batchService.GetBatchByIdAsync(batchId);
                if (batch != null && !string.IsNullOrEmpty(batch.StartedBy))
                {
                    completedByOperator = batch.StartedBy;
                    Console.WriteLine($"✅ Using operator from Batch.StartedBy: {completedByOperator}");
                }
            }

            // Final fallback
            if (string.IsNullOrEmpty(completedByOperator))
            {
                completedByOperator = "Unknown Operator";
                Console.WriteLine("❌ ERROR: No operator found anywhere! Using fallback.");
            }

            Console.WriteLine($"🎯 Completing batch {batchId} with operator: {completedByOperator}");

            // ✅ Complete batch with the operator name
            await _batchService.CompleteBatchAsync(batchId, completedByOperator);

            // Clear session
            _activeSession = null;

            Console.WriteLine($"🎉 Batch {batchId} completed successfully by {completedByOperator}!");
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

            // ✅ Update WeighingRecord as aborted
            if (!string.IsNullOrEmpty(_activeSession.WeighingRecordId))
            {
                var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
                if (record != null)
                {
                    record.Status = WeighingRecordStatus.Aborted;
                    record.AbortReason = reason;
                    record.AbortedBy = abortedBy;
                    record.AbortedDate = DateTime.Now;
                    record.SessionEndTime = DateTime.Now;
                    record.CompletedRepetitions = _activeSession.CurrentRepetition - 1;

                    await _reportService.UpdateWeighingRecordAsync(record);
                }
            }

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

        // ✅ NEW: Cumulative Tolerance Report
        public (bool withinTolerance, string report, decimal overallDeviation) GetCumulativeToleranceReport()
        {
            if (_activeSession == null)
                return (true, "No active session", 0);

            var report = new StringBuilder();
            report.AppendLine("📊 CUMULATIVE TOLERANCE REPORT");
            report.AppendLine($"Repetition {_activeSession.CurrentRepetition} of {_activeSession.TotalRepetitions}");
            report.AppendLine($"Batch: {_activeSession.BatchId}");
            report.AppendLine($"Recipe: {_activeSession.RecipeCode}");
            report.AppendLine(new string('=', 60));
            report.AppendLine();

            decimal totalTarget = 0;
            decimal totalActual = 0;
            bool allWithinTolerance = true;
            int outOfToleranceCount = 0;

            var currentRepTransfers = _activeSession.TransferredIngredients
                .Where(t => t.RepetitionNumber == _activeSession.CurrentRepetition)
                .OrderBy(t => t.IngredientSequence)
                .ToList();

            if (!currentRepTransfers.Any())
            {
                report.AppendLine("No ingredients transferred yet for this repetition.");
                return (true, report.ToString(), 0);
            }

            foreach (var transfer in currentRepTransfers)
            {
                totalTarget += transfer.TargetWeight;
                totalActual += transfer.ActualNetWeight;

                var deviation = transfer.ActualNetWeight - transfer.TargetWeight;
                var deviationPercent = transfer.TargetWeight > 0
                    ? (deviation / transfer.TargetWeight) * 100
                    : 0;

                var withinTolerance = transfer.IsWithinTolerance;

                if (!withinTolerance)
                {
                    allWithinTolerance = false;
                    outOfToleranceCount++;
                }

                // Format the line
                var statusIcon = withinTolerance ? "✓" : "⚠️";
                var deviationSign = deviation >= 0 ? "+" : "";

                report.AppendLine($"{transfer.IngredientSequence}. {transfer.IngredientCode}");
                report.AppendLine($"   Target:    {transfer.TargetWeight:F3} kg");
                report.AppendLine($"   Actual:    {transfer.ActualNetWeight:F3} kg");
                report.AppendLine($"   Deviation: {deviationSign}{deviation:F3} kg ({deviationSign}{deviationPercent:F2}%) {statusIcon}");
                report.AppendLine($"   Range:     {transfer.MinWeight:F3} - {transfer.MaxWeight:F3} kg");
                report.AppendLine();
            }

            report.AppendLine(new string('-', 60));
            report.AppendLine($"TOTALS:");
            report.AppendLine($"Target Total:  {totalTarget:F3} kg");
            report.AppendLine($"Actual Total:  {totalActual:F3} kg");

            var overallDeviation = totalActual - totalTarget;
            var overallDeviationPercent = totalTarget > 0
                ? (overallDeviation / totalTarget) * 100
                : 0;
            var overallSign = overallDeviation >= 0 ? "+" : "";

            report.AppendLine($"Overall Dev:   {overallSign}{overallDeviation:F3} kg ({overallSign}{overallDeviationPercent:F2}%)");
            report.AppendLine();
            report.AppendLine($"Ingredients Within Tolerance: {currentRepTransfers.Count - outOfToleranceCount}/{currentRepTransfers.Count}");
            report.AppendLine();

            if (allWithinTolerance)
            {
                report.AppendLine("✅ ALL INGREDIENTS WITHIN TOLERANCE");
            }
            else
            {
                report.AppendLine($"⚠️ {outOfToleranceCount} INGREDIENT(S) OUT OF TOLERANCE");
            }

            return (allWithinTolerance, report.ToString(), overallDeviation);
        }

        public List<TransferredIngredient> GetTransferHistory(int? repetitionNumber = null)
        {
            if (_activeSession == null)
                return new List<TransferredIngredient>();

            if (repetitionNumber.HasValue)
            {
                return _activeSession.TransferredIngredients
                    .Where(t => t.RepetitionNumber == repetitionNumber.Value)
                    .OrderBy(t => t.IngredientSequence)
                    .ToList();
            }

            return _activeSession.TransferredIngredients
                .OrderBy(t => t.RepetitionNumber)
                .ThenBy(t => t.IngredientSequence)
                .ToList();
        }

        public bool IsScale2Stable(List<decimal> recentReadings, decimal tolerance = SCALE_STABILITY_TOLERANCE)
        {
            if (recentReadings == null || recentReadings.Count < 3)
                return false;

            var max = recentReadings.Max();
            var min = recentReadings.Min();
            var range = max - min;

            return range <= tolerance;
        }

        public decimal CalculateDynamicTransferTolerance(int ingredientsTransferred)
        {
            return BASE_TRANSFER_TOLERANCE + (PER_INGREDIENT_TOLERANCE * ingredientsTransferred);
        }

        // Helper method kept for compatibility
        public (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatus(decimal currentWeight, RecipeIngredient ingredient)
        {
            // This uses absolute weight from scale (not net)
            // Keeping for any legacy code that might use it
            return GetIngredientStatusByNet(currentWeight, ingredient);
        }

        // ✅ PRIVATE: Generate unique record ID
        private async Task<string> GenerateRecordIdAsync()
        {
            var allRecords = await _reportService.GetAllWeighingRecordsAsync();
            var maxId = 0;

            foreach (var record in allRecords)
            {
                if (record.RecordId.StartsWith("RECORD") &&
                    int.TryParse(record.RecordId.Substring(6), out int id))
                {
                    maxId = Math.Max(maxId, id);
                }
            }

            return $"RECORD{(maxId + 1):D3}";
        }

        // ✅ PRIVATE: Save WeighingDetail
        private async Task SaveWeighingDetailAsync(TransferredIngredient transfer)
        {
            if (_activeSession == null || string.IsNullOrEmpty(_activeSession.WeighingRecordId))
                return;

            var detail = new WeighingDetail
            {
                DetailId = await GenerateDetailIdAsync(),
                RecordId = _activeSession.WeighingRecordId,
                BatchId = _activeSession.BatchId,
                RepetitionNumber = transfer.RepetitionNumber,
                IngredientSequence = transfer.IngredientSequence,
                IngredientId = transfer.IngredientId,
                IngredientCode = transfer.IngredientCode,
                IngredientName = transfer.IngredientName,
                TargetWeight = transfer.TargetWeight,
                ActualWeight = transfer.ActualNetWeight,
                MinWeight = transfer.MinWeight,
                MaxWeight = transfer.MaxWeight,
                ToleranceValue = transfer.ToleranceValue,
                BowlCode = transfer.BowlCode,
                BowlType = transfer.BowlType,
                ScaleNumber = 1,
                Unit = "kg",
                Timestamp = transfer.TransferredAt,
                // ✅ NEW: Add Scale 2 tracking data
                Scale2WeightBefore = transfer.Scale2WeightBefore,
                Scale2WeightAfter = transfer.Scale2WeightAfter,
                TransferDeviation = transfer.TransferDeviation
            };

            await _reportService.SaveWeighingDetailAsync(detail);
            Console.WriteLine($"✅ WeighingDetail saved: {detail.DetailId}");
        }

        // ✅ PRIVATE: Generate unique detail ID
        private async Task<string> GenerateDetailIdAsync()
        {
            if (string.IsNullOrEmpty(_activeSession?.WeighingRecordId))
                return "DETAIL0001";

            var details = await _reportService.GetWeighingDetailsByRecordIdAsync(_activeSession.WeighingRecordId);
            var maxId = 0;

            foreach (var detail in details)
            {
                if (detail.DetailId.StartsWith("DETAIL") &&
                    int.TryParse(detail.DetailId.Substring(6), out int id))
                {
                    maxId = Math.Max(maxId, id);
                }
            }

            return $"DETAIL{(maxId + 1):D4}";
        }

        // ✅ PRIVATE: Update WeighingRecord on completion
        private async Task UpdateWeighingRecordOnCompletion()
        {
            if (_activeSession == null || string.IsNullOrEmpty(_activeSession.WeighingRecordId))
                return;

            var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
            if (record == null)
                return;

            // Calculate statistics
            var allDetails = await _reportService.GetWeighingDetailsByRecordIdAsync(_activeSession.WeighingRecordId);

            record.SessionEndTime = DateTime.Now;
            record.Status = WeighingRecordStatus.Completed;
            record.CompletedRepetitions = _activeSession.TotalRepetitions;
            record.TotalIngredientsWeighed = allDetails.Count;
            record.IngredientsWithinTolerance = allDetails.Count(d => d.IsWithinTolerance);
            record.IngredientsOutOfTolerance = allDetails.Count(d => !d.IsWithinTolerance);

            if (allDetails.Any())
            {
                record.AverageDeviation = allDetails.Average(d => Math.Abs(d.Deviation));
                record.MaxDeviation = allDetails.Max(d => Math.Abs(d.Deviation));
            }

            await _reportService.UpdateWeighingRecordAsync(record);
            Console.WriteLine($"✅ WeighingRecord updated: {record.RecordId}");
        }
    }
}