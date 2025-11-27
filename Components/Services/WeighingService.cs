using FirstWeigh.Models;
using Microsoft.Extensions.Logging;
using System.Text;

namespace FirstWeigh.Services
{
    public class WeighingService : IWeighingService
    {
        private readonly IBatchService _batchService;
        private readonly RecipeService _recipeService;
        private readonly ReportService _reportService;
        private readonly ILogger<WeighingService> _logger;
        private WeighingSession? _activeSession;

        // Configuration constants
        private const decimal BASE_TRANSFER_TOLERANCE = 0.050m;
        private const decimal PER_INGREDIENT_TOLERANCE = 0.015m;
        private const decimal BOWL_VERIFICATION_TOLERANCE_PERCENT = 2.0m;
        private const decimal SCALE_STABILITY_TOLERANCE = 0.005m;

        public WeighingService(
            IBatchService batchService,
            RecipeService recipeService,
            ReportService reportService,
            ILogger<WeighingService> logger)
        {
            _batchService = batchService;
            _recipeService = recipeService;
            _reportService = reportService;
            _logger = logger;
        }

        public async Task<bool> UpdateSessionOperatorAsync(string batchId, string operatorName)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
            {
                _logger.LogWarning("UpdateSessionOperator failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _activeSession.OperatorName = operatorName;

            if (!string.IsNullOrEmpty(_activeSession.WeighingRecordId))
            {
                var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
                if (record != null)
                {
                    record.OperatorName = operatorName;
                    await _reportService.UpdateWeighingRecordAsync(record);
                    _logger.LogInformation("WeighingRecord operator updated to: {OperatorName}", operatorName);
                    return true;
                }
            }

            return false;
        }

        public async Task<WeighingSession?> StartWeighingSessionAsync(string batchId)
        {
            _logger.LogInformation("Starting weighing session for batch {BatchId}", batchId);

            var batch = await _batchService.GetBatchByIdAsync(batchId);
            if (batch == null || batch.Status != "InProgress")
            {
                _logger.LogWarning("Cannot start session - Batch {BatchId} not found or not in progress", batchId);
                return null;
            }

            var recipe = await _recipeService.GetRecipeByIdAsync(batch.RecipeId);
            if (recipe == null)
            {
                _logger.LogWarning("Cannot start session - Recipe {RecipeId} not found", batch.RecipeId);
                return null;
            }

            var ingredients = await _recipeService.GetRecipeIngredientsAsync(batch.RecipeId);
            if (ingredients == null || !ingredients.Any())
            {
                _logger.LogWarning("Cannot start session - No ingredients found for recipe {RecipeId}", batch.RecipeId);
                return null;
            }

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

            _logger.LogInformation("Weighing session started - Record: {RecordId}, Batch: {BatchId}, Recipe: {RecipeName}",
                record.RecordId, batchId, recipe.RecipeName);
            return _activeSession;
        }

        public bool SelectBowls(string batchId, string ingredientBowlCode, decimal ingredientBowlWeight,
                                string mixingBowlCode, decimal mixingBowlWeight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
            {
                _logger.LogWarning("SelectBowls failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _activeSession.SelectedIngredientBowlCode = ingredientBowlCode;
            _activeSession.SelectedIngredientBowlWeight = ingredientBowlWeight;
            _activeSession.SelectedMixingBowlCode = mixingBowlCode;
            _activeSession.SelectedMixingBowlWeight = mixingBowlWeight;
            _activeSession.MixingBowlWeightBefore = mixingBowlWeight;

            _logger.LogInformation("Bowls selected - Ingredient: {IngredientBowl} ({IngredientWeight:F3} kg), Mixing: {MixingBowl} ({MixingWeight:F3} kg)",
                ingredientBowlCode, ingredientBowlWeight, mixingBowlCode, mixingBowlWeight);
            return true;
        }

        // ✅ UPDATED: Now uses percentage-based tolerance
        public (bool isValid, string message) VerifyBowlWeight(
            decimal actualWeight,
            decimal recordedWeight,
            string bowlCode,
            decimal tolerancePercent = BOWL_VERIFICATION_TOLERANCE_PERCENT)
        {
            // Calculate tolerance based on percentage of recorded weight
            decimal toleranceAmount = (recordedWeight * tolerancePercent) / 100m;

            // Minimum tolerance of 10g to handle very light bowls
            toleranceAmount = Math.Max(toleranceAmount, 0.010m);

            var difference = Math.Abs(actualWeight - recordedWeight);

            _logger.LogDebug("Bowl verification - Bowl: {BowlCode}, Recorded: {RecordedWeight:F3} kg, Actual: {ActualWeight:F3} kg, Tolerance: {TolerancePercent}% (±{ToleranceAmount:F3} kg)",
                bowlCode, recordedWeight, actualWeight, tolerancePercent, toleranceAmount);

            if (difference <= toleranceAmount)
            {
                _logger.LogInformation("Bowl {BowlCode} verified successfully - Actual: {ActualWeight:F3} kg, Expected: {RecordedWeight:F3} kg",
                    bowlCode, actualWeight, recordedWeight);
                return (true, $"✓ Bowl {bowlCode} verified: {actualWeight:F3} kg");
            }
            else
            {
                _logger.LogWarning("Bowl {BowlCode} verification FAILED - Actual: {ActualWeight:F3} kg, Expected: {RecordedWeight:F3} kg, Difference: {Difference:F3} kg",
                    bowlCode, actualWeight, recordedWeight, difference);
                return (false,
                    $"⚠ Bowl {bowlCode} weight mismatch!\n" +
                    $"Expected: {recordedWeight:F3} kg\n" +
                    $"Actual: {actualWeight:F3} kg\n" +
                    $"Difference: {difference:F3} kg (max: ±{toleranceAmount:F3} kg / {tolerancePercent}%)");
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
            {
                _logger.LogWarning("RecordBowlWeights failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _activeSession.IngredientBowlWeight = ingredientBowlWeight;
            _activeSession.MixingBowlWeightBefore = mixingBowlWeight;
            _activeSession.CurrentStage = WeighingStage.WeighIngredient;

            _logger.LogInformation("Bowls recorded - Ingredient bowl: {IngredientWeight:F3} kg, Mixing bowl: {MixingWeight:F3} kg",
                ingredientBowlWeight, mixingBowlWeight);
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
            {
                _logger.LogWarning("ReadyToTransfer failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _activeSession.NetIngredientWeight = netWeight;

            _logger.LogInformation("Ready to transfer - Net weight: {NetWeight:F3} kg, Ingredient: {IngredientCode}",
                netWeight, _activeSession.CurrentIngredient?.IngredientCode);
            return true;
        }

        public async Task<(bool success, string message, decimal deviation)> ConfirmTransferAsync(string batchId, decimal currentScale2Weight)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
            {
                _logger.LogWarning("ConfirmTransfer failed - No active session for batch {BatchId}", batchId);
                return (false, "No active session", 0);
            }

            var ingredient = _activeSession.CurrentIngredient;
            if (ingredient == null)
            {
                _logger.LogWarning("ConfirmTransfer failed - No current ingredient");
                return (false, "No current ingredient", 0);
            }

            decimal expectedCumulative = _activeSession.TransferredIngredients
                .Where(t => t.RepetitionNumber == _activeSession.CurrentRepetition)
                .Sum(t => t.ActualNetWeight);

            expectedCumulative += _activeSession.NetIngredientWeight;

            decimal scale2Before = _activeSession.MixingBowlWeightBefore;
            decimal actualScale2Net = currentScale2Weight - scale2Before;

            decimal deviation = actualScale2Net - expectedCumulative;
            decimal absDeviation = Math.Abs(deviation);

            int ingredientsTransferred = _activeSession.CurrentIngredientIndex + 1;
            decimal allowedTolerance = CalculateDynamicTransferTolerance(ingredientsTransferred);

            _logger.LogDebug("Transfer Verification - Expected: {Expected:F3} kg, Actual: {Actual:F3} kg, Deviation: {Deviation:F3} kg, Tolerance: ±{Tolerance:F3} kg",
                expectedCumulative, actualScale2Net, deviation, allowedTolerance);

            if (absDeviation > allowedTolerance)
            {
                _logger.LogWarning("Transfer verification FAILED - Deviation {Deviation:F3} kg exceeds tolerance ±{Tolerance:F3} kg",
                    deviation, allowedTolerance);

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

            decimal toleranceValue = (ingredient.TargetWeight * ingredient.TolerancePercentage) / 100;

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
                MinWeight = ingredient.TargetWeight - toleranceValue,
                MaxWeight = ingredient.TargetWeight + toleranceValue,
                ToleranceValue = toleranceValue
            };

            _activeSession.TransferredIngredients.Add(transferRecord);

            await SaveWeighingDetailAsync(transferRecord);

            _logger.LogInformation("Transfer verified - Ingredient: {IngredientCode}, Target: {Target:F3} kg, Actual: {Actual:F3} kg, Within Tolerance: {WithinTolerance}",
                ingredient.IngredientCode, ingredient.TargetWeight, _activeSession.NetIngredientWeight, transferRecord.IsWithinTolerance);

            _activeSession.CurrentIngredientIndex++;

            if (_activeSession.CurrentIngredientIndex >= _activeSession.Ingredients.Count)
            {
                await CompleteRepetitionAsync(batchId);

                if (_activeSession == null)
                {
                    return (true, "Batch completed!", deviation);
                }

                return (true, $"Repetition {_activeSession.CurrentRepetition - 1} complete! Starting repetition {_activeSession.CurrentRepetition}", deviation);
            }

            _logger.LogInformation("Moving to next ingredient: {IngredientCode}", _activeSession.CurrentIngredient?.IngredientCode);
            _activeSession.CurrentStage = WeighingStage.PlaceBowls;

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

            return true;
        }

        public async Task<bool> CompleteRepetitionAsync(string batchId)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
            {
                _logger.LogWarning("CompleteRepetition failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _logger.LogInformation("Repetition {CurrentRep} of {TotalReps} complete for batch {BatchId}",
                _activeSession.CurrentRepetition, _activeSession.TotalRepetitions, batchId);

            await _batchService.UpdateRepetitionProgressAsync(
                batchId,
                _activeSession.CurrentRepetition
            );

            if (_activeSession.CurrentRepetition >= _activeSession.TotalRepetitions)
            {
                _logger.LogInformation("All {TotalReps} repetitions complete - Batch {BatchId} finished!",
                    _activeSession.TotalRepetitions, batchId);
                await CompleteBatchAsync(batchId);
                return true;
            }

            _activeSession.CurrentRepetition++;
            _logger.LogInformation("Starting repetition {CurrentRep} of {TotalReps}",
                _activeSession.CurrentRepetition, _activeSession.TotalRepetitions);

            _activeSession.CurrentIngredientIndex = 0;
            _activeSession.CurrentStage = WeighingStage.PlaceBowls;

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
            {
                _logger.LogWarning("CompleteBatch failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            await UpdateWeighingRecordOnCompletion();

            string completedByOperator = _activeSession.OperatorName;

            if (string.IsNullOrEmpty(completedByOperator))
            {
                _logger.LogWarning("Session operator is empty - Checking WeighingRecord...");

                if (!string.IsNullOrEmpty(_activeSession.WeighingRecordId))
                {
                    var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
                    if (record != null && !string.IsNullOrEmpty(record.OperatorName))
                    {
                        completedByOperator = record.OperatorName;
                        _logger.LogInformation("Using operator from WeighingRecord: {Operator}", completedByOperator);
                    }
                }
            }

            if (string.IsNullOrEmpty(completedByOperator))
            {
                _logger.LogWarning("Still no operator - Checking Batch.StartedBy...");
                var batch = await _batchService.GetBatchByIdAsync(batchId);
                if (batch != null && !string.IsNullOrEmpty(batch.StartedBy))
                {
                    completedByOperator = batch.StartedBy;
                    _logger.LogInformation("Using operator from Batch.StartedBy: {Operator}", completedByOperator);
                }
            }

            if (string.IsNullOrEmpty(completedByOperator))
            {
                completedByOperator = "Unknown Operator";
                _logger.LogError("No operator found anywhere - Using fallback");
            }

            _logger.LogInformation("Completing batch {BatchId} with operator: {Operator}", batchId, completedByOperator);

            await _batchService.CompleteBatchAsync(batchId, completedByOperator);

            _activeSession = null;

            _logger.LogInformation("Batch {BatchId} completed successfully by {Operator}", batchId, completedByOperator);
            return true;
        }

        public Task<bool> PauseSessionAsync(string batchId)
        {
            _logger.LogInformation("Pausing session for batch {BatchId}", batchId);
            _activeSession = null;
            return Task.FromResult(true);
        }

        public async Task<bool> AbortSessionAsync(string batchId, string reason, string abortedBy)
        {
            if (_activeSession == null || _activeSession.BatchId != batchId)
            {
                _logger.LogWarning("AbortSession failed - No active session for batch {BatchId}", batchId);
                return false;
            }

            _logger.LogWarning("Aborting batch {BatchId} - Reason: {Reason}, By: {AbortedBy}", batchId, reason, abortedBy);

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
            _logger.LogInformation("Clearing active session");
            _activeSession = null;
        }

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

        public (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatus(decimal currentWeight, RecipeIngredient ingredient)
        {
            return GetIngredientStatusByNet(currentWeight, ingredient);
        }

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
                Scale2WeightBefore = transfer.Scale2WeightBefore,
                Scale2WeightAfter = transfer.Scale2WeightAfter,
                TransferDeviation = transfer.TransferDeviation
            };

            await _reportService.SaveWeighingDetailAsync(detail);
            _logger.LogDebug("WeighingDetail saved: {DetailId}", detail.DetailId);
        }

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

        private async Task UpdateWeighingRecordOnCompletion()
        {
            if (_activeSession == null || string.IsNullOrEmpty(_activeSession.WeighingRecordId))
                return;

            var record = await _reportService.GetWeighingRecordByIdAsync(_activeSession.WeighingRecordId);
            if (record == null)
                return;

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
            _logger.LogInformation("WeighingRecord {RecordId} updated - Status: {Status}, Compliance: {Compliance:F1}%",
                record.RecordId, record.Status, record.CompliancePercentage);
        }
    }
}