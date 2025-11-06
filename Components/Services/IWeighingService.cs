using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    public interface IWeighingService
    {
        // Session Management
        Task<WeighingSession?> StartWeighingSessionAsync(string batchId);
        WeighingSession? GetActiveSession();
        void ClearActiveSession();
        Task<bool> PauseSessionAsync(string batchId);
        Task<bool> AbortSessionAsync(string batchId, string reason, string abortedBy);
        Task<bool> UpdateSessionOperatorAsync(string batchId, string operatorName);

        // Bowl Selection & Verification
        bool SelectBowls(string batchId, string ingredientBowlCode, decimal ingredientBowlWeight,
                        string mixingBowlCode, decimal mixingBowlWeight);
        (bool isValid, string message) VerifyBowlWeight(decimal actualWeight, decimal recordedWeight,
                                                        string bowlCode, decimal tolerance = 0.05m);
        bool RecordBowlWeights(string batchId, decimal ingredientBowlWeight, decimal mixingBowlWeight);

        // Weight Calculations
        decimal GetNetIngredientWeight(decimal currentScale1Weight);
        (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatusByNet(decimal netWeight, RecipeIngredient ingredient);

        // Transfer Process
        Task<bool> ReadyToTransferAsync(string batchId, decimal netWeight);
        Task<(bool success, string message, decimal deviation)> ConfirmTransferAsync(
            string batchId, decimal currentScale2Weight);

        // ✅ NEW: Completion & Recording
        Task<bool> CompleteIngredientAsync(string batchId, decimal actualWeight,
                                           string bowlCode, string bowlType);
        Task<bool> CompleteRepetitionAsync(string batchId);
        Task<bool> CompleteBatchAsync(string batchId);

        // ✅ NEW: Cumulative Tolerance Reporting
        (bool withinTolerance, string report, decimal overallDeviation) GetCumulativeToleranceReport();
        List<TransferredIngredient> GetTransferHistory(int? repetitionNumber = null);

        // ✅ NEW: Scale Stability
        bool IsScale2Stable(List<decimal> recentReadings, decimal tolerance = 0.005m);

        // ✅ NEW: Dynamic Tolerance Calculation
        decimal CalculateDynamicTransferTolerance(int ingredientsTransferred);

        // Existing helper (kept for compatibility)
        (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatus(decimal currentWeight, RecipeIngredient ingredient);
    }
}