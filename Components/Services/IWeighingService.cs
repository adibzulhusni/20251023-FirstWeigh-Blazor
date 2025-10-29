using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    public interface IWeighingService
    {
        bool SelectBowls(string batchId, string ingredientBowlCode, decimal ingredientBowlWeight,
                 string mixingBowlCode, decimal mixingBowlWeight);
        (bool isValid, string message) VerifyBowlWeight(decimal actualWeight, decimal recordedWeight, string bowlCode, decimal tolerance = 0.01m);
        bool RecordBowlWeights(string batchId, decimal ingredientBowlWeight, decimal mixingBowlWeight);
        decimal GetNetIngredientWeight(decimal currentScale1Weight);
        (string statusColor, string statusIcon, string statusMessage, bool canComplete) GetIngredientStatusByNet(decimal netWeight, RecipeIngredient ingredient);
        Task<bool> ReadyToTransferAsync(string batchId, decimal netWeight);
        Task<bool> ConfirmTransferAsync(string batchId, decimal currentScale2Weight);
        Task<WeighingSession?> StartWeighingSessionAsync(string batchId);
        Task<bool> CompleteIngredientAsync(string batchId, decimal actualWeight);
        Task<bool> PauseSessionAsync(string batchId);
        Task<bool> AbortSessionAsync(string batchId, string reason, string abortedBy);
        WeighingSession? GetActiveSession();
        void ClearActiveSession();

        // Scale calculation helpers
        (string statusColor, string statusIcon, string statusMessage, bool canComplete)
            GetIngredientStatus(decimal currentWeight, RecipeIngredient ingredient);
    }
}