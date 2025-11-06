using FirstWeigh.Models;

namespace FirstWeigh.Services
{
    public interface IBatchService
    {
        Task<List<Batch>> GetAllBatchesAsync();
        Task<Batch?> GetBatchByIdAsync(string batchId);
        Task<List<Batch>> GetBatchesByStatusAsync(string status);
        Task<List<Batch>> GetActiveBatchesAsync(); // InProgress status
        Task<List<Batch>> GetPendingBatchesAsync();
        Task<string> CreateBatchAsync(Batch batch);
        Task<bool> UpdateBatchAsync(Batch batch);
        Task<bool> DeleteBatchAsync(string batchId);
        Task<bool> StartBatchAsync(string batchId, string startedBy);
        Task<bool> AbortBatchAsync(string batchId, string abortedBy, string abortReason);
        Task<bool> UpdateRepetitionProgressAsync(string batchId, int currentRepetition);
        Task<int> GetActiveBatchCountAsync();
        Task<bool> CanStartBatch(); // Check if less than 5 active batches
        Task<bool> CompleteBatchAsync(string batchId, string completedBy);  // ✅ CORRECT - Add second parameter
    }
}