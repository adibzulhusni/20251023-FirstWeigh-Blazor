namespace FirstWeigh.Models
{
    public class WeighingSession
    {
        public string BatchId { get; set; } = string.Empty;
        public string RecipeId { get; set; } = string.Empty;
        public string RecipeName { get; set; } = string.Empty;
        public string RecipeCode { get; set; } = string.Empty;
        public string OperatorName { get; set; } = string.Empty;
        public DateTime StartTime { get; set; }
        public DateTime? SessionStarted { get; set; }
        public DateTime? PlannedStartTime { get; set; }
        public DateTime? PlannedEndTime { get; set; }

        public int CurrentRepetition { get; set; } = 1;
        public int TotalRepetitions { get; set; }
        public int CurrentIngredientIndex { get; set; } = 0;
        public List<RecipeIngredient> Ingredients { get; set; } = new();

        // Stage tracking
        public WeighingStage CurrentStage { get; set; } = WeighingStage.PlaceBowls;

        // Bowl weights recorded in Stage 1
        public decimal IngredientBowlWeight { get; set; } = 0;
        public decimal MixingBowlWeightBefore { get; set; } = 0;

        // Net ingredient weight (for verification)
        public decimal NetIngredientWeight { get; set; } = 0;

        // ✅ NEW: Track actual transferred weights per ingredient
        public List<TransferredIngredient> TransferredIngredients { get; set; } = new();

        // Computed properties
        public RecipeIngredient? CurrentIngredient =>
            Ingredients != null && CurrentIngredientIndex < Ingredients.Count
                ? Ingredients[CurrentIngredientIndex]
                : null;

        public int TotalIngredients => Ingredients?.Count ?? 0;

        public bool IsComplete =>
            CurrentRepetition > TotalRepetitions ||
            (CurrentRepetition == TotalRepetitions && CurrentIngredientIndex >= TotalIngredients);

        // Selected bowls for verification
        public string? SelectedIngredientBowlCode { get; set; }
        public decimal SelectedIngredientBowlWeight { get; set; }
        public string? SelectedMixingBowlCode { get; set; }
        public decimal SelectedMixingBowlWeight { get; set; }

        // ✅ NEW: WeighingRecord tracking for this session
        public string? WeighingRecordId { get; set; }
    }

    // ✅ NEW: Record of each actual transfer
    public class TransferredIngredient
    {
        public int RepetitionNumber { get; set; }
        public int IngredientSequence { get; set; }
        public string IngredientId { get; set; } = string.Empty;
        public string IngredientCode { get; set; } = string.Empty;
        public string IngredientName { get; set; } = string.Empty;
        public decimal TargetWeight { get; set; }           // What recipe called for
        public decimal ActualNetWeight { get; set; }        // What we actually weighed on Scale 1
        public decimal Scale2WeightBefore { get; set; }     // Scale 2 before transfer
        public decimal Scale2WeightAfter { get; set; }      // Scale 2 after transfer
        public decimal TransferDeviation { get; set; }      // Difference in transfer
        public DateTime TransferredAt { get; set; }
        public string BowlCode { get; set; } = string.Empty;
        public string BowlType { get; set; } = string.Empty;

        // Tolerance checking
        public decimal MinWeight { get; set; }
        public decimal MaxWeight { get; set; }
        public decimal ToleranceValue { get; set; }
        public bool IsWithinTolerance => ActualNetWeight >= MinWeight && ActualNetWeight <= MaxWeight;
    }

    // Weighing stages enum
    public enum WeighingStage
    {
        PlaceBowls = 1,      // Stage 1: Place bowls and record weights
        WeighIngredient = 2, // Stage 2: Add ingredient and weigh
        Transfer = 3         // Stage 3: Transfer to mixing bowl
    }
}