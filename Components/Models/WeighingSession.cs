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
        public DateTime? SessionStarted { get; set; } // ✅ ADD THIS
        public DateTime? PlannedStartTime { get; set; }    // ADD THIS LINE
        public DateTime? PlannedEndTime { get; set; }      // ADD THIS LINE
        public int CurrentRepetition { get; set; } = 1;
        public int TotalRepetitions { get; set; }
        public int CurrentIngredientIndex { get; set; } = 0;
        public List<RecipeIngredient> Ingredients { get; set; } = new();

        // ✅ NEW: Stage tracking
        public WeighingStage CurrentStage { get; set; } = WeighingStage.PlaceBowls;

        // ✅ NEW: Bowl weights recorded in Stage 1
        public decimal IngredientBowlWeight { get; set; } = 0;
        public decimal MixingBowlWeightBefore { get; set; } = 0;

        // ✅ NEW: Net ingredient weight (for verification)
        public decimal NetIngredientWeight { get; set; } = 0;

        // Computed properties
        public RecipeIngredient? CurrentIngredient =>
            Ingredients != null && CurrentIngredientIndex < Ingredients.Count
                ? Ingredients[CurrentIngredientIndex]
                : null;

        public int TotalIngredients => Ingredients?.Count ?? 0;

        public bool IsComplete =>
            CurrentRepetition > TotalRepetitions ||
            (CurrentRepetition == TotalRepetitions && CurrentIngredientIndex >= TotalIngredients);
        // ✅ NEW: Selected bowls for verification
        public string? SelectedIngredientBowlCode { get; set; }
        public decimal SelectedIngredientBowlWeight { get; set; }
        public string? SelectedMixingBowlCode { get; set; }
        public decimal SelectedMixingBowlWeight { get; set; }
    }

    // ✅ NEW: Weighing stages enum
    public enum WeighingStage
    {
        PlaceBowls = 1,      // Stage 1: Place bowls and record weights
        WeighIngredient = 2, // Stage 2: Add ingredient and weigh
        Transfer = 3         // Stage 3: Transfer to mixing bowl
    }
}