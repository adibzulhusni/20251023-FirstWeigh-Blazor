// WeighingRecordModel.txt

namespace FirstWeigh.Models
{
    // Master record for the entire batch weighing session
    public class WeighingRecord
    {
        public string RecordId { get; set; } = string.Empty; // RECORD001, RECORD002...
        public string BatchId { get; set; } = string.Empty;
        public string RecipeId { get; set; } = string.Empty;
        public string RecipeCode { get; set; } = string.Empty;
        public string RecipeName { get; set; } = string.Empty;
        public string OperatorName { get; set; } = string.Empty;

        // Session timing
        public DateTime SessionStartTime { get; set; }
        public DateTime? SessionEndTime { get; set; }
        public TimeSpan Duration => SessionEndTime.HasValue
            ? SessionEndTime.Value - SessionStartTime
            : TimeSpan.Zero;

        // Batch details
        public int TotalRepetitions { get; set; }
        public int CompletedRepetitions { get; set; }
        public string Status { get; set; } = WeighingRecordStatus.InProgress; // InProgress, Completed, Aborted

        // Abort information (if applicable)
        public string? AbortReason { get; set; }
        public string? AbortedBy { get; set; }
        public DateTime? AbortedDate { get; set; }

        // Quality metrics (calculated)
        public int TotalIngredientsWeighed { get; set; } // Total count
        public int IngredientsWithinTolerance { get; set; }
        public int IngredientsOutOfTolerance { get; set; }
        public decimal AverageDeviation { get; set; } // Average ± kg
        public decimal MaxDeviation { get; set; } // Worst deviation
        public decimal CompliancePercentage => TotalIngredientsWeighed > 0
            ? (decimal)IngredientsWithinTolerance / TotalIngredientsWeighed * 100
            : 0;

        // Audit trail
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public string CreatedBy { get; set; } = string.Empty;

        // Navigation property
        public List<WeighingDetail> Details { get; set; } = new();
    }

    // Detailed record for each ingredient weighed
    public class WeighingDetail
    {
        public string DetailId { get; set; } = string.Empty; // DETAIL0001, DETAIL0002...
        public string RecordId { get; set; } = string.Empty; // Foreign key to WeighingRecord
        public string BatchId { get; set; } = string.Empty;

        // Repetition and sequence
        public int RepetitionNumber { get; set; }
        public int IngredientSequence { get; set; }

        // Ingredient information
        public string IngredientId { get; set; } = string.Empty;
        public string IngredientCode { get; set; } = string.Empty;
        public string IngredientName { get; set; } = string.Empty;

        // Weight data
        public decimal TargetWeight { get; set; }
        public decimal ActualWeight { get; set; }
        public decimal Deviation => ActualWeight - TargetWeight;
        public decimal DeviationPercentage => TargetWeight > 0
            ? (Deviation / TargetWeight) * 100
            : 0;

        // Tolerance
        public decimal MinWeight { get; set; }
        public decimal MaxWeight { get; set; }
        public decimal ToleranceValue { get; set; } // The actual ± kg tolerance
        public bool IsWithinTolerance => ActualWeight >= MinWeight && ActualWeight <= MaxWeight;

        // Equipment used
        public string BowlCode { get; set; } = string.Empty;
        public string BowlType { get; set; } = string.Empty;
        public int ScaleNumber { get; set; } = 1;
        public string Unit { get; set; } = "kg";

        // Timing
        public DateTime Timestamp { get; set; } = DateTime.Now;

        // Computed properties for display
        public string StatusIcon => IsWithinTolerance ? "✓" : "⚠";
        public string StatusColor => IsWithinTolerance ? "green" : "orange";
        public string StatusText => IsWithinTolerance ? "Within Tolerance" : "Out of Tolerance";
    }

    public static class WeighingRecordStatus
    {
        public const string InProgress = "In Progress";
        public const string Completed = "Completed";
        public const string Aborted = "Aborted";
    }
}