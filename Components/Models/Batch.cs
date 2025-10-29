namespace FirstWeigh.Models
{
    public class Batch
    {
        public string BatchId { get; set; } = string.Empty;
        public string RecipeId { get; set; } = string.Empty;
        public string RecipeName { get; set; } = string.Empty; // Denormalized for display
        public int TotalRepetitions { get; set; }
        public int CurrentRepetition { get; set; } = 0;
        public string Status { get; set; } = "Pending"; // Pending, InProgress, Completed, Aborted

        // Scheduling
        public DateTime? PlannedStartTime { get; set; }
        public DateTime? PlannedEndTime { get; set; }

        // Audit fields
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; }
        public string? StartedBy { get; set; }
        public DateTime? StartedDate { get; set; }
        public DateTime? CompletedDate { get; set; }
        public string? AbortReason { get; set; }
        public string? AbortedBy { get; set; }
        public DateTime? AbortedDate { get; set; }
        public string? Notes { get; set; }

        // Computed properties
        public double ProgressPercentage => TotalRepetitions > 0
            ? (double)CurrentRepetition / TotalRepetitions * 100
            : 0;

        public bool RequiresSupervisorApproval => ProgressPercentage > 40;

        public string StatusBadgeClass => Status switch
        {
            "Pending" => "badge-warning",
            "InProgress" => "badge-primary",
            "Completed" => "badge-success",
            "Aborted" => "badge-danger",
            _ => "badge-secondary"
        };
    }
}