namespace FirstWeigh.Models
{
    public class Bowl
    {
        public string BowlId { get; set; } = string.Empty;
        public string BowlCode { get; set; } = string.Empty;
        public string Category { get; set; } = BowlCategory.RegularBowl; // Regular Bowl or Mixing Bowl
        public string BowlType { get; set; } = string.Empty; // Large, Medium, Small (only for Regular Bowls)
        public decimal Weight { get; set; } = 0; // Measured bowl weight in kg (3 decimal places)
        public string Status { get; set; } = BowlStatus.Available;
        public string CurrentLocation { get; set; } = string.Empty;
        public DateTime? LastUsedDate { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime LastModifiedDate { get; set; } = DateTime.Now;
        public string LastModifiedBy { get; set; } = string.Empty;

        // Display name combining category and type
        public string DisplayType
        {
            get
            {
                if (Category == BowlCategory.MixingBowl)
                    return "Mixing Bowl";
                return BowlType;
            }
        }

        // Calculated property for status color
        public string StatusColor
        {
            get
            {
                return Status switch
                {
                    BowlStatus.Available => "success",
                    BowlStatus.InUse => "primary",
                    BowlStatus.Dirty => "warning",
                    BowlStatus.Maintenance => "danger",
                    _ => "secondary"
                };
            }
        }
    }

    public static class BowlCategory
    {
        public const string RegularBowl = "Regular Bowl";
        public const string MixingBowl = "Mixing Bowl";

        public static List<string> GetAllCategories()
        {
            return new List<string> { RegularBowl, MixingBowl };
        }
    }

    public static class BowlStatus
    {
        public const string Available = "Available";
        public const string InUse = "In Use";
        public const string Dirty = "Dirty";
        public const string Maintenance = "Maintenance";

        public static List<string> GetAllStatuses()
        {
            return new List<string> { Available, InUse, Dirty, Maintenance };
        }
    }

    public static class BowlType
    {
        public const string Large = "Large";
        public const string Medium = "Medium";
        public const string Small = "Small";
        public const string NotApplicable = "N/A"; // For Mixing Bowls

        public static List<string> GetAllTypes()
        {
            return new List<string> { Large, Medium, Small, NotApplicable };
        }
    }

    public class BowlChangeHistory
    {
        public string HistoryId { get; set; } = string.Empty;
        public string BowlId { get; set; } = string.Empty;
        public string BowlCode { get; set; } = string.Empty;
        public string ChangeType { get; set; } = string.Empty; // Status Change, Weight Change, Type Change, Category Change, Location Change
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty; // Required - explain why the change was made
        public DateTime ChangeDate { get; set; } = DateTime.Now;
        public string ChangedBy { get; set; } = string.Empty;
    }
}