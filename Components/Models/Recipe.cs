namespace FirstWeigh.Models
{
    public class Recipe
    {
        public string RecipeId { get; set; } = string.Empty;
        public string RecipeCode { get; set; } = string.Empty;
        public string RecipeName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Status { get; set; } = RecipeStatus.Active;
        public DateTime CreatedDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public string CreatedBy { get; set; } = string.Empty;
        public string LastModifiedBy { get; set; } = string.Empty;

        // Navigation property
        public List<RecipeIngredient> Ingredients { get; set; } = new();
    }


        public class RecipeIngredient
        {
            public string RecipeId { get; set; } = string.Empty;
            public int Sequence { get; set; }
            public string IngredientId { get; set; } = string.Empty;
            public string IngredientCode { get; set; } = string.Empty;
            public string IngredientName { get; set; } = string.Empty;
            public decimal TargetWeight { get; set; }
            public decimal TolerancePercentage { get; set; }
            public int ScaleNumber { get; set; } = 1;
            public string Unit { get; set; } = "kg";

            // ✅ NEW: Bowl size for this ingredient
            public string BowlSize { get; set; } = "Medium"; // Small, Medium, Large

            // NEW: Properties for searchable dropdown
             public string SearchQuery { get; set; } = string.Empty;
             public bool ShowDropdown { get; set; } = false;
             public List<Ingredient> FilteredIngredients { get; set; } = new();


        // Calculated properties
        public decimal MinWeight => TargetWeight - ((TargetWeight * TolerancePercentage) / 100);
            public decimal MaxWeight => TargetWeight + ((TargetWeight * TolerancePercentage) / 100);
        }

        // Bowl size constants for consistency
        public static class BowlSize
        {
            public const string Small = "Small";
            public const string Medium = "Medium";
            public const string Large = "Large";
        }


    public static class RecipeStatus
    {
        public const string Active = "Active";
        public const string Archived = "Archived";
        public const string Draft = "Draft";
    }
}