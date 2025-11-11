namespace FirstWeigh.Models
{
    public class IngredientUsageSummary
    {
        public string IngredientCode { get; set; } = string.Empty;
        public string IngredientName { get; set; } = string.Empty;
        public decimal TotalWeightUsed { get; set; }
        public int TimesUsed { get; set; }
    }
}