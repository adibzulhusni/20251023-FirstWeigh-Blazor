namespace FirstWeigh.Models
{
    public class Ingredient
    {
        public string IngredientId { get; set; } = string.Empty;
        public string IngredientCode { get; set; } = string.Empty;
        public string IngredientName { get; set; } = string.Empty;
        public string PackingType { get; set; } = string.Empty;
        public string UnitOfMeasure { get; set; } = "kg";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime LastModifiedDate { get; set; } = DateTime.Now;
        public string LastModifiedBy { get; set; } = string.Empty;
    }

    public static class UnitOfMeasure
    {
        public const string Kilogram = "kg";
        public const string Gram = "g";
        public const string Liter = "L";
        public const string Milliliter = "mL";

        public static List<string> GetAllUnits()
        {
            return new List<string> { Kilogram, Gram, Liter, Milliliter };
        }
    }

    public static class PackingTypes
    {
        public const string Bag25kg = "25kg Bag";
        public const string Bag50kg = "50kg Bag";
        public const string Bottle1L = "1L Bottle";
        public const string Bottle5L = "5L Bottle";
        public const string Drum200L = "200L Drum";
        public const string Bulk = "Bulk";
        public const string Sack = "Sack";
        public const string Box = "Box";
        public const string Pallet = "Pallet";

        public static List<string> GetAllPackingTypes()
        {
            return new List<string>
            {
                Bag25kg,
                Bag50kg,
                Bottle1L,
                Bottle5L,
                Drum200L,
                Bulk,
                Sack,
                Box,
                Pallet
            };
        }
    }
}