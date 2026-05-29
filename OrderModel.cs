namespace DuctOrderApp.Models
{
    /// <summary>
    /// Represents a single duct order record.
    /// </summary>
    public class OrderModel
    {
        public string Client       { get; set; }
        public string OrderName    { get; set; }
        public string DuctType     { get; set; }
        public bool   Urgent       { get; set; }
        public string DateCreated  { get; set; }
        public bool   Done         { get; set; }
    }
}
