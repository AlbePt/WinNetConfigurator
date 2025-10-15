namespace WinNetConfigurator.Models
{
    public class Cabinet
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public override string ToString() => Name;
    }
}
