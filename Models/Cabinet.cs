// Models/Cabinet.cs
namespace WinNetConfigurator.Models
{
    public class Cabinet
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }
}
