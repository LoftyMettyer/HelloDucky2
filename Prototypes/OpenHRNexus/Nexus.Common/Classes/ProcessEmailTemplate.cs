namespace Nexus.Common.Classes
{
    public class ProcessEmailTemplate
    {
        public int Id { get; set; }
        public string To { get; set; }
        public string BodyTemplate { get; set; }
        public string Subject { get; set; }
    }
}
