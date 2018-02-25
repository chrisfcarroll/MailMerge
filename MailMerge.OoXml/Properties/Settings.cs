namespace MailMerge.OoXml.Properties
{
    public class Settings
    {
        public int MaximumOutputFileSize { get; set; } = 100*1000*1000;
        public string SomeSetting { get; set; }
        public string  TenantId { get; set; }
        public int MaximumMemoryStreamSize { get; set; } = 10*1000*1000;
    }
}
