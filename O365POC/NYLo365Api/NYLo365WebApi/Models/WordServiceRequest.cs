namespace NYLo365WebApi.Models
{
    public class WordServiceRequest
    {
        public string ArtifactType { get; set; }
        public string ContentType { get; set; }
        public string Content { get; set; }
        public string Id { get; set; }
        public bool IsInline { get; set; }
        public ArtifactDetails[] Attachments { get; set; }
        public string Name { get; set; }
        public int Size { get; set; }
    }
}