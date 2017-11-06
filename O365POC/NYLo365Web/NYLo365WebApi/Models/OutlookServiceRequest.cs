namespace NYLo365WebApi.Models
{
    public class OutlookServiceRequest
    {
        public string AttachmentToken { get; set; }
        public string EwsUrl { get; set; }
        public ArtifactDetails[] Attachments { get; set; }
    }
}