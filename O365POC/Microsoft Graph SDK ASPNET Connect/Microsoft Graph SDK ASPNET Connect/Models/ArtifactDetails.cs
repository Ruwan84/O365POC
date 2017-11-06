using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Models
{
    public class ArtifactDetails
    {
        public string ArtifactType { get; set; }
        public string ContentType { get; set; }
        public string Content { get; set; }
        public string Id { get; set; }
        public bool IsInline { get; set; }
        public string Name { get; set; }
        public int Size { get; set; }

        public string Function { get; set; }
        public string DocumentType { get; set; }
        public string LineOfBusiness { get; set; }
        public string BusinessArea { get; set; }
        public string SubBusinessArea { get; set; }
        public string SubFunction { get; set; }
        public string Tower { get; set; }
        public string SubTower { get; set; }
        public string Application { get; set; }
        public string Project { get; set; }
        public DateTime? ExpiryDate { get; set; }
        public string Keyword { get; set; }
        public string Comments { get; set; }

    }
}