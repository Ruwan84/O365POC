using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NYLo365WebApi.Models
{
    public class OutlookServiceResponse
    {
        public bool IsError { get; set; }
        public string Message { get; set; }
        public int AttachmentsProcessed { get; set; }
        public string[] AttachmentNames { get; set; }
    }
}