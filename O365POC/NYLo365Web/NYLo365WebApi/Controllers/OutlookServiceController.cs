using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Http;
using System.Xml;
using NYLo365WebApi.Models;
using Microsoft.SharePoint.Client;
using System.Security;

namespace NYLo365WebApi.Controllers
{
    public class OutlookServiceController : ApiController
    {
        public OutlookServiceResponse PostAttachments(OutlookServiceRequest request)
        {
            OutlookServiceResponse response = new OutlookServiceResponse();

            try
            {
                response = GetAttachmentsFromExchangeServer(request);
            }
            catch (Exception ex)
            {
                response.IsError = true;
                response.Message = ex.Message;
            }

            return response;
        }

        private OutlookServiceResponse GetAttachmentsFromExchangeServer(OutlookServiceRequest request)
        {
            int processedCount = 0;
            List<string> attachmentNames = new List<string>();

            foreach (ArtifactDetails attachment in request.Attachments)
            {
                // Prepare a web request object.
                HttpWebRequest webRequest = WebRequest.CreateHttp(request.EwsUrl);
                webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", request.AttachmentToken));
                webRequest.PreAuthenticate = true;
                webRequest.AllowAutoRedirect = false;
                webRequest.Method = "POST";
                webRequest.ContentType = "text/xml; charset=utf-8";

                // Construct the SOAP message for the GetAttchment operation.
                byte[] bodyBytes = Encoding.UTF8.GetBytes(string.Format(GetAttachmentSoapRequest, attachment.Id));
                webRequest.ContentLength = bodyBytes.Length;

                Stream requestStream = webRequest.GetRequestStream();
                requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                requestStream.Close();

                // Make the request to the Exchange server and get the response.
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                // If the response is okay, create an XML document from the
                // response and process the request.
                if (webResponse.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = webResponse.GetResponseStream();

                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(responseStream);

                    //Trace.Write(xmlDocument.InnerXml);

                    string content = GetContent(xmlDocument);

                    var fileData = System.Convert.FromBase64String(content);
                    MemoryStream ms = new MemoryStream(fileData);
                    
                    //try
                    //{
                    //    FileStream file = new FileStream("C:\\DEV3\\Attachments\\" + attachment.Name, FileMode.Create, FileAccess.Write);
                    //    ms.WriteTo(file);
                    //    file.Close();
                    //    ms.Close();
                    //}
                    //catch(Exception ex)
                    //{
                    //    processedCount++;
                    //    attachmentNames.Add("Error on File Creation: " + ex.Message);
                    //}
                    //Write file to aSharePoint library
                    try
                    {
                        string siteUrl = "https://nylonline.sharepoint.com/sites/KMP";
                        using (ClientContext spContext = new ClientContext(siteUrl))
                        {
                            Web spWeb = spContext.Web;
                            spContext.Credentials = new SharePointOnlineCredentials("sudesh@NYLonline.onmicrosoft.com", GetSecureString("1qaz2wsx@"));
                            spContext.Load(spWeb);
                            spContext.ExecuteQuery();
                            
                            string title = spWeb.Title;

                            var targetFileUrl = String.Format("{0}/{1}", "/Shared Documents", attachment.Name);
                            ms.Position = 0;

                            var list = spContext.Web.Lists.GetByTitle("Documents");
                            spContext.Load(list.RootFolder);
                            spContext.ExecuteQuery();

                            var fileUrl = Path.Combine(list.RootFolder.ServerRelativeUrl, "1"+attachment.Name);

                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, fileUrl, ms, true);

                            //Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, targetFileUrl, ms, true);
                            spContext.ExecuteQuery();
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        processedCount++;
                        attachmentNames.Add("ERROR: " + ex.Message);
                    }

                    // Close the response stream.
                    responseStream.Close();
                    webResponse.Close();

                    processedCount++;
                    attachmentNames.Add(attachment.Name);
                }

            }
            OutlookServiceResponse response = new OutlookServiceResponse
            {
                
                AttachmentNames = attachmentNames.ToArray(),
                AttachmentsProcessed = processedCount
            };
            response.IsError = false;
            return response;
        }

        private string GetContent(XmlDocument attachmentDataXML)
        {
            XmlDocument document = attachmentDataXML;

            XmlNamespaceManager manager = new XmlNamespaceManager(document.NameTable);

            manager.AddNamespace("t", "http://schemas.microsoft.com/exchange/services/2006/types");
            manager.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages");

            XmlNodeList xnList = document.SelectNodes("//t:Content", manager);
            int nodes = xnList.Count;

            if (nodes > 0)
            {
                return xnList[0].InnerXml;
            }

            return null;
        }

        private SecureString GetSecureString(string password)
        {
            SecureString securePassword = new SecureString();
            foreach (Char c in password.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }

        private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
    }
}