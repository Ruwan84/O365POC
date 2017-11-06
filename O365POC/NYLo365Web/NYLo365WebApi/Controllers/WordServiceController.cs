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
    public class WordServiceController : ApiController
    {
        public WordServiceResponse PostAttachments(WordServiceRequest request)
        {
            WordServiceResponse response = new WordServiceResponse();

            try
            {
                response = SubmitFile(request);
            }
            catch (Exception ex)
            {
                response.IsError = true;
                response.Message = ex.Message;
            }

            return response;
        }
        private WordServiceResponse SubmitFile(WordServiceRequest request)
        {
            WordServiceResponse response = new WordServiceResponse();
            var fileData = System.Convert.FromBase64String(request.Content);
            MemoryStream ms = new MemoryStream(fileData);

            //FileStream file = new FileStream("C:\\DEV3\\Attachments\\" + request.attachments[0].name, FileMode.Create, FileAccess.Write);
            //ms.WriteTo(file);
            //file.Close();
            //ms.Close();

            string messsage = "";

            try
            {
                string siteUrl = "https://nylonline.sharepoint.com/sites/KMP";
                using (ClientContext spContext = new ClientContext(siteUrl))
                {
                    messsage += "CP1 <br />";
                    Web spWeb = spContext.Web;
                    spContext.Credentials = new SharePointOnlineCredentials("sudesh@NYLonline.onmicrosoft.com", GetSecureString("1qaz2wsx@"));
                    spContext.Load(spWeb);
                    spContext.ExecuteQuery();
                    messsage += "CP2 <br />";
                    string title = spWeb.Title;
                    messsage += "CP3 <br />";
                    messsage += "CP4 " + title + " <br />";

                    var targetFileUrl = String.Format("{0}/{1}", "/Shared Documents", request.Name);
                    ms.Position = 0;
                    messsage += "CP5 <br />";
                    var list = spContext.Web.Lists.GetByTitle("Documents");
                    spContext.Load(list.RootFolder);
                    spContext.ExecuteQuery();
                    messsage += "CP6 <br />";
                    var fileUrl = Path.Combine(list.RootFolder.ServerRelativeUrl, request.Name);
                    messsage += "CP7 <br />";
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, fileUrl, ms, true);
                    messsage += "CP8 <br />";
                    //Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, targetFileUrl, ms, true);
                    spContext.ExecuteQuery();
                    messsage += "CP9 <br />";
                    response.IsError = false;
                    response.Message = request.Name + "successfully submited to the KMP.";
                    messsage += "CP10 <br />";
                }
            }
            catch (Exception ex)
            {
                response.IsError = true;
                response.Message = "ERROR: " + ex.Message;
            }


            return response;
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

    }
}