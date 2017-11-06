using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft_Graph_SDK_ASPNET_Connect.Helpers;
using Microsoft_Graph_SDK_ASPNET_Connect.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Controllers
{
    public class FileHandlerController : Controller
    {
        GraphService graphService = new GraphService();
        // GET: FileHandler
        public ActionResult Index()
        {
            return View();
            //return View("Graph");
        }
        [Authorize]
        public async Task<ActionResult> Preview()
        {
            var itemsJson =  Request.Form["items"];
            var itemUrls = JsonConvert.DeserializeObject<string[]>(itemsJson);
            string graphApiUrl = itemUrls[0].ToString();
            Session.Add("graphUrl", graphApiUrl);
            @ViewBag.OneDriveUrl = graphApiUrl;
            return View("Metaform");
            /* try { 

             var itemsJson = Request.Form["items"];
             //var itemsJson = "https://graph.microsoft.com/v1.0/shares/u!aHR0cHM6Ly9ueWxvbmxpbmUtbXkuc2hhcmVwb2ludC5jb20vcGVyc29uYWwvc3VkZXNoX255bG9ubGluZV9vbm1pY3Jvc29mdF9jb20vRG9jdW1lbnRzL0RvY3VtZW50MTAuZG9jeD90ZW1wYXV0aD1leUowZVhBaU9pSktWMVFpTENKaGJHY2lPaUp1YjI1bEluMC5leUpoZFdRaU9pSXdNREF3TURBd015MHdNREF3TFRCbVpqRXRZMlV3TUMwd01EQXdNREF3TURBd01EQXZibmxzYjI1c2FXNWxMVzE1TG5Ob1lYSmxjRzlwYm5RdVkyOXRRR1EyTXpjMllUaGtMVFprWTJVdE5EUTNaUzFoTVdFMUxXSmtaREJoT1RnMllXRmxOeUlzSW1semN5STZJakF3TURBd01EQXpMVEF3TURBdE1HWm1NUzFqWlRBd0xUQXdNREF3TURBd01EQXdNQ0lzSW01aVppSTZJakUxTURnNE1qZzNORGtpTENKbGVIQWlPaUl4TlRBNE9URTFNVFE1SWl3aVpXNWtjRzlwYm5SMWNtd2lPaUl3YXpsc2VUTk9Sa2xQU0dobFYzSjNPV2h6ZDFaMmVXa3dNVTFvV1VRd1VqUXpaVU5xVTBzMldreEZQU0lzSW1WdVpIQnZhVzUwZFhKc1RHVnVaM1JvSWpvaU1UQXpJaXdpYVhOc2IyOXdZbUZqYXlJNklsUnlkV1VpTENKamFXUWlPaUpOUkdSdFRtcEpNVTlYVlhSUFZFRXdUMU13TUUxRVFYZE1WR00wVDFkRmRGcEVRWGxPYWxKc1drZEtiVTlFWXpBaUxDSjJaWElpT2lKb1lYTm9aV1J3Y205dlpuUnZhMlZ1SWl3aWMybDBaV2xrSWpvaVRtMUdhMXBFVW10TmFtTjBUV3BCTUZsNU1EQlplbXQ2VEZSck1VMTZWWFJPVkZFelRtcFpORmxxV21wT2JWazFJaXdpYm1GdFpXbGtJam9pTUNNdVpueHRaVzFpWlhKemFHbHdmSE4xWkdWemFFQnVlV3h2Ym14cGJtVXViMjV0YVdOeWIzTnZablF1WTI5dElpd2libWxwSWpvaWJXbGpjbTl6YjJaMExuTm9ZWEpsY0c5cGJuUWlMQ0pwYzNWelpYSWlPaUowY25WbElpd2lZMkZqYUdWclpYa2lPaUl3YUM1bWZHMWxiV0psY25Ob2FYQjhNVEF3TXpObVptWmhOVFpsWmpsbVpVQnNhWFpsTG1OdmJTSXNJblIwSWpvaU1DSXNJblZ6WlZCbGNuTnBjM1JsYm5SRGIyOXJhV1VpT2lJeUluMC5jVEJJTVdKR01FSklSV1pQTUZFeVRIUmFjM280WTJaR1V6WjNUVVZJZUhKRFFYbE9iMWhzU1VodFZUMA/driveItem";
             var itemUrls = JsonConvert.DeserializeObject<string[]>(itemsJson);
             string graphApiUrl = itemUrls[0].ToString();
             //string graphApiUrl = itemsJson;
             // Initialize the GraphServiceClient.
             //GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

             //Getting tokenPart of the grphapi URL
             string toBeSearched = "shares/";
             int ix = graphApiUrl.IndexOf(toBeSearched);
             string graphToken = "";
             if (ix != -1)
             {
                 graphToken = graphApiUrl.Substring(ix + toBeSearched.Length);

             }
             GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();
             var request = graphClient.Shares[graphToken];
             var foundFile = await request.Request().GetAsync();

             MemoryStream stream = (MemoryStream)await graphClient.Me.Drive.Items[foundFile.Id].Content.Request().GetAsync();

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

                     var targetFileUrl = String.Format("{0}/{1}", "/Shared Documents", foundFile.Name);
                     stream.Position = 0;

                     var list = spContext.Web.Lists.GetByTitle("Documents");
                     spContext.Load(list.RootFolder);
                     spContext.ExecuteQuery();

                     var fileUrl = Path.Combine(list.RootFolder.ServerRelativeUrl, foundFile.Name);

                     Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, fileUrl, stream, true);

                     //Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, targetFileUrl, ms, true);
                     spContext.ExecuteQuery();


                 }
             }
             catch (Exception ex)
             {
                 //processedCount++;
                 //attachmentNames.Add("ERROR: " + ex.Message);
             }
         }
             catch (ServiceException se)
             {

             }

             return View("FiileHandler");*/




        }
        [Authorize]
        public  async Task<FileHandlerResponce> SubmitKM(MetaData requestp)
        {
            FileHandlerResponce fsRes = new FileHandlerResponce();
            string msg = "";
            try
            {
                
                //var itemsJson = "https://graph.microsoft.com/v1.0/shares/u!aHR0cHM6Ly9ueWxvbmxpbmUtbXkuc2hhcmVwb2ludC5jb20vcGVyc29uYWwvc3VkZXNoX255bG9ubGluZV9vbm1pY3Jvc29mdF9jb20vRG9jdW1lbnRzL0RvY3VtZW50MTAuZG9jeD90ZW1wYXV0aD1leUowZVhBaU9pSktWMVFpTENKaGJHY2lPaUp1YjI1bEluMC5leUpoZFdRaU9pSXdNREF3TURBd015MHdNREF3TFRCbVpqRXRZMlV3TUMwd01EQXdNREF3TURBd01EQXZibmxzYjI1c2FXNWxMVzE1TG5Ob1lYSmxjRzlwYm5RdVkyOXRRR1EyTXpjMllUaGtMVFprWTJVdE5EUTNaUzFoTVdFMUxXSmtaREJoT1RnMllXRmxOeUlzSW1semN5STZJakF3TURBd01EQXpMVEF3TURBdE1HWm1NUzFqWlRBd0xUQXdNREF3TURBd01EQXdNQ0lzSW01aVppSTZJakUxTURnNE1qZzNORGtpTENKbGVIQWlPaUl4TlRBNE9URTFNVFE1SWl3aVpXNWtjRzlwYm5SMWNtd2lPaUl3YXpsc2VUTk9Sa2xQU0dobFYzSjNPV2h6ZDFaMmVXa3dNVTFvV1VRd1VqUXpaVU5xVTBzMldreEZQU0lzSW1WdVpIQnZhVzUwZFhKc1RHVnVaM1JvSWpvaU1UQXpJaXdpYVhOc2IyOXdZbUZqYXlJNklsUnlkV1VpTENKamFXUWlPaUpOUkdSdFRtcEpNVTlYVlhSUFZFRXdUMU13TUUxRVFYZE1WR00wVDFkRmRGcEVRWGxPYWxKc1drZEtiVTlFWXpBaUxDSjJaWElpT2lKb1lYTm9aV1J3Y205dlpuUnZhMlZ1SWl3aWMybDBaV2xrSWpvaVRtMUdhMXBFVW10TmFtTjBUV3BCTUZsNU1EQlplbXQ2VEZSck1VMTZWWFJPVkZFelRtcFpORmxxV21wT2JWazFJaXdpYm1GdFpXbGtJam9pTUNNdVpueHRaVzFpWlhKemFHbHdmSE4xWkdWemFFQnVlV3h2Ym14cGJtVXViMjV0YVdOeWIzTnZablF1WTI5dElpd2libWxwSWpvaWJXbGpjbTl6YjJaMExuTm9ZWEpsY0c5cGJuUWlMQ0pwYzNWelpYSWlPaUowY25WbElpd2lZMkZqYUdWclpYa2lPaUl3YUM1bWZHMWxiV0psY25Ob2FYQjhNVEF3TXpObVptWmhOVFpsWmpsbVpVQnNhWFpsTG1OdmJTSXNJblIwSWpvaU1DSXNJblZ6WlZCbGNuTnBjM1JsYm5SRGIyOXJhV1VpT2lJeUluMC5jVEJJTVdKR01FSklSV1pQTUZFeVRIUmFjM280WTJaR1V6WjNUVVZJZUhKRFFYbE9iMWhzU1VodFZUMA/driveItem";
                //var itemUrls = Session["graphUrl"].ToString();
                string graphApiUrl = Session["graphUrl"].ToString();
                //string graphApiUrl = itemsJson;
                // Initialize the GraphServiceClient.
                //GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                //Getting tokenPart of the grphapi URL
                string toBeSearched = "shares/";
                int ix = graphApiUrl.IndexOf(toBeSearched);
                string graphToken = "";
                if (ix != -1)
                {
                    graphToken = graphApiUrl.Substring(ix + toBeSearched.Length);

                }
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();
                var request = graphClient.Shares[graphToken];
                var foundFile = await request.Request().GetAsync();

                MemoryStream stream = (MemoryStream)await graphClient.Me.Drive.Items[foundFile.Id].Content.Request().GetAsync();

                try
                {
                    string siteUrl = "https://nylonline.sharepoint.com/sites/IBM";
                    using (ClientContext spContext = new ClientContext(siteUrl))
                    {
                        Web spWeb = spContext.Web;
                        spContext.Credentials = new SharePointOnlineCredentials("hassan@NYLonline.onmicrosoft.com", GetSecureString("kmp@2017"));
                        spContext.Load(spWeb);
                        spContext.ExecuteQuery();

                        string title = spWeb.Title;

                        var targetFileUrl = String.Format("{0}/{1}", "/IT Business Management Documents", foundFile.Name);
                        stream.Position = 0;

                        var list = spContext.Web.Lists.GetByTitle("IT Business Management Documents");
                        spContext.Load(list.RootFolder);
                        spContext.ExecuteQuery();

                        var fileUrl = Path.Combine(list.RootFolder.ServerRelativeUrl, foundFile.Name);
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, fileUrl, stream, true);
                        spContext.ExecuteQuery();

                        //Updating MetaData

                        Microsoft.SharePoint.Client.File newFile = spContext.Web.GetFileByServerRelativeUrl(fileUrl);
                        ListItem item = newFile.ListItemAllFields;

                        ArtifactDetails attachment = new ArtifactDetails();
                        attachment.Function = requestp.Functiont;
                        attachment.DocumentType = requestp.DocumentType;
                        attachment.LineOfBusiness = requestp.LineOfBusiness;
                        attachment.BusinessArea = requestp.BusinessArea;
                        attachment.SubBusinessArea = requestp.SubBusinessArea;
                        attachment.SubFunction = requestp.SubFunction;
                        attachment.Tower = requestp.Tower;
                        attachment.SubTower = requestp.SubTower;
                        attachment.Application = requestp.Application;
                        attachment.Project = requestp.Project;
                        if(requestp.ExpiryDate == string.Empty)
                        {
                            attachment.ExpiryDate = null;
                        }
                        else
                        {
                            attachment.ExpiryDate = Convert.ToDateTime(requestp.ExpiryDate);
                        }                        
                        attachment.Keyword = requestp.Keyword;
                        attachment.Comments = requestp.Comments;

                        UpdateTaxonomyFields(item, attachment);


                        fsRes.IsError = false;
                        fsRes.Message = "Successfully Uploaded!";
                    }
                }
                catch (Exception ex)
                {
                    fsRes.IsError = true;
                    fsRes.Message = ex.Message ;
                    msg = ex.Message;
                }
            }
            catch (ServiceException se)
            {
                fsRes.IsError = true;
                fsRes.Message = se.Message;
                msg = se.Message;
            }
            msg = "OK";
            
            return fsRes;
        }

        private void UpdateTaxonomyFields(ListItem item, ArtifactDetails attachment)
        {
            var ctx = item.Context;
            var list = item.ParentList;

            var fldFunction = list.Fields.GetByInternalNameOrTitle("MMFunction");
            var taxFieldFunction = ctx.CastTo<TaxonomyField>(fldFunction);
            TaxonomyFieldValue termValueFunction = new TaxonomyFieldValue();
            termValueFunction.Label = attachment.Function.Split('|')[0];
            termValueFunction.TermGuid = attachment.Function.Split('|')[1];
            termValueFunction.WssId = -1;
            taxFieldFunction.SetFieldValueByValue(item, termValueFunction);

            var fldDocumentType = list.Fields.GetByInternalNameOrTitle("MMDocumentType");
            var taxFieldDocumentType = ctx.CastTo<TaxonomyField>(fldDocumentType);
            TaxonomyFieldValue termValueDocumentType = new TaxonomyFieldValue();
            termValueDocumentType.Label = attachment.DocumentType.Split('|')[0];
            termValueDocumentType.TermGuid = attachment.DocumentType.Split('|')[1];
            termValueDocumentType.WssId = -1;
            taxFieldDocumentType.SetFieldValueByValue(item, termValueDocumentType);

            var fldLineofBusiness = list.Fields.GetByInternalNameOrTitle("MMLineofBusiness");
            var taxFieldLineofBusiness = ctx.CastTo<TaxonomyField>(fldLineofBusiness);
            TaxonomyFieldValue termValueLineofBusiness = new TaxonomyFieldValue();
            termValueLineofBusiness.Label = attachment.LineOfBusiness.Split('|')[0];
            termValueLineofBusiness.TermGuid = attachment.LineOfBusiness.Split('|')[1];
            termValueLineofBusiness.WssId = -1;
            taxFieldLineofBusiness.SetFieldValueByValue(item, termValueLineofBusiness);

            var fldBusinessArea = list.Fields.GetByInternalNameOrTitle("MMBusinessArea");
            var taxFieldBusinessArea = ctx.CastTo<TaxonomyField>(fldBusinessArea);
            TaxonomyFieldValue termValueBusinessArea = new TaxonomyFieldValue();
            termValueBusinessArea.Label = attachment.BusinessArea.Split('|')[0];
            termValueBusinessArea.TermGuid = attachment.BusinessArea.Split('|')[1];
            termValueBusinessArea.WssId = -1;
            taxFieldBusinessArea.SetFieldValueByValue(item, termValueBusinessArea);

            var fldSubBusinessArea = list.Fields.GetByInternalNameOrTitle("MMSubBusinessArea");
            var taxFieldSubBusinessArea = ctx.CastTo<TaxonomyField>(fldSubBusinessArea);
            TaxonomyFieldValue termValueSubBusinessArea = new TaxonomyFieldValue();
            termValueSubBusinessArea.Label = attachment.SubBusinessArea.Split('|')[0];
            termValueSubBusinessArea.TermGuid = attachment.SubBusinessArea.Split('|')[1];
            termValueSubBusinessArea.WssId = -1;
            taxFieldSubBusinessArea.SetFieldValueByValue(item, termValueSubBusinessArea);

            var fldSubFunction = list.Fields.GetByInternalNameOrTitle("MMSubFunction");
            var taxFieldSubFunction = ctx.CastTo<TaxonomyField>(fldSubFunction);
            TaxonomyFieldValue termValueSubFunction = new TaxonomyFieldValue();
            termValueSubFunction.Label = attachment.SubFunction.Split('|')[0];
            termValueSubFunction.TermGuid = attachment.SubFunction.Split('|')[1];
            termValueSubFunction.WssId = -1;
            taxFieldSubFunction.SetFieldValueByValue(item, termValueSubFunction);

            item["Base_x0020_Content"] = "Document";
            item["Organization"] = "Insurance Technology";
            item["Project_x0020_ID"] = attachment.Project;

            item.Update();
            ctx.Load(item);
            ctx.ExecuteQuery();
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