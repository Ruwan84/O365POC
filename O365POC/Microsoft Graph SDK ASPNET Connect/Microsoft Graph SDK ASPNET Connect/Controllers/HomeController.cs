/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.Graph;
using Microsoft_Graph_SDK_ASPNET_Connect.Helpers;
using Microsoft_Graph_SDK_ASPNET_Connect.Models;
using Resources;
using System.Net.Http;
using System;
using System.IO;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Security;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Controllers
{
    public class HomeController : Controller
    {
        GraphService graphService = new GraphService();

        public ActionResult Index()
        {
            return View("Graph");
        }

        [Authorize]
        // Get the current user's email address from their profile.
        public async Task<ActionResult> GetMyEmailAddress()
        {
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                //string url = "https://graph.microsoft.com/v1.0/shares/u!aHR0cHM6Ly9ueWxvbmxpbmUtbXkuc2hhcmVwb2ludC5jb20vcGVyc29uYWwvc3VkZXNoX255bG9ubGluZV9vbm1pY3Jvc29mdF9jb20vRG9jdW1lbnRzL0RvY3VtZW50MTAuZG9jeD90ZW1wYXV0aD1leUowZVhBaU9pSktWMVFpTENKaGJHY2lPaUp1YjI1bEluMC5leUpoZFdRaU9pSXdNREF3TURBd015MHdNREF3TFRCbVpqRXRZMlV3TUMwd01EQXdNREF3TURBd01EQXZibmxzYjI1c2FXNWxMVzE1TG5Ob1lYSmxjRzlwYm5RdVkyOXRRR1EyTXpjMllUaGtMVFprWTJVdE5EUTNaUzFoTVdFMUxXSmtaREJoT1RnMllXRmxOeUlzSW1semN5STZJakF3TURBd01EQXpMVEF3TURBdE1HWm1NUzFqWlRBd0xUQXdNREF3TURBd01EQXdNQ0lzSW01aVppSTZJakUxTURnNE1qZzNORGtpTENKbGVIQWlPaUl4TlRBNE9URTFNVFE1SWl3aVpXNWtjRzlwYm5SMWNtd2lPaUl3YXpsc2VUTk9Sa2xQU0dobFYzSjNPV2h6ZDFaMmVXa3dNVTFvV1VRd1VqUXpaVU5xVTBzMldreEZQU0lzSW1WdVpIQnZhVzUwZFhKc1RHVnVaM1JvSWpvaU1UQXpJaXdpYVhOc2IyOXdZbUZqYXlJNklsUnlkV1VpTENKamFXUWlPaUpOUkdSdFRtcEpNVTlYVlhSUFZFRXdUMU13TUUxRVFYZE1WR00wVDFkRmRGcEVRWGxPYWxKc1drZEtiVTlFWXpBaUxDSjJaWElpT2lKb1lYTm9aV1J3Y205dlpuUnZhMlZ1SWl3aWMybDBaV2xrSWpvaVRtMUdhMXBFVW10TmFtTjBUV3BCTUZsNU1EQlplbXQ2VEZSck1VMTZWWFJPVkZFelRtcFpORmxxV21wT2JWazFJaXdpYm1GdFpXbGtJam9pTUNNdVpueHRaVzFpWlhKemFHbHdmSE4xWkdWemFFQnVlV3h2Ym14cGJtVXViMjV0YVdOeWIzTnZablF1WTI5dElpd2libWxwSWpvaWJXbGpjbTl6YjJaMExuTm9ZWEpsY0c5cGJuUWlMQ0pwYzNWelpYSWlPaUowY25WbElpd2lZMkZqYUdWclpYa2lPaUl3YUM1bWZHMWxiV0psY25Ob2FYQjhNVEF3TXpObVptWmhOVFpsWmpsbVpVQnNhWFpsTG1OdmJTSXNJblIwSWpvaU1DSXNJblZ6WlZCbGNuTnBjM1JsYm5SRGIyOXJhV1VpT2lJeUluMC5jVEJJTVdKR01FSklSV1pQTUZFeVRIUmFjM280WTJaR1V6WjNUVVZJZUhKRFFYbE9iMWhzU1VodFZUMA/driveItem";
                string url = "u!aHR0cHM6Ly9ueWxvbmxpbmUtbXkuc2hhcmVwb2ludC5jb20vcGVyc29uYWwvc3VkZXNoX255bG9ubGluZV9vbm1pY3Jvc29mdF9jb20vRG9jdW1lbnRzL0RvY3VtZW50MTAuZG9jeD90ZW1wYXV0aD1leUowZVhBaU9pSktWMVFpTENKaGJHY2lPaUp1YjI1bEluMC5leUpoZFdRaU9pSXdNREF3TURBd015MHdNREF3TFRCbVpqRXRZMlV3TUMwd01EQXdNREF3TURBd01EQXZibmxzYjI1c2FXNWxMVzE1TG5Ob1lYSmxjRzlwYm5RdVkyOXRRR1EyTXpjMllUaGtMVFprWTJVdE5EUTNaUzFoTVdFMUxXSmtaREJoT1RnMllXRmxOeUlzSW1semN5STZJakF3TURBd01EQXpMVEF3TURBdE1HWm1NUzFqWlRBd0xUQXdNREF3TURBd01EQXdNQ0lzSW01aVppSTZJakUxTURnNE1qZzNORGtpTENKbGVIQWlPaUl4TlRBNE9URTFNVFE1SWl3aVpXNWtjRzlwYm5SMWNtd2lPaUl3YXpsc2VUTk9Sa2xQU0dobFYzSjNPV2h6ZDFaMmVXa3dNVTFvV1VRd1VqUXpaVU5xVTBzMldreEZQU0lzSW1WdVpIQnZhVzUwZFhKc1RHVnVaM1JvSWpvaU1UQXpJaXdpYVhOc2IyOXdZbUZqYXlJNklsUnlkV1VpTENKamFXUWlPaUpOUkdSdFRtcEpNVTlYVlhSUFZFRXdUMU13TUUxRVFYZE1WR00wVDFkRmRGcEVRWGxPYWxKc1drZEtiVTlFWXpBaUxDSjJaWElpT2lKb1lYTm9aV1J3Y205dlpuUnZhMlZ1SWl3aWMybDBaV2xrSWpvaVRtMUdhMXBFVW10TmFtTjBUV3BCTUZsNU1EQlplbXQ2VEZSck1VMTZWWFJPVkZFelRtcFpORmxxV21wT2JWazFJaXdpYm1GdFpXbGtJam9pTUNNdVpueHRaVzFpWlhKemFHbHdmSE4xWkdWemFFQnVlV3h2Ym14cGJtVXViMjV0YVdOeWIzTnZablF1WTI5dElpd2libWxwSWpvaWJXbGpjbTl6YjJaMExuTm9ZWEpsY0c5cGJuUWlMQ0pwYzNWelpYSWlPaUowY25WbElpd2lZMkZqYUdWclpYa2lPaUl3YUM1bWZHMWxiV0psY25Ob2FYQjhNVEF3TXpObVptWmhOVFpsWmpsbVpVQnNhWFpsTG1OdmJTSXNJblIwSWpvaU1DSXNJblZ6WlZCbGNuTnBjM1JsYm5SRGIyOXJhV1VpT2lJeUluMC5jVEJJTVdKR01FSklSV1pQTUZFeVRIUmFjM280WTJaR1V6WjNUVVZJZUhKRFFYbE9iMWhzU1VodFZUMA/driveItem";



                var request = graphClient.Shares[url];
                //var request = graphClient.Shares[url].Root;
                var foundFile = await request.Request().GetAsync();

                // var stream = await graphClient.Drive.Items[foundFile.AdditionalData[]]
                //MemoryStream stream = (MemoryStream)await graphClient.Me.Drive.Items["01ZZ4ET4K6GAHSZAUZYFEICYWLLL3Y6PH4"].Content.Request().GetAsync();
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

                






                using (var client = new WebClient())
                {
                    //var result = client.OpenRead(new Uri(foundFile.AdditionalData["@microsoft.graph.downloadUrl"].ToString()));
                    //FileStream  sr = new StreamReader(result);
                    //jsonData = sr.ReadToEnd();
                    FileStream fileStream = System.IO.File.OpenRead(@"https://graph.microsoft.com:443/v1.0/drives/b!J03dakwgk0yVNVR2aLbG-Q7vzjGftsRCh1ZDbPgZFG8MyUd2a8IeQbGjYsCorbAy/root:");
                    //var fileStream = client.OpenRead(new Uri(foundFile.AdditionalData["@microsoft.graph.downloadUrl"].ToString()));
                    byte[] fileData = new byte[fileStream.Length];
                    fileStream.Read(fileData, 0, fileData.Length);
                    fileStream.Close();
                    //return fileData;
                }

                //OneDriveInfo driveInfo = JsonConvert.DeserializeObject<OneDriveInfo>(jsonData);


                //var bytes = File.ReadAllBytes(foundFile.AdditionalData["@microsoft.graph.downloadUrl"].ToString());
                Stream currentUserPhotoStream = null;

                //DriveItem uploadedFile = null;

                try
                {
                    MemoryStream fileStream = new MemoryStream();
                    currentUserPhotoStream = await graphClient.Me.Drive.Root.ItemWithPath(foundFile.Name).Content.Request().GetAsync();

                }


                catch (ServiceException)
                {
                    return null;
                }

               

                /*var response = await graphClient.HttpProvider.SendAsync(foundFile);

                var stream = await response.Content.ReadAsStreamAsync();*/
                //GraphServiceClient graphClient1 = SDKHelper.GetAuthenticatedClient();
                //var request1 = graphClient.Shares[url].Items[foundFile.Id].Content;
                var request1 = graphClient.Shares[url].Items[foundFile.Id].Content.Request().GetHttpRequestMessage(); ;
                request1.RequestUri = new Uri(request1.RequestUri.AbsoluteUri.Replace("graph.microsoft.com", "api.onedrive.com"));
                var response = await graphClient.HttpProvider.SendAsync(request1);

                //var stream = await response.Content.ReadAsStreamAsync();
                //var content = await request1.Request().GetAsync();



                //var foundFile = await graphClient.Shares[UrlToSharingToken(sharingLink)].Root.Children.Request().GetAsync();

                //var request = graphClient.Shares[UrlToSharingToken(sharingLink)].Items[foundFile[0].Id].Content.Request().GetHttpRequestMessage();

                //request.RequestUri = new Uri(request.RequestUri.AbsoluteUri.Replace("graph.microsoft.com", "api.onedrive.com"));

                //var response = await graphClient.HttpProvider.SendAsync(request);

                //var stream = await response.Content.ReadAsStreamAsync();


                // Get the current user's email address. 

                ViewBag.Email = await graphService.GetMyEmailAddress(graphClient);
                return View("Graph");
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + se.Error.Message });
            }
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
        [Authorize]
        // Send mail on behalf of the current user.
        public async Task<ActionResult> SendEmail()
        {
            if (string.IsNullOrEmpty(Request.Form["email-address"]))
            {
                ViewBag.Message = Resource.Graph_SendMail_Message_GetEmailFirst;
                return View("Graph");
            }

            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Build the email message.
                Message message = await graphService.BuildEmailMessage(graphClient, Request.Form["recipients"], Request.Form["subject"]);
                
                // Send the email.
                await graphService.SendEmail(graphClient, message);

                // Reset the current user's email address and the status to display when the page reloads.
                ViewBag.Email = Request.Form["email-address"];
                ViewBag.Message = Resource.Graph_SendMail_Success_Result;
                return View("Graph");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + se.Error.Message });
            }
        }

        public ActionResult About()
        {
            return View();
        }
    }
}