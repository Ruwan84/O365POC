/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Graph;
using Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Models
{
    public class GraphService
    {

        // Get the current user's email address from their profile.
        public async Task<string> GetMyEmailAddress(GraphServiceClient graphClient)
        {
            //IEnumerable<Microsoft.Graph.Option> urlId = '';
            string url = "https://graph.microsoft.com/v1.0/shares/u!aHR0cHM6Ly9ueWxvbmxpbmUtbXkuc2hhcmVwb2ludC5jb20vcGVyc29uYWwvc3VkZXNoX255bG9ubGluZV9vbm1pY3Jvc29mdF9jb20vRG9jdW1lbnRzL0RvY3VtZW50MTAuZG9jeD90ZW1wYXV0aD1leUowZVhBaU9pSktWMVFpTENKaGJHY2lPaUp1YjI1bEluMC5leUpoZFdRaU9pSXdNREF3TURBd015MHdNREF3TFRCbVpqRXRZMlV3TUMwd01EQXdNREF3TURBd01EQXZibmxzYjI1c2FXNWxMVzE1TG5Ob1lYSmxjRzlwYm5RdVkyOXRRR1EyTXpjMllUaGtMVFprWTJVdE5EUTNaUzFoTVdFMUxXSmtaREJoT1RnMllXRmxOeUlzSW1semN5STZJakF3TURBd01EQXpMVEF3TURBdE1HWm1NUzFqWlRBd0xUQXdNREF3TURBd01EQXdNQ0lzSW01aVppSTZJakUxTURnNE1qZzNORGtpTENKbGVIQWlPaUl4TlRBNE9URTFNVFE1SWl3aVpXNWtjRzlwYm5SMWNtd2lPaUl3YXpsc2VUTk9Sa2xQU0dobFYzSjNPV2h6ZDFaMmVXa3dNVTFvV1VRd1VqUXpaVU5xVTBzMldreEZQU0lzSW1WdVpIQnZhVzUwZFhKc1RHVnVaM1JvSWpvaU1UQXpJaXdpYVhOc2IyOXdZbUZqYXlJNklsUnlkV1VpTENKamFXUWlPaUpOUkdSdFRtcEpNVTlYVlhSUFZFRXdUMU13TUUxRVFYZE1WR00wVDFkRmRGcEVRWGxPYWxKc1drZEtiVTlFWXpBaUxDSjJaWElpT2lKb1lYTm9aV1J3Y205dlpuUnZhMlZ1SWl3aWMybDBaV2xrSWpvaVRtMUdhMXBFVW10TmFtTjBUV3BCTUZsNU1EQlplbXQ2VEZSck1VMTZWWFJPVkZFelRtcFpORmxxV21wT2JWazFJaXdpYm1GdFpXbGtJam9pTUNNdVpueHRaVzFpWlhKemFHbHdmSE4xWkdWemFFQnVlV3h2Ym14cGJtVXViMjV0YVdOeWIzTnZablF1WTI5dElpd2libWxwSWpvaWJXbGpjbTl6YjJaMExuTm9ZWEpsY0c5cGJuUWlMQ0pwYzNWelpYSWlPaUowY25WbElpd2lZMkZqYUdWclpYa2lPaUl3YUM1bWZHMWxiV0psY25Ob2FYQjhNVEF3TXpObVptWmhOVFpsWmpsbVpVQnNhWFpsTG1OdmJTSXNJblIwSWpvaU1DSXNJblZ6WlZCbGNuTnBjM1JsYm5SRGIyOXJhV1VpT2lJeUluMC5jVEJJTVdKR01FSklSV1pQTUZFeVRIUmFjM280WTJaR1V6WjNUVVZJZUhKRFFYbE9iMWhzU1VodFZUMA/driveItem";
                      
            // Get the current user. 
            // This sample only needs the user's email address, so select the mail and userPrincipalName properties.
            // If the mail property isn't defined, userPrincipalName should map to the email for all account types. 
            User me = await graphClient.Me.Request().Select("mail,userPrincipalName").GetAsync();
            var res = await graphClient.Shares.Request().Select(url).GetAsync();
                

            return me.Mail ?? me.UserPrincipalName;
        }

        // Send an email message from the current user.
        public async Task SendEmail(GraphServiceClient graphClient, Message message)
        {
            await graphClient.Me.SendMail(message, true).Request().PostAsync();
        }

        // Create the email message.
        public async Task<Message> BuildEmailMessage(GraphServiceClient graphClient, string recipients, string subject)
        {

            // Get current user photo
            Stream photoStream = await GetCurrentUserPhotoStreamAsync(graphClient);


            // If the user doesn't have a photo, or if the user account is MSA, we use a default photo

            if ( photoStream == null)
            {
                photoStream = System.IO.File.OpenRead(System.Web.Hosting.HostingEnvironment.MapPath("/Content/test.jpg"));
            }

            MemoryStream photoStreamMS = new MemoryStream();
            // Copy stream to MemoryStream object so that it can be converted to byte array.
            photoStream.CopyTo(photoStreamMS);

            DriveItem photoFile = await UploadFileToOneDrive(graphClient, photoStreamMS.ToArray());

            MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
            attachments.Add(new FileAttachment
            {
                ODataType = "#microsoft.graph.fileAttachment",
                ContentBytes = photoStreamMS.ToArray(),
                ContentType = "image/png",
                Name = "me.png"
            });

            Permission sharingLink = await GetSharingLinkAsync(graphClient, photoFile.Id);

            // Add the sharing link to the email body.
            string bodyContent = string.Format(Resource.Graph_SendMail_Body_Content, sharingLink.Link.WebUrl);

            // Prepare the recipient list.
            string[] splitter = { ";" };
            string[] splitRecipientsString = recipients.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            List<Recipient> recipientList = new List<Recipient>();
            foreach (string recipient in splitRecipientsString)
            {
                recipientList.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient.Trim()
                    }
                });
            }

            // Build the email message.
            Message email = new Message
            {
                Body = new ItemBody
                {
                    Content = bodyContent,
                    ContentType = BodyType.Html,
                },
                Subject = subject,
                ToRecipients = recipientList,
                Attachments = attachments
            };
            return email;
        }

        // Gets the stream content of the signed-in user's photo. 
        // This snippet doesn't work with consumer accounts.
        public async Task<Stream> GetCurrentUserPhotoStreamAsync(GraphServiceClient graphClient)
        {
            Stream currentUserPhotoStream = null;

            try
            {
                currentUserPhotoStream = await graphClient.Me.Photo.Content.Request().GetAsync();

            }

            // If the user account is MSA (not work or school), the service will throw an exception.
            catch (ServiceException)
            {
                return null;
            }

            return currentUserPhotoStream;

        }

        // Uploads the specified file to the user's root OneDrive directory.
        public async Task<DriveItem> UploadFileToOneDrive(GraphServiceClient graphClient, byte[] file)
        {
            DriveItem uploadedFile = null;

            try
            {
                MemoryStream fileStream = new MemoryStream(file);
                uploadedFile = await graphClient.Me.Drive.Root.ItemWithPath("me.png").Content.Request().PutAsync<DriveItem>(fileStream);

            }


            catch (ServiceException)
            {
                return null;
            }

            return uploadedFile;
        }

        public static async Task<Permission> GetSharingLinkAsync(GraphServiceClient graphClient, string Id)
        {
            Permission permission = null;

            try
            {
                permission = await graphClient.Me.Drive.Items[Id].CreateLink("view").Request().PostAsync();
            }

            catch (ServiceException)
            {
                return null;
            }

            return permission;
        }

    }
}