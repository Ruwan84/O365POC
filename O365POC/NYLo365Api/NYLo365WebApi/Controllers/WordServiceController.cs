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
using Microsoft.SharePoint.Client.Taxonomy;

namespace NYLo365WebApi.Controllers
{
    public class WordServiceController : ApiController
    {
        public WordServiceResponse PostDocuments(WordServiceRequestNew request1)
        {                                          
            WordServiceResponse response = new WordServiceResponse();

            try
            {
                response = SubmitFile(request1);
            }
            catch (Exception ex)
            {
                response.IsError = true;
                response.Message = ex.Message;
            }

            return response;
        }

        private WordServiceResponse SubmitFile(WordServiceRequestNew request)
        {
            WordServiceResponse response = new WordServiceResponse();
            var fileData = System.Convert.FromBase64String(request.Content);
            MemoryStream ms = new MemoryStream(fileData);

            //FileStream file = new FileStream("C:\\DEV3\\Attachments\\" + request.attachments[0].name, FileMode.Create, FileAccess.Write);
            //ms.WriteTo(file);
            //file.Close();
            //ms.Close();

            //string messsage = "";

            try
            {
                string siteUrl = "https://nylonline.sharepoint.com/sites/ibm";
                using (ClientContext spContext = new ClientContext(siteUrl))
                {
                    //messsage += "CP1 <br />";
                    Web spWeb = spContext.Web;
                    spContext.Credentials = new SharePointOnlineCredentials("hassan@NYLonline.onmicrosoft.com", GetSecureString("kmp@2017"));
                    spContext.Load(spWeb);
                    spContext.ExecuteQuery();
                    //messsage += "CP2 <br />";
                    string title = spWeb.Title;
                    //messsage += "CP3 <br />";
                    //messsage += "CP4 " + title + " <br />";

                    var targetFileUrl = String.Format("{0}/{1}", "/IT Business Management Documents", request.Name);
                    ms.Position = 0;
                    //messsage += "CP5 <br />";
                    var list = spContext.Web.Lists.GetByTitle("IT Business Management Documents");
                    spContext.Load(list.RootFolder);
                    spContext.ExecuteQuery();
                    //messsage += "CP6 <br />";
                    var fileUrl = Path.Combine(list.RootFolder.ServerRelativeUrl, request.Name);
                    //messsage += "CP7 <br />";
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, fileUrl, ms, true);
                    //messsage += "CP8 <br />";
                    //Microsoft.SharePoint.Client.File.SaveBinaryDirect(spContext, targetFileUrl, ms, true);
                    spContext.ExecuteQuery();


                    Microsoft.SharePoint.Client.File newFile = spContext.Web.GetFileByServerRelativeUrl(fileUrl);
                    ListItem item = newFile.ListItemAllFields;

                    ArtifactDetails attachment = new ArtifactDetails();
                    attachment.Function = request.Function;
                    attachment.DocumentType = request.DocumentType;
                    attachment.LineOfBusiness = request.LineOfBusiness;
                    attachment.BusinessArea = request.BusinessArea;
                    attachment.SubBusinessArea = request.SubBusinessArea;
                    attachment.SubFunction = request.SubFunction;
                    attachment.Tower = request.Tower;
                    attachment.SubTower = request.SubTower;
                    attachment.Application = request.Application;
                    attachment.Project = request.Project;
                    attachment.ExpiryDate = request.ExpiryDate;
                    attachment.Keyword = request.Keyword;
                    attachment.Comments = request.Comments;

                    UpdateTaxonomyFields(item, attachment);
                 
                    response.IsError = false;
                    response.Message = request.Name + " successfully uploaded to the KM Portal.";                     
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

    }
}